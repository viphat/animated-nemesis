require 'spreadsheet'
require 'roo'
require 'roo-xls'
require 'json'

class CodeListToolsService < BaseService

  def paste_codelist(params,json_file,codelist_file)
    ap __method__
    begin
      ap params['codelist']
      codelist = build_codelist_array(params['codelist'])
      json_data = build_from_json_file(json_file)
      options = json_data['options']

      if params[:dual_languages].present? && params[:dual_languages].to_sym == :on
        options['dual_languages'] = true
      else
        options['dual_languages'] = false
      end

      codelist = read_codelist_file(codelist_file,params['sheet'].to_i,codelist,options["dual_languages"])

      options['export_data_type'] = options['export_data_type'].to_sym
      options['export_first'] = options['export_first'].to_sym

      indexes = build_index_from_json(json_data['indexes'])
      data = build_raw_data_from_json(json_data['data'])

      if params['output_file_name'].present? && options['output_file_name'].present?
        if params['output_file_name'] != options['output_file_name']
          options['output_file_name'] = params['output_file_name']
        end
      end

      data = parse_data(data,codelist,options)
    rescue Exception => e
      ap e
      raise e
    end

    # Xuat Data
    begin
      data_tools = DataToolsService.new
      result_file = data_tools.read_and_export_data(json_file,options,params,true,data,indexes)
    rescue Exception => e
      ap e
      result_file = data_tools.log_file
    end
    result_file
  end

  private

  def parse_data(data,codelist,options=nil)
    ap __method__
    codelist.each do |item|
      if item.code.length == 0
        next
      end
      data.each do |datum|
        if datum.question != item.question
          next
        end
        if item.filters != "" && !(datum.filters.include?(item.filters))
          next
        end
        datum = insert_label(datum,item,options)
        datum = insert_into(datum)
        datum = insert_into(datum) if options["dual_languages"] == true

        datum.codelist = true
      end
    end
    data
  end

  def integer?(str)
    /\A[+-]?\d+\z/ === str
  end

  def insert_label(data,codelist,options=nil)
    ap __method__
    codes = []
    codelist.code.each do |c|
      codes << c.code
    end
    p data.sheet_name
    data.val.each do |item|
      if integer?(item['count'][0])
        index = codes.index(item['count'][0].to_i)
        unless index.nil?
          label = codelist.code[index].label
          item['count'].insert(1,label)
          item['percent'].insert(1,label)
          if codelist.code[index].sub_label.present?
            sub_label = codelist.code[index].sub_label
            item['count'].insert(2,sub_label)
            item['percent'].insert(2,sub_label)
          end
          if codelist.code[index].bold == true
            item['bold'] = true
          end
          next
        end
      end
      item['count'].insert(1,"")
      item['percent'].insert(1,"")
      if options["dual_languages"] == true
        item['count'].insert(1,"")
        item['percent'].insert(1,"")
      end
    end
    data_codes = []
    data.val.each do |x|
      data_codes << x['percent'][0].to_i
    end
    c_index = 0
    codelist.code.each do |c|
      index = data_codes.index(c.code)
      unless index.nil?
        data.val.insert(c_index,data.val.delete_at(index))
        data_codes.insert(c_index,data_codes.delete_at(index))
        c_index += 1
      end
    end
    data
  end

  def insert_into(data)
    ap __method__
    data.header.insert(1,"") if data.header.present?
    data.resp.insert(1,"") if data.resp.present?
    data.wtd_resp.insert(1,"") if data.wtd_resp.present?
    data.means.insert(1,"") if data.means.present?
    data.mode.insert(1,"") if data.mode.present?
    data.medians.insert(1,"") if data.medians.present?
    data.std_deviation.insert(1,"") if data.std_deviation.present?
    data.totals_percent.insert(1,"") if data.totals_percent.present?
    data.totals_count.insert(1,"") if data.totals_count.present?
    data
  end

  def build_from_json_file(json_file)
    ap __method__
    file = File.open(json_file,"rb")
    contents = file.read
    # Parse Json Data to Object
    json_data = JSON.parse(contents)
    json_data
  end

  def build_index_from_json(indexes_hash)
    ap __method__
    indexes = []
    indexes_hash.each do |item|
      index = Index.new
      index.build!(item)
      indexes << index
    end
    indexes
  end

  def build_raw_data_from_json(draft)
    ap __method__
    data = []
    draft.each do |item|
      raw_data = RawData.new
      raw_data.build!(item)
      data << raw_data
    end
    data
  end

  def read_codelist_file(codelist_file,sheet,codelist,is_dual_languages)
    ap __method__
    # Spreadsheet.client_encoding = 'UTF-8'
    excel = Roo::Spreadsheet.open(codelist_file)
    sheet = excel.sheet(sheet-1)
    codelist.each_with_index do |item, index|
      next if item.qbegin == 0 || item.question == 'New Row'
      ap "#{item.qbegin}----#{item.qend}"
      codelist.each do |prev|
        if item.qbegin == prev.qbegin && item.qend == prev.qend
          item.code = prev.code
          break
        end
      end
      next if item.code.count > 0
      index = 1

      (item.qbegin..item.qend).each do |row_index|
        row = sheet.row(row_index)
        if row[0].is_a? Numeric
          if is_dual_languages
            code = Code.new(row[0].to_i, row[1].to_s, index,row[2].to_s)
          else
            code = Code.new(row[0].to_i, row[1].to_s, index)
          end
          code.bold = true if sheet.font(row_index,2).present? && sheet.font(row_index,2).bold?
          item.code.push code
          index = index + 1
        end
      end
    end
    codelist
  end

  def build_codelist_array(codelist_params)
    ap __method__
    codelist = []
    codelist_tmp = []
    if codelist_params.include?("},")
      codelist_params.split("},").each do |str|
        codelist_tmp.push str + "}"
      end
      codelist_tmp[-1] = codelist_tmp[-1][0..-2]
    end

    if (codelist_tmp.length > 0)
      codelist_tmp.each do |codelist_row|
        codelist.push ActiveSupport::JSON.decode(codelist_row)
      end
    else
      codelist.push ActiveSupport::JSON.decode(codelist_params)
    end
    converted_codelist = []
    codelist.each_with_index do |item,index|
      unless item['question'] == 'New Row'
        s = CodeList.new(item['question'],item['filter'],item['qbegin'].to_i,item['qend'].to_i)
        converted_codelist.push s
      end
    end
    converted_codelist
  end

end