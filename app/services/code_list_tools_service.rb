require 'spreadsheet'
require 'roo'
require 'roo-xls'
require 'json'

class CodeListToolsService < BaseService

  def paste_codelist
    sheet = 1
    codelist_file = "public/uploads/Code List Final MOVING -2812015 - CS revised - V02.xls"
    json_file = "public/uploads/VGh1IEZlYiAyNiAxNzoxNzowMCAyMDE1.json"
    codelist_string = "{\"question\":\"A9.1\",\"filter\":\"\",\"qbegin\":\"21\",\"qend\":\"137\"},{\"question\":\"B3 nho ve fanpage\",\"filter\":\"\",\"qbegin\":\"21\",\"qend\":\"137\"},{\"question\":\"B10 diem nho qc\",\"filter\":\"\",\"qbegin\":\"21\",\"qend\":\"137\"},{\"question\":\"D2.2 khong thich\",\"filter\":\"\",\"qbegin\":\"141\",\"qend\":\"185\"},{\"question\":\"D2.1 thich\",\"filter\":\"\",\"qbegin\":\"189\",\"qend\":\"304\"},{\"question\":\"B7 xac nhan xem qc\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B8 thay qc o dau\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B9 nhan hieu trong doan qc\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B11 thong diep\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B12 muc do thich\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B13 diem thich\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B14 diem khong thich\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B15 muc do doc dao khac biet\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B16 diem doc dao khac biet\",\"filter\":\"\",\"qbegin\":\"375\",\"qend\":\"399\"},{\"question\":\"B17 muc do hap dan\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B18 muc do phu hop\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B20 muc do vui nhon\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B21 muc do muon mua\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B23 diem cai thien\",\"filter\":\"\",\"qbegin\":\"403\",\"qend\":\"412\"},{\"question\":\"New Row\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"}"
    codelist = build_codelist_array(codelist_string)
    codelist = read_codelist_file("#{Rails.root}/#{codelist_file}",sheet,codelist)
    data = build_data_from_json("#{Rails.root}/#{json_file}")
    options = build_options_from_json("#{Rails.root}/#{json_file}")
    data = parse_data(data,codelist)
    data
  end

  private

  def parse_data(data,codelist)
    codelist.each do |item|
      data.each do |datum|
        if datum.question != item.question
          next
        end
        if item.filters != "" && !(datum.filters?(item.filters))
          next
        end
        datum = insert_label(datum,item)
        datum = insert_into(datum)
      end
    end
    # Insert them mot o trong vao Means, Mode, Medians, Std deviation, Totals_count, Totals_Percent, Resp, Wtd_Resp, Header
    # Val - Insert Label
    # Val - Sap xep theo Index
    data
  end

  def integer?(str)
    /\A[+-]?\d+\z/ === str
  end

  def insert_label(data,codelist)
    codes = []
    codelist.code.each do |c|
      codes << c.code
    end
    data.val.each do |item|
      if integer?(item['count'][0])
        index = codes.index(item['count'][0].to_i)
        unless index.nil?
          # Tim Thay
          label = codelist.code[index].label
          item['count'].insert(1,label)
          item['percent'].insert(1,label)
          next
        end
      end
      item['count'].insert(1,"")
      item['percent'].insert(1,"")
    end
    c_index = 0
    codelist.code.each do |c|
      data_codes = []
      data.val.each { |x| data_codes << x['percent'][0].to_i }
      index = data_codes.index(c.code)
      unless index.nil?
        data.val.insert(c_index,data.val.delete_at(index))
        c_index += 1
      end
    end
    data
  end

  def insert_into(data)
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

  def build_options_from_json(json_file)
    file = File.open(json_file,"rb")
    contents = file.read
    # Parse Json Data to Object
    json_data = JSON.parse(contents)
    json_data['options']
  end

  def build_data_from_json(json_file)
    # Read Json Data from File
    file = File.open(json_file,"rb")
    contents = file.read
    # Parse Json Data to Object
    json_data = JSON.parse(contents)
    draft = json_data['data']
    data = []
    draft.each do |item|
      raw_data = RawData.new
      raw_data.build!(item)
      data << raw_data
    end
    data
  end

  def read_codelist_file(codelist_file,sheet,codelist)
    # Spreadsheet.client_encoding = 'UTF-8'
    excel = Roo::Spreadsheet.open(codelist_file)
    sheet = excel.sheet(sheet-1)

    codelist.each_with_index do |item,index|
      if item.qbegin == 0 || item.question == 'New Row'
        next
      end
      index = 1
      (item.qbegin..item.qend).each do |row_index|
        row = sheet.row(row_index)
        if row[0].is_a? Numeric
          code = Code.new(row[0].to_i, row[1].to_s, index)
          code.bold = true if sheet.font(row_index,2).bold?
          item.code.push code
          index = index + 1
        end
      end
    end
    codelist
  end

  def build_codelist_array(codelist_params)
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