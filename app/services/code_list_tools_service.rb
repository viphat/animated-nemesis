require 'spreadsheet'
require 'roo'
require 'roo-xls'
class CodeListToolsService < BaseService

  def codelist_process
    sheet = 1
    codelist_file = "public/uploads/Code List Final MOVING -2812015 - CS revised - V02.xls"
    codelist_string = "{\"question\":\"A9.1\",\"filter\":\"\",\"qbegin\":\"21\",\"qend\":\"137\"},{\"question\":\"B3 nho ve fanpage\",\"filter\":\"\",\"qbegin\":\"21\",\"qend\":\"137\"},{\"question\":\"B10 diem nho qc\",\"filter\":\"\",\"qbegin\":\"21\",\"qend\":\"137\"},{\"question\":\"D2.2 khong thich\",\"filter\":\"\",\"qbegin\":\"141\",\"qend\":\"185\"},{\"question\":\"D2.1 thich\",\"filter\":\"\",\"qbegin\":\"189\",\"qend\":\"304\"},{\"question\":\"B7 xac nhan xem qc\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B8 thay qc o dau\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B9 nhan hieu trong doan qc\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B11 thong diep\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B12 muc do thich\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B13 diem thich\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B14 diem khong thich\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B15 muc do doc dao khac biet\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B16 diem doc dao khac biet\",\"filter\":\"\",\"qbegin\":\"375\",\"qend\":\"399\"},{\"question\":\"B17 muc do hap dan\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B18 muc do phu hop\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B20 muc do vui nhon\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B21 muc do muon mua\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"},{\"question\":\"B23 diem cai thien\",\"filter\":\"\",\"qbegin\":\"403\",\"qend\":\"412\"},{\"question\":\"New Row\",\"filter\":\"\",\"qbegin\":\"\",\"qend\":\"\"}"
    codelist = build_codelist_array(codelist_string)
    codelist = read_codelist_file("#{Rails.root}/#{codelist_file}",1,codelist)
  end

  def build_options(params)
    options = {
      codelist_file: '',
      data_file: '',
      codelist_sheet: 1,
      output_file_name: false
    }
    # Arrange Code
    # Tach Cot
    options
  end

  def check_file_exists(filename)
    ap "check_file_exists - #{filename}"
    filepath = "#{Rails.root}/public/uploads/#{filename}"
    return File.exist?(filepath)
  end

  def read_codelist_file(codelist_file,sheet,codelist)
    # reload!; cl = CodeListToolsService.new; cl.run;

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