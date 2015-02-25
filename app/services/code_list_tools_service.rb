require 'roo'

class CodeListToolsService < BaseService

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
    filepath = "#{Rails.root}/public/uploads/#{filename}"
    return File.exist?(filepath)
  end

  def read_codelist_file(codelist_file)
    Roo::Excelx.new("myspreadsheet.xlsx")
  end

  def build_codelist_array(codelist_params)
    codelist = []
    codelist_tmp = []

    if codelist_params.include?("},")
      codelist_params.split("},").each do |str|
        codelist_tmp.push str + "}"
      end
    end

    if (codelist_tmp.length > 0)
      codelist_tmp.each do |codelist_row|
        codelist.push ActiveSupport::JSON.decode(codelist_row)
      end
    else
      codelist.push ActiveSupport::JSON.decode(codelist_params)
    end

    codelist
  end
end