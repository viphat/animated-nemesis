require 'axlsx'
# https://gist.github.com/randym/3980434
# https://gist.github.com/randym/3179305

class DataToolsService < BaseService
  EXPORT_DATA_TYPE = [:count, :percent, :both ]

  def run(file,options)

    @indexes = []

    helper_obj = HelperService.new
    read_csv_obj = ReadCsvService.new

    # encode = helper_obj.extract_zip_file("#{Rails.root}/public/uploads/BILLIO2.zip")
    encode = helper_obj.extract_zip_file("#{Rails.root}/#{file}")

    src_folder = RAILS_TEMP_PATH + encode + "/"

    p = Axlsx::Package.new
    p.use_shared_strings = true


    data = read_csv_obj.read_all_csv_files_in_folder(src_folder,options,@indexes)

    wb = p.workbook

    all_in_one_sheet = nil
    write_excel_object = WriteExcelService.new

    index_sheet = wb.add_worksheet(name: 'INDEX' ) if options['build_index']

    index_helper = IndexHelper.new
    all_in_one_sheet = wb.add_worksheet(name: 'DATA' ) if options['all_in_one']

    data.each_with_index do |x,i|
      if options['all_in_one']
        @indexes[i].link = (write_excel_object.export_data_to_excel(wb,x,options,all_in_one_sheet)).to_s
        @indexes[i].link = "'DATA'!A#{@indexes[i].link}"
      else
        wb.add_worksheet(name: x.sheet_name) do |sheet|
          write_excel_object.export_data_to_excel(wb,x,options,sheet)
          @indexes[i].link = "'#{x.sheet_name}'!A1"
        end
      end
    end
    # ap @indexes
    blue_link = wb.styles.add_style :fg_color => '0000FF'
    index_helper.process_and_write_indexes_to_excel(index_sheet,@indexes, blue_link) if options['build_index']


    if options['output_file_name']
      export_excel_file = "#{Rails.root}/public/uploads/#{options['output_file_name']}.xlsx"
    else
      export_excel_file = "#{Rails.root}/public/uploads/#{helper_obj.generate_name()}.xlsx"
    end

    p.serialize export_excel_file

    helper_obj.delete_folder_after_process(src_folder)

    export_excel_file
  end

  # protected

  def build_options(params)
    options = {
      'build_index' => true,
      'all_in_one' => false,
      'clean_empty_code' => false,
      'clean_empty_table' => false,
      'clean_empty_header' => false,
      'output_file_name' => false,
      'num_of_digits' => 0, # Number of Digits after decimal
      'export_data_type' => :both,
      'orders' => {
        'question' => 0,
        'filters' => 0,
        'base' => 0,
        'wtd_resp' => 0,
        'resp' => 0,
        'header_and_data' => 0,
        'totals' => 0,
        'std_deviation' => 0,
        'means' => 0,
        'mode' => 0,
        'medians' => 0
      }
    }

    options['build_index'] = true if params['build_index'].present?
    options['all_in_one'] = true if params['all_in_one'].present?
    options['clean_empty_code'] = true if params['clean_empty_code'].present?
    options['clean_empty_table'] = true if params['clean_empty_table'].present?
    options['clean_empty_header'] = true if params['clean_empty_header'].present?
    options['num_of_digits'] = params['num_of_digits']

    options['export_data_type'] = :percent if params['export_data_type'] = 'percent'
    options['export_data_type'] = :count if params['export_data_type'] = 'count'
    options['export_data_type'] = :both if params['export_data_type'] = 'both'

    if params['output_file_name'].present? &&  params['output_file_name'] != ''
      options['output_file_name'] = File.basename(params['output_file_name'])
    end

    build_orders(params,options)

    options
  end

  private

  def build_orders(params,options)
    orders = params['orders']
    index = 1
    orders.split(",").each do |o|
      options['orders'][o] = index
      index += 1
    end
  end

end