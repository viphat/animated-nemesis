require 'axlsx'
# https://gist.github.com/randym/3980434
# https://gist.github.com/randym/3179305

class DataToolsService < BaseService
  attr_accessor :log_file

  def initialize()
    @log_file = nil
  end

  def run(file,options)
    @indexes = []

    helper_obj = HelperService.new


    # Setup Log File and Excel File
    export_excel_file = "#{Rails.root}/public/uploads/#{helper_obj.generate_name()}.xlsx"
    @log_file = "#{Rails.root}/public/uploads/#{helper_obj.generate_name()}.txt"

    if options['output_file_name']
      export_excel_file = "#{Rails.root}/public/uploads/#{options['output_file_name']}.xlsx"
      @log_file = "#{Rails.root}/public/uploads/#{options['output_file_name']}.txt"
    end

    log_file = File.open(@log_file,'w')
    log_file.write("Logs File for #{File.basename(file)}\n")
    log_file.write("#{Time.zone.now}")

    # Setup Var
    read_csv_obj = ReadCsvService.new
    index_helper = IndexHelper.new
    write_excel_object = WriteExcelService.new
    encode = helper_obj.extract_zip_file("#{Rails.root}/#{file}") # Extract File ZIP
    src_folder = RAILS_TEMP_PATH + "csv/"  + encode + "/"
    p = Axlsx::Package.new
    p.use_shared_strings = true
    wb = p.workbook
    all_in_one_sheet = nil
    index_sheet = wb.add_worksheet(name: 'INDEX' ) if options['build_index']
    all_in_one_sheet = wb.add_worksheet(name: 'DATA' ) if options['all_in_one']

    begin
      # Read CSV Folder
      data = read_csv_obj.read_all_csv_files_in_folder(src_folder,options,@indexes,log_file)

      log_file.write("\nWriting Data After Processing:\n\n\n")

      # Write to Excel File
      data.each_with_index do |x,i|
        log_file.write("\n#{x.sheet_name}")
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
      # Build Index if needed
      blue_link = wb.styles.add_style :fg_color => '0000FF'
      index_helper.process_and_write_indexes_to_excel(index_sheet,@indexes, blue_link) if options['build_index']

      # Finalize and Save to Excel File
      p.serialize export_excel_file

    rescue Exception => e
      log_file.write("\n#{e}")
      raise e
      Airbrake.notify_or_ignore(
        e,
        :parameters    => params,
        :cgi_data      => ENV.to_hash
      )
    ensure
      # Delete CSV Files Folder
      helper_obj.delete_folder_after_process(src_folder)
      # Close Log File
      log_file.close
    end
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

    options['export_data_type'] = :percent if params['export_data_type'] == 'percent'
    options['export_data_type'] = :count if params['export_data_type'] == 'count'
    options['export_data_type'] = :both if params['export_data_type'] == 'both'

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