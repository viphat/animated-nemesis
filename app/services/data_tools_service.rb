require 'axlsx'
# https://gist.github.com/randym/3980434
# https://gist.github.com/randym/3179305

class DataToolsService < BaseService
  attr_accessor :log_file

  def initialize()
    @log_file = nil
  end

  def read_question_and_build_json

  end

  def read_question_and_write_to_file(file,options,params)
    helper_obj = HelperService.new
    # Setup Log File and Excel File
    full_file_path = exported_file_path(options,helper_obj)
    ap "#{full_file_path}"
    @log_file = "#{full_file_path}.txt"
    @indexes = []
    log_file = File.open(@log_file,'w')
    log_file.write("Logs File for #{File.basename(file)}\n")
    log_file.write("#{Time.zone.now}")
    codelist_tools_service = CodeListToolsService.new
    unless codelist_tools_service.check_file_exists("#{File.basename(file)}")
      log_file.write("\nFile #{file} doesn't exits on Server")
      log_file.close
      raise "File #{file} doesn't exits on Server"
    end
    read_csv_obj = ReadCsvService.new
    encode = helper_obj.extract_zip_file("#{Rails.root}/#{file}")
    begin
      src_folder = RAILS_TEMP_PATH + "csv/"  + encode + "/"
      data = read_csv_obj.read_all_csv_files_in_folder(src_folder,options,@indexes,log_file)

      log_file.write("\nWrite Cookies\n\n\n")
      write_questions_to_file(data,full_file_path)
    rescue Exception => e
      log_file.write("\n#{e}")
      Airbrake.notify_or_ignore(
        e,
        :parameters    => params,
        :cgi_data      => ENV.to_hash
      )
      ap e
      raise e
    ensure
      helper_obj.delete_folder_after_process(src_folder)
      log_file.close
    end
    "#{full_file_path}.csv"
  end

  def export_data(file,options,params)
    @indexes = []
    helper_obj = HelperService.new
    # Setup Log File and Excel File
    full_file_path = exported_file_path(options,helper_obj)
    export_excel_file = "#{full_file_path}.xlsx"
    @log_file = "#{full_file_path}.txt"

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
      Airbrake.notify_or_ignore(
        e,
        :parameters    => params,
        :cgi_data      => ENV.to_hash
      )
      ap e
      raise e
    ensure
      helper_obj.delete_folder_after_process(src_folder)
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

  def exported_file_path(options,helper)
    ("#{Rails.root}/public/uploads/" +
    (options['output_file_name']? options['output_file_name'] : helper.generate_name()))
  end

  def build_orders(params,options)
    orders = params['orders']
    index = 1
    orders.split(",").each do |o|
      options['orders'][o] = index
      index += 1
    end
  end

  def write_questions_to_file(data,output_file_name)
    questions = []
    data.each do |d|
      questions.push d.question
    end
    questions = questions.uniq
    # directory = "public/uploads"
    # path = File.join(directory, File.basename(output_file_name,".*") + ".csv")
    # path = File.join(directory, "#{output_file_name}.csv")
    path = "#{output_file_name}.csv"
    File.open(path, "wb") { |f|
      f.puts("\"#\",\"Question\",\"Filter\",\"Begin\",\"End\"")
      questions.each do |q|
        f.puts("\"\",\"#{q}\",\"\",\"\",\"\"")
      end
    }
    path
  end

end