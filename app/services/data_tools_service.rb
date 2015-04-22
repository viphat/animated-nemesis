require 'axlsx'
# https://gist.github.com/randym/3980434
# https://gist.github.com/randym/3179305

class DataToolsService < BaseService
  attr_accessor :log_file

  def initialize()
    @log_file = nil
  end

  def read_question_and_build_json(file,options,params)
    ap __method__
    @indexes = []
    helper_obj = HelperService.new
    read_csv_obj = ReadCsvService.new
    full_file_path = exported_file_path(options,helper_obj)
    log_file = build_log_file(file,full_file_path)

    begin
      encode = helper_obj.extract_zip_file(file)
      src_folder = RAILS_TEMP_PATH + "csv/"  + encode + "/"
      data = read_csv_obj.read_all_csv_files_in_folder(src_folder,options,@indexes,log_file)
      log_file.write("\n\n\nWrite JSON Data\n")
      full_file_path = write_data_to_json_file(data,options,@indexes,full_file_path)
    rescue Exception => e
      log_file.write("\n\n\n#{e}")
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
    full_file_path
  end

  def read_question_and_write_to_file(file,options,params)
    ap __method__
    @indexes = []
    helper_obj = HelperService.new
    read_csv_obj = ReadCsvService.new
    full_file_path = exported_file_path(options,helper_obj)
    log_file = build_log_file(file,full_file_path)
    begin
      encode = helper_obj.extract_zip_file(file)
      src_folder = RAILS_TEMP_PATH + "csv/"  + encode + "/"
      data = read_csv_obj.read_all_csv_files_in_folder(src_folder,options,@indexes,log_file)
      log_file.write("\n\n\nWrite Cookies\n")
      full_file_path = write_questions_to_file(data,full_file_path)
    rescue Exception => e
      log_file.write("\n\n\n#{e}")
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
    full_file_path
  end

  def read_and_export_data(file,options,params,is_codelist=false,data=nil,indexes=nil)
    ap __method__
    @indexes = []

    @indexes = indexes if is_codelist

    helper_obj = HelperService.new
    full_file_path = exported_file_path(options,helper_obj)

    export_excel_file = "#{full_file_path}.xlsx"
    log_file = build_log_file(file,full_file_path)

    read_csv_obj = ReadCsvService.new
    index_helper = IndexHelper.new
    write_excel_object = WriteExcelService.new

    begin
      # Read CSV Folder
      unless is_codelist
        encode = helper_obj.extract_zip_file(file) # Extract File ZIP
        src_folder = RAILS_TEMP_PATH + "csv/"  + encode + "/"
        data = read_csv_obj.read_all_csv_files_in_folder(src_folder,options,@indexes,log_file)
      end
      # Write to Excel File
      log_file.write("\n\nWriting Data After Processing:\n")
      p = Axlsx::Package.new
      p.use_shared_strings = true
      wb = p.workbook
      all_in_one_sheet = nil
      index_sheet = wb.add_worksheet(name: 'INDEX' ) if options['build_index']
      all_in_one_sheet = wb.add_worksheet(name: 'DATA' ) if options['all_in_one']

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
      log_file.write("\n\n\n#{e}")
      Airbrake.notify_or_ignore(
        e,
        :parameters    => params,
        :cgi_data      => ENV.to_hash
      )
      ap e
      raise e
    ensure
      unless is_codelist
        helper_obj.delete_folder_after_process(src_folder)
      end
      log_file.close
    end
    export_excel_file
  end

  # protected

  def build_options(params)
    options = {
      'build_index' => false,
      'all_in_one' => false,
      'clean_empty_code' => false,
      'clean_empty_table' => false,
      'clean_empty_header' => false,
      'output_file_name' => false,
      'num_of_digits' => 0, # Number of Digits after decimal
      'export_data_type' => :both,
      'export_first' => :count,
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
    options['export_data_type'] = :both if params['export_data_type'] == 'both_count' || params['export_data_type'] == 'both_percent'

    options['export_first'] = :count if params['export_data_type'] == 'both_count'
    options['export_first'] = :percent if params['export_data_type'] == 'both_percent'

    if params['output_file_name'].present? &&  params['output_file_name'] != ''
      options['output_file_name'] = File.basename(params['output_file_name'])
    end

    build_orders(params,options)
    options
  end

  private

  def build_log_file(file,full_file_path)
    @log_file = "#{full_file_path}.txt"
    log_file = File.open(@log_file,'w')
    log_file.write("Logs File for #{File.basename(file)}\n")
    log_file.write("#{Time.zone.now}")
    helper_obj = HelperService.new
    unless helper_obj.check_file_exists("#{File.basename(file)}")
      log_file.write("\n\n\nFile #{file} doesn't exits on Server")
      log_file.close
      raise "File #{file} doesn't exits on Server"
    end
    log_file
  end

  def exported_file_path(options,helper)
    (
      "#{Rails.root}/public/uploads/" +
      (options['output_file_name']? "#{options['output_file_name']}_#{options['export_data_type']}" : "#{helper.generate_name()}_#{options['export_data_type']}" )
    )
  end

  def build_orders(params,options)
    orders = params['orders']
    index = 1
    orders.split(",").each do |o|
      options['orders'][o] = index
      index += 1
    end
  end

  def write_data_to_json_file(data,options,indexes,output_file_name)

    path = "#{output_file_name}.json"
    File.open(path,"wb") { |f|
      f.puts({
        data: data,
        options: options,
        indexes: indexes
      }.to_json)
    }
    path
  end

  def write_questions_to_file(data,output_file_name)
    questions = []
    data.each do |d|
      questions.push d.question
    end
    questions = questions.uniq
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