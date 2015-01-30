require 'axlsx'
# https://gist.github.com/randym/3980434
# https://gist.github.com/randym/3179305

class DataProcessingService < BaseService
  EXPORT_DATA_TYPE = [:count, :percent, :both ]

  def run
    helper_obj = HelperService.new

    encode = helper_obj.extract_zip_file("#{Rails.root}/public/uploads/HAZEL.zip")

    src_folder = RAILS_TEMP_PATH + encode + "/"



    p = Axlsx::Package.new
    p.use_shared_strings = true
    options = Hash.new
    options = {
      'build_index' => false,
      'all_in_one' => true,
      'num_of_digits' => 0, # Number of Digits after decimal
      'export_data_type' => :both,
      'orders' => {
        'question' => 1,
        'filters' => 2,
        'base' => 0,
        'wtd_resp' => 0,
        'resp' => 3,
        'header_and_data' => 4,
        'totals' => 8,
        'std_deviation' => 0,
        'means' => 5,
        'mode' => 6,
        'medians' => 7
      }
    }
    data = read_all_csv_files_in_folder(src_folder,options)
    export_excel_file = "#{Rails.root}/public/uploads/" + helper_obj.generate_name() + '.xlsx'


    wb = p.workbook

    all_in_one_sheet = nil
    write_excel_object = WriteExcelService.new
    all_in_one_sheet = wb.add_worksheet(name: 'DATA' ) if options['all_in_one']
    data.each do |x|
      if options['all_in_one']
        write_excel_object.export_data_to_excel(wb,x,options,all_in_one_sheet)
      else
        wb.add_worksheet(name: x['sheet_name']) do |sheet|
          write_excel_object.export_data_to_excel(wb,x,options,sheet)
        end
      end
    end

    p.serialize export_excel_file

    helper_obj.delete_folder_after_process(src_folder)

  end

  protected

  def read_all_csv_files_in_folder(src_folder,options)
    # Get All CSV file in Folder (and sub folder)
    # folder = Dir.glob(dir_path + "*.CSV")
    folder = Dir.glob(File.join(src_folder,"**","*.CSV"))
    # Sort Folder
    data = []
    read_csv_object = ReadCsvService.new
    folder.sort_by! { |f| File.basename(f) }
    folder.each do |csv_file|
      data << read_csv_object.read_csv(csv_file,options)
    end
    data
  end


end