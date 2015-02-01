require 'axlsx'
# https://gist.github.com/randym/3980434
# https://gist.github.com/randym/3179305

class DataProcessingService < BaseService
  EXPORT_DATA_TYPE = [:count, :percent, :both ]

  def run

    @indexes = []

    helper_obj = HelperService.new

    encode = helper_obj.extract_zip_file("#{Rails.root}/public/uploads/BILLIO2.zip")

    src_folder = RAILS_TEMP_PATH + encode + "/"

    p = Axlsx::Package.new
    p.use_shared_strings = true
    options = Hash.new
    options = {
      'build_index' => true,
      'codelist' => false,
      'all_in_one' => false,
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
    # ap data

    export_excel_file = "#{Rails.root}/public/uploads/" + helper_obj.generate_name() + '.xlsx'

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
    ap @indexes
    blue_link = wb.styles.add_style :fg_color => '0000FF'
    index_helper.process_and_write_indexes_to_excel(index_sheet,@indexes, blue_link) if options['build_index']

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
    index_helper = IndexHelper.new
    folder.sort_by! { |f| File.basename(f) }
    folder.each do |csv_file|
      d = read_csv_object.read_csv(csv_file,options)
      data << d
      @indexes << index_helper.build_index_object(d)
    end
    data
  end


end