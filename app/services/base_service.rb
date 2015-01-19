require 'csv'
require 'zip'
require 'base64'
require 'axlsx'
# https://gist.github.com/randym/3980434
# https://gist.github.com/randym/3179305
class BaseService
  EXPORT_DATA_TYPE = [:count, :percent, :both ]
  def run
    encode = extract_zip_file("#{Rails.root}/public/uploads/HAZEL.zip")
    src_folder = RAILS_TEMP_PATH + encode + "/"
    data = read_all_csv_files_in_folder(src_folder)
    p = Axlsx::Package.new
    p.use_shared_strings = true
    options = Hash.new
    options = {
      'build_index' => false,
      'all_in_one' => false,
      'export_data_type' => :percent,
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
    export_excel_file = "#{Rails.root}/public/uploads/" + generate_name + '.xlsx'
    wb = p.workbook
    data.each do |x|
      export_data_to_excel(wb,x,options)
    end
    p.serialize export_excel_file
    delete_folder_after_process(src_folder)
  end

  protected

  def read_all_csv_files_in_folder(src_folder)
    # Get All CSV file in Folder (and sub folder)
    # folder = Dir.glob(dir_path + "*.CSV")
    folder = Dir.glob(File.join(src_folder,"**","*.CSV"))
    # Sort Folder
    data = []
    folder.sort_by! { |f| File.basename(f) }
    folder.each do |csv_file|
      data << read_csv(csv_file)
    end
    data
  end

  def trim_and_downcase_a_string(string)
    string.split.join(" ").downcase unless string.nil?
  end

  def read_csv(csv_file)
    # max_cols = 0
    data = Hash.new
    header_flag = false
    header_label_flag = 0
    header_label = ""
    value = []
    data['val'] = []
    csv_file_name = File.basename(csv_file)
    data['sheet_name'] = csv_file_name[0] + csv_file_name[2..5].to_i.to_s
    CSV.foreach(csv_file) do |row|
      line = $INPUT_LINE_NUMBER
      # max_cols = row.length if row.length > max_cols
      unless row.all? { |x| x.blank? }
        first_cell = trim_and_downcase_a_string(row[0])
        # ap first_cell
        case
          when first_cell.include?('wtd. resp.') then
            data['wtd_resp'] = row
          when first_cell.include?('resp.') then
            data['resp'] = row
          when first_cell.include?('base') then
            data['base'] = first_cell.upcase
          when first_cell.include?('table') then
            data['table_name'] = first_cell.capitalize
          when first_cell.include?('filters')
            data['filters'] = first_cell.capitalize
          when first_cell === 'means' then
            data['means'] = row
          when first_cell === 'medians' then
            data['medians'] = row
          when first_cell === 'mode' then
            data['mode'] = row
          when first_cell === 'std. deviation' then
            data['std_deviation'] = row
          when first_cell === 'totals' then
            data['totals_count'] = row
          when !first_cell.blank? && data['table_name'] != nil && data['table_name'].include?(first_cell)
            data['question'] = first_cell.capitalize
          else
            if !header_flag
              if header_label.present? && header_label_flag + 1 != line
                # Header
                header_flag = true
                header_label = header_label.split.join(" ")
                data['header_label'] = header_label
                data['header'] = row
                data['header'][0] = header_label
              else
                # Header Label (Truong hop bi xuong dong khi Header qua ngan nhung Label Header qua dai nua)
                header_label += " " + row.second
                header_label_flag = line
              end
            else
              # Data
              unless first_cell.blank?
                # Count
                unless data['totals_count'].present?
                  value = {}
                  value['count'] = row
                end
              else

                # %
                if data['totals_count'].present?
                  data['totals_percent'] = row
                  data['totals_percent'][0] = data['totals_count'][0]
                else
                  value['percent'] = row
                  # Fill with 0
                  value['percent'][0] = value['count'][0]
                  value['percent'].map! do |x|
                    if x.blank?
                      x = "0"
                    else
                      x
                    end
                  end
                  data['val'] << value
                end
              end
            end
        end
      end #unless
    end #foreach CSV
    # ap data
    data
  end

  def add_question(sheet,data,options,predefined_styles)
    sheet.add_row([data['question']], widths: [:ignore] * data['header'].count,
                    style: predefined_styles['bold']) if !data['question'].nil?
  end

  def add_filters(sheet,data,options,predefined_styles)
    sheet.add_row([data['filters']], style: predefined_styles['blue_bold']) if !data['filters'].nil?
  end

  def add_base(sheet,data,options,predefined_styles)
    sheet.add_row([data['base']]) if !data['base'].nil?
  end

  def add_wtd_resp(sheet,data,options,predefined_styles)
    sheet.add_row(data['wtd_resp'], style: predefined_styles['blue_bold']) if !data['wtd_resp'].nil?
  end

  def add_resp(sheet,data,options,predefined_styles)
    sheet.add_row(data['resp'], style: predefined_styles['red_bold']) if !data['resp'].nil?
  end

  def add_header_and_data(sheet,data,options,predefined_styles)
      style_for_header = [predefined_styles['red_bold']]
      (data['header'].count - 1).times { style_for_header << predefined_styles['red_bold_with_center'] }
      start_row = sheet.add_row(data['header'], style: style_for_header,:widths => [:ignore] * data['header'].count)
      style_for_data = [predefined_styles['bold_border']]
      (data['header'].count - 1).times { style_for_data << predefined_styles['border'] }
      data['val'].each do |value|
        sheet.add_row(value['count'], style: style_for_data) if options['export_data_type'] == :count
        sheet.add_row(value['percent'], style: style_for_data) if options['export_data_type'] == :percent
      end
      last_cell = sheet.rows.last.cells.last
      sheet.auto_filter = Axlsx::cell_r(0,start_row.index) + ":" + Axlsx::cell_r(last_cell.index,last_cell.row.index)
  end

  def add_totals(sheet,data,options,predefined_styles)

      sheet.add_row(data['totals_count'],style: predefined_styles['bold']) if options['export_data_type'] == :count && !data['totals_count'].nil?
      sheet.add_row(data['totals_percent'],style: predefined_styles['bold']) if options['export_data_type'] == :percent && !data['totals_percent'].nil?

  end


  def add_std_deviation(sheet,data,options,predefined_styles)
    sheet.add_row(data['std_deviation'],style: predefined_styles['bold']) if !data['std_deviation'].nil?
  end

  def add_means(sheet,data,options,predefined_styles)
    sheet.add_row(data['means'],style: predefined_styles['bold']) if !data['means'].nil?
  end


  def add_mode(sheet,data,options,predefined_styles)
    sheet.add_row(data['mode'],style: predefined_styles['bold']) if !data['mode'].nil?
  end


  def add_medians(sheet,data,options,predefined_styles)
    sheet.add_row(data['medians'],style: predefined_styles['bold']) if !data['medians'].nil?
  end

  def export_data_to_excel(wb,data,options={})
    # http://axlsx.blog.randym.net/

    # Setup Styles
    styles = wb.styles
    predefined_styles = Hash.new

    predefined_styles['border'] = styles.add_style(:border => Axlsx::STYLE_THIN_BORDER)
    predefined_styles['bold'] = styles.add_style(:b => true)
    predefined_styles['bold_border'] = styles.add_style(:b => true, :border => Axlsx::STYLE_THIN_BORDER)
    predefined_styles['red_bold_with_center'] = styles.add_style(:b => true, :fg_color=>"FF0000",:alignment => { :horizontal => :center })
    predefined_styles['red_bold'] = styles.add_style(:b => true, :fg_color=>"FF0000")
    predefined_styles['red'] = styles.add_style(:fg_color=>"FF0000")
    predefined_styles['blue_bold'] = styles.add_style(:b => true, :fg_color=>"0000FF")

    sorted_order = options['orders'].sort_by{ |k,v| v }
    # ap sorted_order
    # binding.pry

    # Started to Write Data
    wb.add_worksheet(name: data['sheet_name']) do |sheet|

      sorted_order.each do |key,value|
        if value > 0
          eval("add_#{key}(sheet,data,options,predefined_styles)")
        end
      end

    end
  end

  def write_to_sheet(excel_file,sheet=1)
    # Choose what will display - Drag and Drop to Order
    # 2 Options All File in One Sheet (Tuong tu Tools Merge Sheet - Build Index or Not), 1 File Per Sheet (Build Index or Not)
    # 2 Style - SPSS Style (Count and Percentage) or Original Style ( Count or Percentage )
    # Paste Codelist for OA (Export a List with All Unique Question with Filters, Require User to Fill this List and Upload Codelist to Server to Continue Process)
  end

  def delete_folder_after_process(folder)
    # Tuong tu lenh rm -rf cua Unix
    FileUtils.rm_rf(folder)
  end

  def generate_name()
    Base64.encode64(Time.now.ctime)[0..-2]
  end

  def extract_zip_file(zip_file_path)
    # Overwritte if Exists
    # Zip.on_exists_proc = true
    # zip_file = RAILS_TEMP_PATH + zip_file_name
    encode = generate_name()
    des_folder = RAILS_TEMP_PATH + encode + "/"
    # Zip.on_exists_proc = true
    Dir.mkdir(des_folder) unless Dir.exists?(des_folder)
    Zip::File.open(zip_file_path) do |file|
      file.each do |entry|
        # puts "Extracting #{entry.name}"
        entry.extract(des_folder + entry.name)
      end
    end
    encode
  end

end