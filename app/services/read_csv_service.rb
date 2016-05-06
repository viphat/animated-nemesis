require 'csv'
require 'string'

class ReadCsvService < BaseService

  def read_all_csv_files_in_folder(src_folder,options,indexes,log_file=nil)
    # Get All CSV file in Folder (and sub folder)
    # folder = Dir.glob(dir_path + "*.CSV")
    folder = Dir.glob(File.join(src_folder,"**","*.CSV"))
    # Sort Folder
    data = []
    index_helper = IndexHelper.new
    log_file.write("\nReading Data from Source:\n\n\n") unless log_file.nil?
    folder.sort_by! { |f| File.basename(f) }
    folder.each do |csv_file|
      ap csv_file
      log_file.write("\n#{File.basename(csv_file)}") unless log_file.nil?
      d = read_csv(csv_file,options)
      data << d if d
      indexes << index_helper.build_index_object(d) if d
    end
    data
  end

  def read_csv(csv_file,options)
    ap "#{__method__} #{csv_file}"
    # max_cols = 0
    csv_text = File.read(csv_file).encode!('UTF-8', 'binary', invalid: :replace, undef: :replace, replace: '')

    csv = CSV.parse(csv_text, headers: false)
    data = RawData.new
    header_flag = false
    header_label_flag = 0
    header_label = ""
    value = {}
    csv_file_name = File.basename(csv_file)
    data.sheet_name = csv_file_name[0] + csv_file_name[2..5].to_i.to_s
    helper_obj = HelperService.new
    # ap csv_file
    # CSV.foreach(csv_file) do |row|
    csv.each_with_index do |row,line|
      next if row.all? { |x| x == "" }
      # line = $INPUT_LINE_NUMBER
      # max_cols = row.length if row.length > max_cols
      unless row.all? { |x| x.blank? }
        first_cell = helper_obj.trim_and_downcase_a_string(row[0])
        ap first_cell
        case
          when first_cell.include?('base') then
            data.base = first_cell.upcase
          when first_cell.include?('weights:') then
            data.weight = first_cell.upcase
          when first_cell.include?('wtd. resp.') then
            data.wtd_resp = row
          when first_cell.include?('resp.') then
            data.resp = row
          when first_cell.starts_with?('table') then
            data.table_name = first_cell.capitalize
          when data.question.nil? && !first_cell.blank? && data.table_name != nil && data.table_name.include?(first_cell)
            data.question = first_cell.capitalize
          when first_cell.include?('filters')
            data.filters = first_cell.capitalize
          when first_cell === 'means' then
            data.means = row
          when first_cell === 'medians' then
            data.medians = row
          when first_cell === 'mode' then
            data.mode = row
          when first_cell === 'std. deviation' then
            data.std_deviation = row
          when first_cell === 'totals' then
            data.totals_count = row
            if check_empty_code(data.totals_count)
              data.totals_percent = row
              data.totals_percent[0] = data.totals_count[0]
            end
          else
            if header_flag == false
              if header_label.present? && header_label_flag + 1 != line
                # Header
                header_flag = true
                header_label = header_label.split.join(" ")
                data.header_label = header_label
                data.header = row
                data.header[0] = header_label
              else
                # Header Label (Truong hop bi xuong dong khi Header qua ngan nhung Label Header qua dai nua)
                header_label += " " + row.join(" ").split.join(" ")
                header_label_flag = line
              end
            else
              # Data
              unless first_cell.blank?
                # Count
                if data.totals_count.nil?
                  value = {}
                  value['count'] = row
                  if check_empty_code(value['count'])
                    value['percent'] = row
                    value['percent'][0] = value['count'][0]
                    data.val << value if (options['clean_empty_code'] == false)
                  end
                end
              else
                # %
                if data.totals_count.present?
                  row.map! do |x|
                    if x.blank?
                      x = "0"
                    else
                      x.is_number? ? x.to_f.round(options['num_of_digits'].to_i) : x
                    end
                  end
                  data.totals_percent = row
                  data.totals_percent[0] = data.totals_count[0]
                else
                  value['percent'] = row
                  # Fill with 0
                  value['percent'][0] = value['count'][0]
                  value['percent'].map! do |x|
                    if x.blank?
                      x = "0"
                    else
                      x.is_number? ? x.to_f.round(options['num_of_digits'].to_i) : x
                    end
                  end
                  data.val << value if (options['clean_empty_code'] == false) || (options['clean_empty_code'] == true &&  !check_empty_code(value['count']))
                end
              end
            end
        end
      end #unless
    end #foreach CSV
    return false if options['clean_empty_table'] && check_empty_table(data.resp)
    if options['clean_empty_header'] == true
      clean_columns = find_zero_header_columns(data)
      if clean_columns.present?
        clean_columns.each_with_index do |col_to_delete,index|
          data = delete_header_column(data,col_to_delete-index)
        end
      end
    end
    data
  end

  def check_empty_code(val_arr)
    val_arr[1..-1].map(&:to_i).inject(:+) == 0
  end

  def check_empty_header(resp_arr,i)
    resp_arr[i] == 0
  end

  def find_zero_header_columns(data)
    zero = []
    if data.resp.present?
      data.resp.each_with_index do |val,index|
        if val.to_i == 0
          zero << index
        end
      end
    end
    if data.wtd_resp.present?
      data.wtd_resp.each_with_index do |val,index|
        if val.to_i == 0
          zero << index
        end
      end
    end
    zero.delete_at(0)
    zero.uniq
  end

  def delete_header_column(data,col_to_delete)
    data.resp.delete_at(col_to_delete) if data.resp.present?
    data.wtd_resp.delete_at(col_to_delete) if data.wtd_resp.present?
    data.header.delete_at(col_to_delete) if data.header.present?
    data.means.delete_at(col_to_delete) if data.means.present?
    data.mode.delete_at(col_to_delete) if data.mode.present?
    data.medians.delete_at(col_to_delete) if data.medians.present?
    data.std_deviation.delete_at(col_to_delete) if data.std_deviation.present?
    data.totals_percent.delete_at(col_to_delete) if data.totals_percent.present?
    data.totals_count.delete_at(col_to_delete) if data.totals_count.present?
    data.val.each do |value|
      value['count'].delete_at(col_to_delete) if value['count'].present?
      value['percent'].delete_at(col_to_delete) if value['percent'].present?
    end
    data
  end

  def check_empty_table(resp_arr)
    resp_arr[1..-1].map(&:to_i).inject(:+) == 0
  end

end
