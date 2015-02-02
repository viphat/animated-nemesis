require 'csv'
require 'string'

class ReadCsvService < BaseService

  def read_all_csv_files_in_folder(src_folder,options,indexes)
    # Get All CSV file in Folder (and sub folder)
    # folder = Dir.glob(dir_path + "*.CSV")
    folder = Dir.glob(File.join(src_folder,"**","*.CSV"))
    # Sort Folder
    data = []
    index_helper = IndexHelper.new
    folder.sort_by! { |f| File.basename(f) }
    folder.each do |csv_file|
      d = read_csv(csv_file,options)
      data << d if d
      indexes << index_helper.build_index_object(d) if d
    end
    data
  end

  protected

  def read_csv(csv_file,options)
    # max_cols = 0
    data = RawData.new
    header_flag = false
    header_label_flag = 0
    header_label = ""
    value = {}
    csv_file_name = File.basename(csv_file)
    data.sheet_name = csv_file_name[0] + csv_file_name[2..5].to_i.to_s
    helper_obj = HelperService.new
    # ap csv_file
    CSV.foreach(csv_file) do |row|

      line = $INPUT_LINE_NUMBER
      # max_cols = row.length if row.length > max_cols
      unless row.all? { |x| x.blank? }
        first_cell = helper_obj.trim_and_downcase_a_string(row[0])
        # ap first_cell
        case
          when first_cell.include?('wtd. resp.') then
            data.wtd_resp = row
          when first_cell.include?('resp.') then
            data.resp = row
          when first_cell.include?('base') then
            data.base = first_cell.upcase
          when first_cell.include?('table') then
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
          else
            if header_flag == false
              # binding.pry
              if header_label.present? && header_label_flag + 1 != line
                # Header
                header_flag = true
                header_label = header_label.split.join(" ")
                data.header_label = header_label
                data.header= row
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
                  data.val << value if (options['clean_empty_table'] == false) || (options['clean_empty_table'] == true &&  !check_empty_code(value['count']))
                end
              end
            end
        end
      end #unless
    end #foreach CSV
    return false if options['clean_empty_table'] && check_empty_table(data.resp)
    data
  end

  private

  def check_empty_code(val_arr)
    val_arr[1..-1].inject(:+) == 0
  end

  def check_empty_header(resp_arr,i)
    resp_arr[i] == 0
  end

  def delete_header_column(data,col_to_delete)

  end

  def check_empty_table(resp_arr)
    resp_arr[1..-1].map(&:to_i).inject(:+) == 0
  end

end