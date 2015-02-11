require 'zip'
require 'base64'

class HelperService < BaseService

  def delete_folder_after_process(folder)
    # Tuong tu lenh rm -rf cua Unix
    FileUtils.rm_rf(folder)
  end

  def generate_name()
    Base64.encode64(Time.zone.now.ctime)[0..-2]
  end

  def extract_zip_file(zip_file_path)
    # Overwritte if Exists
    # Zip.on_exists_proc = true
    # zip_file = RAILS_TEMP_PATH + zip_file_name
    encode = generate_name()
    des_folder = RAILS_TEMP_PATH + "csv/"  + encode + "/"
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

  def trim_and_downcase_a_string(string)
    string.split.join(" ").downcase unless string.nil?
  end


end