class DataFile < ActiveRecord::Base

  def self.save_excel(excel_file)
    name = excel_file.original_filename
    directory = "#{Rails.root}/public/uploads"
    path = File.join(directory, name)
    # write the file
    File.open(path, "wb") { |f| f.write(excel_file.read) }
    path
  end

  def self.save(data_file)
    name = data_file.original_filename
    # helper_obj = HelperService.new
    # new_name = "#{helper_obj.generate_name}.zip"
    # if File.extname(new_name) != ".zip"
    #   name = File.basename(new_name) + ".zip"
    # end
    directory = "#{Rails.root}/public/uploads"
    # create the file path
    path = File.join(directory, name)
    # write the file
    File.open(path, "wb") { |f| f.write(data_file.read) }
    path
  end
end