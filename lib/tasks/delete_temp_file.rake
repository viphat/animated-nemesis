desc "Delete User's Temporary Files and Folder"
task :delete_files_and_folders => :environment do

  p "---Delete tmp Folder"
  p "-------#{RAILS_TEMP_PATH}"
  FileUtils.rm_rf(Dir.glob("#{RAILS_TEMP_PATH}csv/*"))

  p "---Delete Upload and Result Files"
  p "------#{Rails.root}/public/uploads/"
  FileUtils.rm_rf(Dir.glob("#{Rails.root}/public/uploads/*"))

end


