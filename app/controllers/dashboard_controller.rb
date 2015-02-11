class DashboardController < ApplicationController

  def index

  end

  def create
    data_file = params[:data_file]
    if File.extname(data_file.original_filename) != ".zip"
      flash[:notice] = "Định dạng Tập tin không được chấp nhận. <br/> Bạn chỉ được upload <strong>tập tin zip</strong>. "
      render "index"
      return
    end
    data_tools = DataToolsService.new
    begin
      uploaded_file = DataFile.save(data_file)
      options = Hash.new
      options = data_tools.build_options(params)
      result_file = data_tools.run(uploaded_file,options,params)
      unless result_file == false
        send_file result_file
      end
    rescue Exception => e
      send_file data_tools.log_file
      # raise e
    end

  end

  private
  def strong_params
    params.permit(:data_file,:output_file_name,:export_data_type, :num_of_digits,
                  :build_index,:all_in_one,:clean_empty_code,:clean_empty_table,
                  :clean_empty_header,:orders
                 )
  end
end
