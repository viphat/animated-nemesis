class DashboardController < ApplicationController

  def index
  end

  def codelist

  end

  def download_csv

    data_tools = DataToolsService.new
    begin
      if params['data_file'].present?
        data_file = params[:data_file]
        if File.extname(data_file.original_filename) != ".zip"
          flash[:notice] = "Định dạng Tập tin không được chấp nhận. <br/> Bạn chỉ được upload <strong>tập tin zip</strong>. "
          render "index"
          return
        end
        uploaded_file = DataFile.save(data_file)
      else
        uploaded_file = "#{Rails.root}/public/uploads/#{params['local_storage_data']}"
      end
      options = Hash.new
      options = data_tools.build_options(params)
      # binding.pry
      result_file = data_tools.read_question_and_write_to_file(uploaded_file,options,params)
      send_file result_file
    rescue Exception => e
      # binding.pry
      send_file data_tools.log_file
    end
  end

  def codelist_process
    codelist_tools = CodeListToolsService.new

    if params['codelist_file'].present?
      codelist_file = DataFile.save_excel(params['codelist_file'])
    else
      codelist_file = "public/uploads/#{params['local_storage_codelist']}"
    end

    if params['data_file'].present?
      data_file = DataFile.save_excel(params['data_file'])
    else
      data_file = "public/uploads/#{params['local_storage_data']}"
    end
    codelist = codelist_tools.build_codelist_array(params['codelist'])
    codelist = codelist_tools.read_codelist_file("#{Rails.root}/#{codelist_file}",params['sheet'].to_i,codelist)
    # Read Codelist File
    render nothing: true
  end

  def check_file_exists
    codelist_tools = CodeListToolsService.new
    if codelist_tools.check_file_exists(params[:file])
      render json: {status: true}
    else
      render json: {status: false}
    end
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
      result_file = data_tools.export_data(uploaded_file,options,params)
      unless result_file == false
        # cookies["DataTools.codelist_csv_file"] = File.basename(result_file,".*") + ".csv"
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
