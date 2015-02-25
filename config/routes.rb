Rails.application.routes.draw do
  root 'dashboard#index'
  post '/data_tools/post', to: 'dashboard#create', as: 'run_data_tools'
  post '/codelist_tools/post', to: 'dashboard#codelist_process', as: 'run_codelist_tools'
  get '/codelist', to: 'dashboard#codelist'
  get '/check_file_exists', to: 'dashboard#check_file_exists', as: 'check_file_exists'
  get '/download_csv', to: 'dashboard#download_csv', as: 'download_csv'
end
