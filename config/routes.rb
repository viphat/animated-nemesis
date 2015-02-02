Rails.application.routes.draw do
  root 'dashboard#index'
  post '/data_tools/post', to: 'dashboard#create', as: 'run_data_tools'
end
