class ApplicationController < ActionController::Base
  # Prevent CSRF attacks by raising an exception.
  # For APIs, you may want to use :null_session instead.
  # protect_from_forgery with: :exception
  before_filter :authenticate
  protect_from_forgery with: :null_session

  protected
  def authenticate
    # authenticate_or_request_with_http_basic do |username, password|
    #   username == "intage" && password == "qweCXZ!@#"
    # end
  end

end
