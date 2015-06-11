Airbrake.configure do |config|
  config.api_key = 'ae51f4226931ebe48b21aa9f87b2b442'
  config.host    = 'errbit.ubrand.vn'
  config.port    = 80
  config.secure  = config.port == 443
end