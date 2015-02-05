Airbrake.configure do |config|
  config.api_key = 'd51d9523944e4f095128c72f041502ae'
  config.host    = 'err.viphat.name'
  config.port    = 80
  config.secure  = config.port == 443
end