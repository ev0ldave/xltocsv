$:.push File.expand_path("../lib",__FILE__)
require 'xltocsv'
require 'xltocsv/version'

Gem::Specification.new do |s|
  s.name        = 'xltocsv'
  s.version     = '0.0.1'
  s.date        = '2013-09-12'
  s.summary     = "digery do"
  s.description = "A simple excel to csv converter"
  s.authors     = ["David Andrew"]
  s.email       = ['ev0ldave@gmail.com']
  s.license       = 'MIT'

  s.executables = %w(xltocsv)

  s.files += Dir.glob("lib/**/*")
  s.files += Dir.glob("bin/**/*")

  s.add_dependency('trollop')
  s.add_dependency('roo')

end
