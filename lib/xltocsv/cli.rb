require 'rubygems'
require 'trollop'
require 'iconv'
require 'roo'
require 'fileutils'
require 'tempfile'


module Xltocsv
  class Cli

    # Parse the input JSON files. We expect symbol for keys throughout
    def self.execute(args)
      Cli.new(parse_args(args)).convert
    end

    def initialize(opts)
      @opts = opts
    end

    def convert
      ## change excel files to csv
      if @opts[:format] == "all"
        xls(@opts[:input],@opts[:output],@opts[:filter])
        xlsx(@opts[:input],@opts[:output],@opts[:filter])
      elsif @opts[:format] == "xls"
        xls(@opts[:input],@opts[:output],@opts[:filter])
      elsif @opts[:format] == "xlsx"
        xlsx(@opts[:input],@opts[:output],@opts[:filter])
      else
        raise ArgumentError, 'unrecognized format'
      end
    end

    def xls(input, output, regex)
      Dir["#{input}/*.xls"].each do |file|  
        file_path = "#{file}"
        file_basename = File.basename(file, ".xlsx")
        xls = Roo::Excelx.new(file_path)
        xls.to_csv("#{output}/#{file_basename}.csv")
        unless regex.nil?
          filter("#{output}/#{file_basename}.csv")
        end
      end
    end

    def xlsx(input, output, regex)
      Dir["#{input}/*.xlsx"].each do |file|  
        file_path = "#{file}"
        file_basename = File.basename(file, ".xlsx")
        xls = Roo::Excelx.new(file_path)
        xls.to_csv("#{output}/#{file_basename}.csv")
        unless regex.nil?
          filter("#{output}/#{file_basename}.csv")
        end
      end
    end

    def filter(file)
      tmp = Tempfile.new("extract")
      open(file, 'r').each { |l| tmp << l unless l.chomp =~ /(#{@opts[:filter].join(')|(')})/ }
      tmp.close
      FileUtils.mv(tmp.path, file)
    end

    def self.parse_args(args)
      @opts = Trollop::options args do
        opt :input, "Directory containing xls/xlsx files", :type => :string, :require => true
        opt :output, "Output Directory for csv", :type => :string, :default => "/tmp"
        opt :format, "Original filetype (xls, xlsx or all)", :type => :string, :default => "all"
        opt :regex, "regular expressions to exlude lines from csv", :type => :strings
        version "xltocsv #{Xltocsv::VERSION}"
      end 
      Trollop::die :input, "This is a required path to xls/xlsx files" if opts[:input].nil? 
    end  

  end
end
