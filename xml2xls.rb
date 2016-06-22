#! /bin/sh
exec ruby -S -x "$0" "$@"
#! ruby
# -*- coding: utf-8 -*-

require 'rubygems'
require 'optparse'
require 'csv'
require 'rexml/document'
require 'axlsx'


def process_args
  options = {}

  OptionParser.new do |opts|
    opts.banner = "Usage : ruby #{__FILE__} [options]"
    opts.version = '1.0.0'

    opts.on("-x", "--xml VALUE", "Input XML File") do |v|
      options[:xml] = File.expand_path( File.dirname(__FILE__) ) + '/' + v.to_s
    end
    opts.on("-e", "--xls [VALUE]", "Output Excel File") do |v|
      options[:xls] = File.expand_path( File.dirname(__FILE__) ) + '/' + v.to_s
    end
    opts.parse!
  end

  options
end

# options check
def check_options( options )
  if options[:xml].nil?
    puts "#{options[:xml]} not set."
    return false
  elsif !File.exists?( options[:xml] )
      puts "#{options[:xml]} not exists."
      return false
  end

  if options[:xls].nil?
    options[:xls] = File.basename( options[:xml], ".xml") + '.xlsx'
  end
  if File.exists?( options[:xls] )
    puts "#{options[:xls].to_s} exists yet."
    return false
  end

  return true
end

def setting_header
  header = Array.new
  header << "ID"
  header << "EventClass"
  header << "TextData"
  header << "Duration"
  header << "SPID"
  header << "DatabaseID"
  header << "DatabaseName"
  header << "LoginName"
  header << "CPU"
  header << "呼び出し元"

  header
end

def generate_xlsx( xlsx_name, doc, header )

end

def main
  options = process_args
  if !check_options( options )
    exit
  end

  doc = REXML::Document.new( open( options[:xml] ) )
  header = setting_header

  worksheet_name = File.basename( options[:xls], ".xlsx" )

  Axlsx::Package.new do |pkg|
    pkg.workbook.add_worksheet(name: worksheet_name ) do |sheet|
      # ヘッダーの追加
      sheet.add_row header

      id = 0
      doc.elements.each('TraceData/Events/Event[@id="12"]') do |elm|
        id         = id + 1
        event_name = elm.attributes["name"].to_s
        text_data  = elm.get_elements('./Column[@name="TextData"]').first.text.gsub(/(\r\n|\r|\n)/, " ")
        dbid       = elm.get_elements('./Column[@name="DatabaseID"]').first.text + ""
        login_name = elm.get_elements('./Column[@name="LoginName"]').first.text + ""
        duration   = elm.get_elements('./Column[@name="Duration"]').first.text + ""
        db_name    = elm.get_elements('./Column[@name="DatabaseName"]').first.text + ""
        spid       = elm.get_elements('./Column[@name="SPID"]').first.text + ""
        cpu        = elm.get_elements('./Column[@name="CPU"]').first.text + ""

        sheet.add_row [ id.to_s, event_name, text_data, duration, spid, dbid, db_name, login_name, cpu, "" ]
      end
    end
    pkg.serialize( options[:xls] )
  end

end

main 
