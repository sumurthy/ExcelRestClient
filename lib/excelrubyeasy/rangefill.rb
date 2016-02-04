require 'uri'
require 'json'

module ExcelRubyEasy
  module Model
	class RangeFill

		attr_accessor :color
		attr_reader :worksheetName, :address

		def initialize (parms={}) 
			@color = parms[:color]
			@worksheetName = parms[:worksheetName]
			@address = parms[:address]			
		end

		def update
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Worksheets('#{URI.escape @worksheetName}')/Range('#{address}')/Format/Fill")
	        parms = {
				color: @color
	        }
	        # Remove empty or null values
	        parms.delete_if {|key, value| value.nil? || value.to_s.empty? }
	        
	        request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_PATCH)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	        j = JSON.parse(response.body, {:symbolize_names => true}) 
			@color = j[:color]
	     	return self
		end

		def clear
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Worksheets('#{URI.escape @worksheetName}')/Range('#{address}')/Format/Fill/Clear")
	        request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
	        response = ExcelRubyEasy::HttpAction::do_http(uri, request, parms.to_json)
	        j = JSON.parse(response.body, {:symbolize_names => true}) 
			@color = j[:color]
	     	return self
		end		
	end
  end
end