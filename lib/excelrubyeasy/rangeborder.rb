require 'uri'
require 'json'

module ExcelRubyEasy
  module Model
	class RangeBorder

		attr_accessor :color, :sideIndex, :style, :weight				
		attr_reader :worksheetName, :address

		def initialize (parms={}) 
			@color = parms[:color]
			@sideIndex = parms[:sideIndex]
			@style = parms[:style]
			@weight = parms[:weight]					
			@worksheetName = parms[:worksheetName]
			@address = parms[:address]		
		end

		def update
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Worksheets('#{URI.escape @name}')/Range('#{address}')/Format/Fill")
	        parms = {
				color: @color, 
				sideIndex: @sideIndex, 
				style: @style, 
				weight: @weight, 				
	        }
	        # Remove empty or null values
	        parms.delete_if {|key, value| value.nil? || value.to_s.empty? }
	        
	        request = Net::HTTP::Patch.new(uri.request_uri, ExcelRubyEasy::HEADERS_PATCH)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	        j = JSON.parse(response.body, {:symbolize_names => true}) 
			@color = j[:color]
			@sideIndex = j[:sideIndex]
			@style = j[:style]
			@weight = j[:weight] 					       
	     	return self
	     	
		end	
	end
  end
end