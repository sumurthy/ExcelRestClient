require 'uri'
require 'json'

module ExcelRubyEasy
  module Model
	class RangeFont

		attr_accessor :bold, :color, :italic, :name, :size,	:underline
		attr_reader :worksheetName, :address

		def initialize (parms={}) 
			@bold = parms[:bold]
			@color = parms[:color]
			@italic = parms[:italic]
			@name = parms[:name]
			@size = parms[:size]
			@underline = parms[:underline]
			@worksheetName = parms[:worksheetName]
			@address = parms[:address]		
		end

		def update
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Worksheets('#{URI.escape @worksheetName}')/Range('#{address}')/Format/Font")
	        parms = {
				bold: @bold, 
				color: @color, 
				italic: @italic,
				name: @name,
				size: @size,
				underline: @underline 
			}
	        # Remove empty or null values
	        parms.delete_if {|key, value| value.nil? || value.to_s.empty? }
			
	        request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_PATCH)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	        j = JSON.parse(response.body, {:symbolize_names => true}) 
			@bold = j[:bold]
			@color = j[:color]
			@italic = j[:italic]
			@name = j[:name]
			@size = j[:size]
			@underline = j[:underline]				       
	     	return self
		end	
	end
  end
end