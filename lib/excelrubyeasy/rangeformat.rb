require 'uri'
require 'json'

require_relative 'rangefont'
require_relative 'rangefill'
require_relative 'rangeborder'

module ExcelRubyEasy
  module Model
	class RangeFormat 

		attr_accessor :horizontalAlignment, :verticalAlignment, :wrapText
		attr_reader :borders, :fill, :font, :worksheetName, :address


		def initialize (parms={}) 
			@horizontalAlignment = parms[:horizontalAlignment]
			@verticalAlignment = parms[:verticalAlignment]
			@wrapText = parms[:wrapText]
			@worksheetName = parms[:worksheetName]
			@address = parms[:address]			
		end

		def update()
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Worksheets('#{URI.escape @worksheetName}')/Range('#{address}')/Format")
	        parms = {
				horizontalAlignment: @horizontalAlignment,
				verticalAlignment: @verticalAlignment,
				wrapText:  @wrapText
	        }
	        puts "format before: #{parms}"
	        # Remove empty or null values
	        parms.delete_if {|key, value| value.nil? || value.to_s.empty? }
	        puts "format after: #{parms}"			        
	        
	        request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_PATCH)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	        j = JSON.parse(response.body, {:symbolize_names => true}) 

			@horizontalAlignment = j[:horizontalAlignment]
			@verticalAlignment = j[:verticalAlignment]
			@wrapText = j[:wrapText]		
	     	return self
		end

		def fill()
			puts "$$$$$$$$$$$$$$$$$$$$$ Calling: #{@worksheetName}, #{@address}"
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @worksheetName}')/Range('#{address}')/Format/Fill")
			parms = {
				color: j[:color],
				worksheetName: @worksheetName,
				address: @address
			}
			return ExcelRubyEasy::Model::RangeFill.new(parms)			
		end

		def font()
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @worksheetName}')/Range('#{address}')/Format/Font")
			parms = {
				bold: j[:bold],
				color: j[:color],
				italic: j[:italic],
				name: j[:name],
				size: j[:size],
				underline: j[:underline],
				worksheetName: @worksheetName,
				address: @address
			}
			return ExcelRubyEasy::Model::RangeFont.new(parms)			
		end

		def borders()
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @worksheetName}')/Range('#{address}')/Format/Borders")
			response_array = j[:value]
			return_array = []
			response_array.each do |item|
				parms={}
				parms = {
					color: item[:color],
					sideIndex: item[:sideIndex],
					style: item[:style],
					weight: item[:weight],
					worksheetName: @worksheetName,
					address: @address
				}
				return_array << ExcelRubyEasy::Model::RangeBorder.new(parms)
			end
			return return_array
		end

	end
  end
end