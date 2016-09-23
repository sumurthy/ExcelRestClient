require 'uri'
require 'json'
require 'base64'

module ExcelRubyEasy
  module Model
	class Chart 

		attr_accessor :height, :left, :name, :top, :width, :id
		attr_reader :worksheetid, :worksheetname

		def initialize (parms={}) 
			@height = parms[:height]
			@left = parms[:left]
			@name = parms[:name]
			@id = parms[:id]
			@top = parms[:top]		
			@width = parms[:width]
			@worksheetid = parms[:worksheetid]
			@worksheetname = parms[:worksheetname]
		end
		
		# Return base64 string

		def get_chart_image
	      	@name = 'Chart' + Random.rand(10000..99999).to_s if (@name.nil? || @name.empty?)		
			j = ExcelRubyEasy::HttpAction::doGetRequest_base64 (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @worksheetid}')/Charts('#{URI.escape @name}')/Image(width=0,height=0,fittingMode='fit')")			
			base64_image = j[:value]			
			return base64_image
		end

		def set_position(startcell=nil, endcell=nil) 	      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @worksheetid}')/Charts('#{URI.escape @id}')/setPosition")
	        parms = {
				startCell: startcell,
				endCell: endcell 
	        }
	        request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	     	if response.kind_of? Net::HTTPSuccess
	        	true
	        else
	        	false
	        end
		end

		def set_source(sourcerange=nil, seriesby=nil) 	      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @worksheetid}')/Charts('#{URI.escape @id}')/setData")
	        parms = {
				sourceData: sourcerange,
	        }
	        if !seriesby.nil? && !seriesby.empty?
				parms[:seriesBy] = seriesby 
	        end

	        request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	     	if response.kind_of? Net::HTTPSuccess
	        	true
	        else
	        	false
	        end
		end
		def set_position(startcell=nil, endcell=nil) 	      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @worksheetid}')/Charts('#{URI.escape @id}')/setPosition")
	        parms = {
				startCell: startcell
	        }
	        if !endcell.nil? && !endcell.empty?
				parms[:endCell] = endcell 
	        end

	        request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	     	if response.kind_of? Net::HTTPSuccess
	        	true
	        else
	        	false
	        end
		end

		def update()
	      	@name = 'Chart' + Random.rand(10000..99999).to_s if (@name.nil? || @name.empty?)		
			uri = URI.parse(ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @worksheetid}')/Charts('#{URI.escape @id}')")

	        parms = {
				height: @height,
				left: @left, 
				name: @name,
				top: @top, 
				width: @width, 
	        }
	        request = Net::HTTP::Patch.new(uri.request_uri, ExcelRubyEasy::HEADERS_PATCH)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	        j = JSON.parse(response.body, {:symbolize_names => true}) 
	     	#return j["value"].
			@height = j[:height]
			@left = j[:left] 
			@name = j[:name]
			@top = j[:top]
			@width = j[:width ]
	     	return self
		end

		def delete
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @worksheetid}')/Charts('#{URI.escape @id}')")
	        request = Net::HTTP::Delete.new(uri.request_uri, ExcelRubyEasy::HEADERS_DELETE)
	        response = ExcelRubyEasy::HttpAction::do_http(uri, request)
	     	if response.kind_of? Net::HTTPSuccess
	        	true
	        else
	        	false
	        end
		end

	end
  end
end