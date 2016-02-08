require 'uri'
require 'json'

module ExcelRubyEasy
  module Model
	class Worksheet 

		attr_accessor :position, :name
		attr_reader :id, :visibility, :charts, :tables

		def initialize (parms={}) 
			@position = parms[:position]
			@name = parms[:name]
			@id = parms[:id]
			@visibility = parms[:visibility]		
		end

		def update()
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Worksheets('#{URI.escape @id}')")
			puts "$$$$$$$$$$ name inside class: #{@name}"
	        parms = {
	            name: @name,
	            position: @position
	        }
	        request = Net::HTTP::Patch.new(uri.request_uri, ExcelRubyEasy::HEADERS_PATCH)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	        j = JSON.parse(response.body, {:symbolize_names => true}) 
	     	#return j["value"].
			@position = j[:position]
			@name = j[:name]
			@id = j[:id]
			@visibility = j[:visibility]	     	
	     	return self
		end

		def delete
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Worksheets('#{URI.escape @id}')")
	        request = Net::HTTP::Delete.new(uri.request_uri, ExcelRubyEasy::HEADERS_DELETE)
	        response = ExcelRubyEasy::HttpAction::do_http(uri, request)
	     	if response.kind_of? Net::HTTPSuccess
	        	true
	        else
	        	false
	        end
		end

		def get_usedrange (valuesOnly = true)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @id}')/UsedRange(valuesOnly=#{valuesOnly})")
			parms = load_rangeparms(j)
			return ExcelRubyEasy::Model::Range.new(parms)		
		end

		def get_cell (row=nil, column=nil)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @id}')/Cell(row=#{row.to_s},column=#{column.to_s})")
			parms = load_rangeparms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_range(address=nil)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @id}')/Range(address='#{address}')")
			parms = load_rangeparms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end	 

		def get_chart(id=nil, name=nil)
			if !id.nil?
				j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver +  "Worksheets('#{URI.escape @id}')/Charts('#{URI.escape id}')")
			else
				j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver +  "Worksheets('#{URI.escape @id}')/Charts(#{name})")
			end
			parms = {		
				height: j[:height],
				left: j[:left],
				name: j[:name],
				top: j[:top],   
				width: j[:width],
				id: j[:id],
				worksheetid: @id
			}			
			return ExcelRubyEasy::Model::Chart.new(parms)
		end

		def charts
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @id}')/Charts")
			response_array = j[:value]
			return_array = []
			parms={}
			response_array.each do |item|
				parms = {		
					height: item[:height],
					left: item[:left],
					name: item[:name],
					top: item[:top],   
					width: item[:width],
					id: item[:id], 
					worksheetid: @id  					  
				}
				return_array << ExcelRubyEasy::Model::Chart.new(parms)
			end
			return return_array
		end

		def tables
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @id}')/Tables")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]
		    REQUEST_RESPONSE_SAVE_LIST["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE_LIST["res"] = REQUEST_RESPONSE["res"]		    
			response_array = j[:value]
			return_array = []
			parms={}
			response_array.each do |item|
				parms = {
					id: item[:id],
					name: item[:name], 
					showHeaders: item[:showHeaders],
					showTotals: item[:showTotals], 
					style: item[:style], 
				}
				return_array << ExcelRubyEasy::Model::Table.new(parms)
			end
			return return_array
		end

		def add_table(rangeAddress=nil, hasHeaders=true) 

			#rangeAddress = (@name + '!' + rangeAddress) unless rangeAddress.include?('!')

			uri = URI.parse(ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @id}')/Tables/$/Add")
			parms = {
				address: rangeAddress,
				hasHeaders: hasHeaders
			}
			request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
			response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]
			j = JSON.parse(response.body, {:symbolize_names => true}) 
			parms={}
			parms = {
				id: j[:id],
				name: j[:name], 
				showHeaders: j[:showHeaders],
				showTotals: j[:showTotals], 
				style: j[:style]
			}
			return ExcelRubyEasy::Model::Table.new(parms)  
		end

		def add_chart(type=nil, sourcedata=nil, seriesby=nil) 

			uri = URI.parse(ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape @id}')/Charts/$/Add")
			parms = {
				type: type,
				sourcedata: sourcedata,
				seriesby: seriesby
			}
			request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
			response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]
			j = JSON.parse(response.body, {:symbolize_names => true}) 
			parms={}
			parms = {
				id: j[:id],
				height: j[:height],
				left: j[:left], 
				name: j[:name],
				top: j[:top], 
				width: j[:width],
				worksheetid: @id
			}
			return ExcelRubyEasy::Model::Chart.new(parms)  
		end


		private def load_rangeparms(j=nil)
			parms = {}
			parms = {
				address: j[:address],
				addressLocal: j[:addressLocal],
				cellCount: j[ :cellCount],
				columnCount: j[:columnCount],
				columnIndex: j[:columnIndex],
				formulas: j[:formulas],
				formulasLocal: j[:formulasLocal],
				numberFormat: j[:numberFormat],
				rowCount: j[:rowCount],
				rowIndex: j[:rowIndex],
				text: j[:text],
				values: j[:values],
				valueTypes: j[:valueTypes]
			}
			return parms
		end



	end
  end
end

	