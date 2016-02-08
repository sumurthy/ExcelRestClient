require 'uri'
require 'json'
require_relative 'tablerow'
require_relative 'tablecolumn'

module ExcelRubyEasy
  module Model
	class Table 

		attr_accessor :id, :name, :showHeaders, :showTotals, :style
		attr_reader :rows, :columns

		def initialize (parms={}) 
			@id = parms[:id]
			@name = parms[:name]
			@showHeaders = parms[:showHeaders]
			@showTotals = parms[:showTotals]		
			@style = parms[:style]
			puts "calling with #{@id}"
		end

		def update()
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Tables('#{@id}')")
	        parms = {
				name: @name, 
				showHeaders: @showHeaders,
				showTotals: @showTotals, 
				style: @style, 
	        }
	        if parms[:name].nil? || parms[:name].empty?
	        	parms.delete(:name)
	        end
	        if parms[:style].nil? || parms[:style].empty?
	        	parms.delete(:style)
	        end
	        request = Net::HTTP::Patch.new(uri.request_uri, ExcelRubyEasy::HEADERS_PATCH)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	        j = JSON.parse(response.body, {:symbolize_names => true}) 
	     	#return j["value"].
			@id = j[:id]
			@name = j[:name]
			@showHeaders = j[:showHeaders]
			@showTotals = j[:showTotals]
			@style = j[:style]
	     	return self
		end

		def delete
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Tables('#{@id}')")       
	        request = Net::HTTP::Delete.new(uri.request_uri, ExcelRubyEasy::HEADERS_DELETE)
	        response = ExcelRubyEasy::HttpAction::do_http(uri, request)
	     	if response.kind_of? Net::HTTPSuccess
	        	true
	        else
	        	false
	        end
		end

		def convert_to_range 
			puts "convert to range2a"
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s} "
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Tables('#{@id}')/convertToRange")
			request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
			response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, nil)
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]
	     	if response.kind_of? Net::HTTPSuccess
	        	true
	        else
	        	false
	        end
		end

		def add_row(values=nil, index=nil) 
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}, rangeAddress = #{values.to_s} "
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Tables('#{@id}')/Rows")
			parms = {
				"values" => values,
				"index" => index
			}
			request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
			response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]

			j = JSON.parse(response.body, {:symbolize_names => true}) 
			parms = {
				index: j[:index],
				values: j[:values],
				tableId: @id
			}
			return ExcelRubyEasy::Model::TableRow.new(parms)
		end

		def rows 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{@id}')/Rows")
		    REQUEST_RESPONSE_SAVE_LIST["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE_LIST["res"] = REQUEST_RESPONSE["res"]

			response_array = j[:value]
			return_array = []
			response_array.each do |item|
				parms={}
				parms = {
					index: item[:index],
					values: item[:values],
					tableId: @id
				}
				return_array << ExcelRubyEasy::Model::TableRow.new(parms)
			end
			return return_array
		end

		def add_column(values=nil, index=nil) 
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}, rangeAddress = #{values.to_s} "
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Tables('#{@id}')/Columns")
			parms = {
				"values" => values,
				"index" => index
			}
			request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
			response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]

			j = JSON.parse(response.body, {:symbolize_names => true}) 
			parms = {
				id: j[:id],
				name: j[:name],
				index: j[:index],
				values: j[:values],
				tableId: @id
			}
			return ExcelRubyEasy::Model::TableColumn.new(parms)
		end

		def columns 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{@id}')/Columns")
		    REQUEST_RESPONSE_SAVE_LIST["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE_LIST["res"] = REQUEST_RESPONSE["res"]

			response_array = j[:value]
			return_array = []
			response_array.each do |item|
				parms={}
				parms = {
					id: item[:id],
					name: item[:name],
					index: item[:index],
					values: item[:values],
					tableId: @id					
				}
				return_array << ExcelRubyEasy::Model::TableColumn.new(parms)
			end
			return return_array
		end

		def get_row(index=nil)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{@id}')/Rows/$/ItemAt(index=#{index})")
			parms = {
				index: j[:index],
				values: j[:values],
				tableId: @id
			}
			return ExcelRubyEasy::Model::TableRow.new(parms)			
		end

		def get_column(id=nil)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{@id}')/Columns(#{id})")
			parms = {
				id: j[:id],
				name: j[:name],
				index: j[:index],
				values: j[:values],
				tableId: @id
			}
			return ExcelRubyEasy::Model::TableColumn.new(parms)			
		end

		def get_databodyrange
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{@id}')/DataBodyRange")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]			
			parms = load_rangeparms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_headerrowrange
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{@id}')/HeaderRowRange")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]			
			parms = load_rangeparms(j)
			return ExcelRubyEasy::Model::Range.new(parms)			
		end	

		def get_totalrowrange
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{@id}')/TotalRowRange")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]			
			parms = load_rangeparms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_range(address=nil)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{@id}')/Range")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]			
			parms = load_rangeparms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
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