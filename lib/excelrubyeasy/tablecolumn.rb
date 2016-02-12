require 'uri'
require 'json'

module ExcelRubyEasy
  module Model
	class TableColumn

		attr_accessor :id, :name, :index, :values, :tableId

		def initialize (parms={}) 
			@id = parms[:id]
			@name = parms[:name]
			@index = parms[:index]
			@values = parms[:values]
			@tableId = parms[:tableId]
		end

		def get_range()
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{@tableId}')/Columns(#{index})/Range")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]

			parms = load_rangeparms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end	

		def update()
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Tables('#{@tableId}')/Columns('#{@id}')")
	        parms = {
				values: @values, 
	        }
	        request = Net::HTTP::Patch.new(uri.request_uri, ExcelRubyEasy::HEADERS_PATCH)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	        j = JSON.parse(response.body, {:symbolize_names => true}) 
	     	#return j["value"].
			@index = j[:index]
			@values = j[:values]
	     	return self
		end

		def delete
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Tables('#{@tableId}')/Columns('#{@id}')")
	        request = Net::HTTP::Delete.new(uri.request_uri, ExcelRubyEasy::HEADERS_DELETE)
	        response = ExcelRubyEasy::HttpAction::do_http(uri, request)	        
	     	if response.kind_of? Net::HTTPSuccess
	        	true
	        else
	        	false
	        end
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

		def get_databodyrange
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{@tableId}')/Columns('#{@id}')/DataBodyRange")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]

			parms = load_rangeparms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_headerrowrange
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{@tableId}')/Columns('#{@id}')/HeaderRowRange")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]

			parms = load_rangeparms(j)
			return ExcelRubyEasy::Model::Range.new(parms)			
		end	

		def get_totalrowrange
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{@tableId}')/Columns('#{@id}')/TotalRowRange")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]

			parms = load_rangeparms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end
		
	end
  end
end