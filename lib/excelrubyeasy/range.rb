require 'uri'
require 'json'
require_relative 'rangeformat'
require_relative 'worksheet'

module ExcelRubyEasy
  module Model
	class Range 

		attr_accessor :address, 
					:addressLocal,
					:cellCount, 
					:columnCount,
					:columnIndex,
					:formulas,
					:formulasLocal,
					:numberFormat,
					:rowCount,
					:rowIndex,
					:text,
					:values,
					:valueTypes

		attr_reader :worksheetName, :format, :worksheet
		
		def initialize (parms={}) 
			@address = parms[:address]
			@addressLocal = parms[:addressLocal]
			@cellCount = parms[:cellCount]
			@columnCount = parms[:columnCount]
			@columnIndex = parms[:columnIndex]
			@formulas = parms[:formulas]
			@formulasLocal = parms[:formulasLocal]
			@numberFormat = parms[:numberFormat]
			@rowCount = parms[:rowCount]
			@rowIndex = parms[:rowIndex]
			@text = parms[:text]
			@values = parms[:values]
			@valueTypes = parms[:valueTypes]
			# Sheet name can have more than one bang. Get the correct sheet name (string before last !)
			@worksheetName = @address[0...@address.rindex('!')]
		end

		def worksheet
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')")
			parms = {		
				position: j[:position],
				name: j[:name],
				id:	j[:id],
				visibility: j[:visibility] 
			}
			return ExcelRubyEasy::Model::Worksheet.new(parms)
		end

		def format
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range(address='#{@address}')/Format")
			parms = {		
				horizontalAlignment:  j[:horizontalAlignment],
				verticalAlignment:  j[:verticalAlignment],
				wrapText: j[:wrapText],
				worksheetName: @worksheetName,
				address: @address	
			}
			return ExcelRubyEasy::Model::RangeFormat.new(parms)
		end

		def sync(isValues=false, isNumformat=false, isFormulas=false, isFormulasLocal=false )
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Worksheets('#{URI.escape @worksheetName}')/Range(address='#{@address}')")
	        parms = {}
	        parms[:values] = @values if isValues
	        parms[:numberFormat] = @values if isNumformat
	        parms[:formulas] = @values if isFormulas	        
	        parms[:formulasLocal] = @values if formulasLocal	        	        
	        request = Net::HTTP::Patch.new(uri.request_uri, ExcelRubyEasy::HEADERS_PATCH)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	        j = JSON.parse(response.body, {:symbolize_names => true}) 
			@address = j[:address]
			@addressLocal = j[:addressLocal]
			@cellCount = j[:cellCount]
			@columnCount = j[:columnCount]
			@columnIndex = j[:columnIndex]
			@formulas = j[:formulas]
			@formulasLocal = j[:formulasLocal]
			@numberFormat = j[:numberFormat]
			@rowCount = j[:rowCount]
			@rowIndex = j[:rowIndex]
			@text = j[:text]
			@values = j[:values]
			@valueTypes = j[:valueTypes]
	     	return self
		end

		def delete(shift=nil)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Worksheets('#{URI.escape @address.split('!').first}')/Range(address='#{@address}')/Delete")
	        parms = {
	        	shift: shift
	        }	        
	        request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	     	if response.kind_of? Net::HTTPSuccess
	        	true
	        else
	        	false
	        end
		end


		def clear(applyTo=nil)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Worksheets('#{URI.escape @address.split('!').first}')/Range(address='#{@address}')/Clear")
	        parms = {
	        	applyTo: applyTo
	        }	        
	        request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	     	if response.kind_of? Net::HTTPSuccess
	        	true
	        else
	        	false
	        end
		end

		def insert(shift=nil)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			uri = URI.parse(ExcelRubyEasy::Client.excelserver+"Worksheets('#{URI.escape @address.split('!').first}')/Range(address='#{@address}')/Insert")

	        parms = {
	        	shift: shift
	        }	        
	        request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
	        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	        j = JSON.parse(response.body, {:symbolize_names => true}) 
			parms = load_parms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_boundingrect(anotherRange=nil) 
			if anotherRange.is_a? String
				rangeAddress = anotherRange
			else
				rangeAddress = anotherRange.address
			end

			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range('#{rangeAddress}')/BoundingRect")
			parms = load_parms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_intersection(anotherRange=nil) 
			if anotherRange.is_a? String
				rangeAddress = anotherRange
			else
				rangeAddress = anotherRange.address
			end

			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range(address='#{rangeAddress}')/Intersection")
			parms = load_parms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_usedrange
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range(address='#{@address}')/UsedRange")
			parms = load_parms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end


		def get_cell(row=nil, column=nil) 
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range(address='#{@address}')/Cell(row=#{row.to_s},column=#{column.to_s})")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]
			parms = load_parms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_offsetrange(rowOffset=nil, coulmnOffset=nil) 
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range(address='#{@address}')/OffsetRange(rowOffset=#{rowOffset.to_s},coulmnOffset=#{coulmnOffset.to_s})")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]			
			parms = load_parms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_column(column=nil) 
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range(address='#{@address}')/Column(#{column.to_s})")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]			
			parms = load_parms(j)			
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_row(row=nil) 
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range(address='#{@address}')/Row(#{row.to_s})")
		    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
		    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]			
			parms = load_parms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end


		def get_entirerow 
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range(address='#{@address}')/EntireRow")
			parms = load_parms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_entirecolumn
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range(address='#{@address}')/EntireColumn")
			parms = load_parms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_lastcell

			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range(address='#{@address}')/LastCell")
			parms = load_parms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_lastrow

			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range(address='#{@address}')/LastRow")
			parms = load_parms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		def get_lastcolumn

			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheetName}')/Range(address='#{@address}')/LastColumn")
			parms = load_parms(j)
			return ExcelRubyEasy::Model::Range.new(parms)
		end

		private

		def load_parms(j=nil)
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