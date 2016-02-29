require 'net/https'
require 'uri'
require 'json'
require_relative 'worksheet.rb'
require_relative 'chart.rb'
require_relative 'table.rb'
require_relative 'range.rb'
require_relative 'nameditem.rb'
require_relative 'errors.rb'
require_relative 'httpaction.rb'


module ExcelRubyEasy

class Client

	class << self
	  attr_accessor :session_id, :persist_changes, :excelfileid
	end
	
	attr_accessor :client_id, :secret, :redirect_uri, :authcode, :access_token, :refresh_token, :auth_code, :auth_url, :resource
				  

	
	include ExcelRubyEasy::Logging

	def initialize(parms={})
		@client_id = parms[:client_id] || nil
		@secret = parms[:secret] || nil	
		@redirect_uri = parms[:redirect_uri] || nil
		@auth_url = parms[:auth_url] || nil
		@resource = parms[:resource] || nil

		if parms[:persist_changes]
		 	ExcelRubyEasy::Client.persist_changes = true
		 else
		 	ExcelRubyEasy::Client.persist_changes = false
		end
	end

	##
    # Returns the authorization URL that the application use to re-direct the user for 
    # login and authorization purpose.
    #
    # @return Authozation URL
	def get_authurl
		logger.debug "D, #{__method__.to_s}"
		params = {
            "client_id" => @client_id,
            "response_type" => "code",
            "redirect_uri" => @redirect_uri,
            "prompt" => "consent"
        }
        auth_uri = URI::Generic.new("https", nil, @auth_url, nil, nil, "authorize", 
        							 nil, nil, nil)
        auth_uri.query = URI.encode_www_form(params)
        logger.debug "D, #{__method__.to_s},  #{auth_uri.to_s}"
        return auth_uri.to_s
	end

	##
	# Return Access token for a given authorization code. 
	# If session is initialized with an auth-code, then it'll 
	# used to return the access token
	#
	def set_access_token(authcodeParam = nil, update_refToken = false) 
		logger.debug "D, #{__method__.to_s}, resource= #{resource}, authcode passed = #{authcodeParam} "

		puts "D, #{__method__.to_s}, url = #{auth_url}, resource= #{resource}, authcode passed = #{authcodeParam} "

      	auth_code = authcodeParam unless authcodeParam.nil?  
		uri = URI.parse("https://#{auth_url}token")
        request = Net::HTTP::Post.new(uri.request_uri)
        resource = resource 
        params = {
            "grant_type" => "authorization_code",
            "client_id" => @client_id,
            #"client_secret" => CGI.escape(@client_secret),
            "client_secret" => @secret,
            "code" => auth_code,
            "redirect_uri" => @redirect_uri,
            "resource" => @resource
        }
        # request.set_form_data(params)
        logger.debug "D, 1) calling auth token "	
        response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, params)
        logger.debug "D, 2) ending auth token "	
        j = JSON.parse(response.body) 
        @access_token = j['access_token']
        ExcelRubyEasy::TOKEN["token"] = "Bearer " + @access_token
		ExcelRubyEasy::HEADERS_POST_BASIC["Authorization"] = ExcelRubyEasy::TOKEN["token"]        
		ExcelRubyEasy::HEADERS_POST["Authorization"] = ExcelRubyEasy::TOKEN["token"]        
		ExcelRubyEasy::HEADERS_GET["Authorization"] = ExcelRubyEasy::TOKEN["token"]        
		ExcelRubyEasy::HEADERS_GET_ALL["Authorization"] = ExcelRubyEasy::TOKEN["token"]        
		ExcelRubyEasy::HEADERS_PATCH["Authorization"] = ExcelRubyEasy::TOKEN["token"]        
		ExcelRubyEasy::HEADERS_DELETE["Authorization"] = ExcelRubyEasy::TOKEN["token"]        

        if update_refToken
            @refresh_token = j['refresh_token']
        end
        logger.debug "D,Returning "

     	return 
    end

	def search_for_excelfiles
		ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
		j = ExcelRubyEasy::HttpAction::doGetRequest (ONEDRIVE_SEARCH)
	    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
	    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]			
		response_array = j[:value]
		return_array = []
		parms = {}
		response_array.each do |item|
			parms = {		
				name: item[:name],
				id: item[:id],
				size: item[:size]
			}
			return_array << parms
		end
		return return_array
	end

	def self.excelserver
		return ExcelRubyEasy::EXCEL_BASE + ExcelRubyEasy::RESOURCE_PATH["path"]
	end

	def set_excelfile_and_session(id=nil)
		ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
		ExcelRubyEasy::Client.excelfileid  = id
		ExcelRubyEasy::RESOURCE_PATH["path"] =	id + '/workbook/'
	 	ExcelRubyEasy::Client.session_id = ExcelRubyEasy::Client.create_sessionid(ExcelRubyEasy::Client.persist_changes)			 	
		puts "------------"
		puts "#{ExcelRubyEasy::RESOURCE_PATH["path"]}"
		puts "------------"
		return	
	end

	def self.create_sessionid(persistChanges=nil)

		return
		persistChanges = ExcelRubyEasy::Client.persist_changes

		uri = URI.parse(ExcelRubyEasy::Client.excelserver + 'CreateSession')

		parms = {
			persistChanges: persistChanges
		}
		
		request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST_BASIC)
		response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json, false)
		j = JSON.parse(response.body, {:symbolize_names => true}) 
		ExcelRubyEasy::HEADERS_POST["Workbook-Session-Id"] = j[:id]
		ExcelRubyEasy::HEADERS_PATCH["Workbook-Session-Id"] = j[:id]
		ExcelRubyEasy::HEADERS_DELETE["Workbook-Session-Id"] = j[:id]
		ExcelRubyEasy::HEADERS_GET["Workbook-Session-Id"] = j[:id]
		ExcelRubyEasy::HEADERS_GET_ALL["Workbook-Session-Id"] = j[:id]		

		ExcelRubyEasy::Client.session_id = j[:id] 
		#logger.debug "Returning "
	end

	def refresh_session
		uri = URI.parse(ExcelRubyEasy::Client.excelserver+'RefreshSession')
		parms = {
			type: type
		}
		request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
		response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
	    if !response.kind_of? Net::HTTPSuccess
	        raise ExcelRubyEasy::ClientError.new
	    end		
		return 
	end

	def close_session
		uri = URI.parse(ExcelRubyEasy::Client.excelserver+'CloseSession')
		parms = {
			type: type
		}
		request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
		response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)
		return 		
	end


	##
	# Add Table
	# to do: also accept the Range object. 

	def add_table(rangeAddress=nil, hasHeaders=true) 
		ExcelRubyEasy::logger.debug "D, #{__method__.to_s}, rangeAddress = #{rangeAddress} "

		uri = URI.parse(ExcelRubyEasy::Client.excelserver+'Tables/$/Add')
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


	##
	# Add Worksheet
	#

	def add_worksheet(name=nil)
		ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		
		name = 'Sheet' + Random.rand(10000..99999).to_s if (name.nil? || name.empty?)		
		uri = URI.parse(ExcelRubyEasy::Client.excelserver+'Worksheets')
		parms = {
			name: name
		}
		request = Net::HTTP::Post.new(uri.request_uri, ExcelRubyEasy::HEADERS_POST)
		response = ExcelRubyEasy::HttpAction::http_sync_with_body(uri, request, parms.to_json)

	    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
	    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]

		j = JSON.parse(response.body, {:symbolize_names => true}) 
		parms={}
		parms = {
			position: j[:position],
			name: j[:name],
			id: j[:id],
			visibility: j[:visibility]
		}
		return ExcelRubyEasy::Model::Worksheet.new(parms)		
	end

	##
	# List Objects
	#
	def list_objects(type=nil)
		ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"
		response = []
		j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + type)
	    REQUEST_RESPONSE_SAVE_LIST["req"] = REQUEST_RESPONSE["req"]
	    REQUEST_RESPONSE_SAVE_LIST["res"] = REQUEST_RESPONSE["res"]
	    REQUEST_RESPONSE_SAVE["req"] = REQUEST_RESPONSE["req"]
	    REQUEST_RESPONSE_SAVE["res"] = REQUEST_RESPONSE["res"]

		response_array = j[:value]
		return_array = []
		parms={}
		case type.split('/').last.capitalize
		when 'Worksheets'
			response_array.each do |item|
				parms = {		
					position: item[:position],
					name: item[:name],
					id:	item[:id],
					visibility: item[:visibility] 
				}
				return_array << ExcelRubyEasy::Model::Worksheet.new(parms)
			end
		when 'Charts'
			response_array.each do |item|
				parms = {		
					height: item[:height],
					left: item[:left],
					name: item[:name],
					top: item[:top],   
					width: item[:width]  
				}
				return_array << ExcelRubyEasy::Model::Chart.new(parms)
			end			
		when 'Tables'
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

		when 'Names'		
			response_array.each do |item|
				parms = {
					name: item[:name], 
					type: item[:type], 
					value: item[:value],
					visible: item[:visible],					
				}
				return_array << ExcelRubyEasy::Model::NamedItem.new(parms)
			end		 
		
		end
		return return_array
	end

	def get_range(name=nil, worksheet=nil, address=nil)
		ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 

		if (worksheet.nil? || worksheet.empty?)	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Names('#{URI.escape name}')/Range")										
		else
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape worksheet}')/Range(address='#{address}')")
		end
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
		puts "Creating0"
		return ExcelRubyEasy::Model::Range.new(parms)
	end	 

	def list_workbook_charts(worksheets=[])
		return_array = []
		worksheets.each do |ws|
			puts
			puts "Reading chart for #{ws.name}" 
			puts
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape ws.id}')/Charts")
			response_array = j[:value]
			parms={}
			response_array.each do |item|
				parms = {		
					height: item[:height],
					left: item[:left],
					name: item[:name],
					top: item[:top],   
					width: item[:width],
					id: item[:id], 
					worksheetid: ws.id,
					worksheetname: ws.name  					  
				}
				return_array << ExcelRubyEasy::Model::Chart.new(parms)
			end
		end
		puts '*********************'
		puts return_array
		puts '*********************'		
		return return_array
	end

	def get_worksheet(id=nil, name=nil)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			if !id.nil?
				j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape id}')")
			else
				j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Worksheets('#{URI.escape name}')")				
			end
			parms = {		
				position: j[:position],
				name: j[:name],
				id:	j[:id],
				visibility: j[:visibility] 
			}
			return ExcelRubyEasy::Model::Worksheet.new(parms)
	end

	def get_nameditem(name=nil)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"      		      	
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Names('#{URI.escape name}')")
			parms = {
				name: j[:name], 
				type: j[:type], 
				value: j[:value],
				visible: j[:visible],					
			}
			return ExcelRubyEasy::Model::NamedItem.new(parms)
	end

	def get_table(id=nil, name=nil)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"  

			if !id.nil?
				j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{id}')")
			else
				j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Tables('#{URI.escape name}')")				
			end
			parms = {
				id: j[:id],
				name: j[:name], 
				showHeaders: j[:showHeaders],
				showTotals: j[:showTotals], 
				style: j[:style], 
			}			
			return ExcelRubyEasy::Model::Table.new(parms)
	end

	
end
end