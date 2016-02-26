require 'net/https'
require 'uri'
require 'json'
require_relative 'errors'
require 'logger'

module ExcelRubyEasy
	class HttpAction

		# Clean-up the parameters convert them to Strings. 
	    def self.clean_params(params)
	        r = {}
	        params.each do |k,v|
	            r[k] = v.to_s if not v.nil?
	        end
	        r
	    end

		include ExcelRubyEasy::Logging

			##
		# Generic GET Request initiator
		#
		def self.doGetRequest(geturl, useJsonFormat=true,isRetry=true)
	        ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		
	        uri = URI.parse(geturl)
	        request = Net::HTTP::Get.new(uri, ExcelRubyEasy::HEADERS_GET)

	        response = ExcelRubyEasy::HttpAction::do_http(uri, request,isRetry)
	        if useJsonFormat
	            return JSON.parse(response.body, {:symbolize_names => true})
	        else
	            return response.body
	        end
	    end

		def self.doGetRequest_base64(geturl, isRetry=true)
	        ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		
	        uri = URI.parse(geturl)
	        request = Net::HTTP::Get.new(uri, ExcelRubyEasy::HEADERS_GET_ALL)
	        response = ExcelRubyEasy::HttpAction::do_http(uri, request,isRetry)
	        return JSON.parse(response.body, {:symbolize_names => true})
	        
            return response.body
	    end


		def self.do_http(uri, request,isRetry=true) 
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"
	        http = Net::HTTP.new(uri.host, uri.port)
	        http.use_ssl = true
	        http.verify_mode = OpenSSL::SSL::VERIFY_NONE
	#        http.verify_mode = OpenSSL::SSL::VERIFY_PEER
	#        http.ca_file = ExcelAPI::TRUSTED_CERT_FILE
	        #http.set_debug_output(Logger.new("foo.txt") ) 
	        http.set_debug_output($stdout)    
	        response = http.request(request)

	        ExcelRubyEasy::REQUEST_RESPONSE["req"] = request
			ExcelRubyEasy::REQUEST_RESPONSE["res"] = response                                  

	        # New error handling
	        #ExcelRubyEasy::logger.debug(response.to_hash)

	        if response.kind_of? Net::HTTPSuccess
	        	response
	        else	        	
	   #      	if response.kind_of?(Net::HTTPNotFound) && isRetry
				# 	#request["Workbook-Session-id"] = ExcelRubyEasy::Client.create_sessionid(ExcelRubyEasy::Client.persist_changes)
				# 	response = http.request(request)
				#     ExcelRubyEasy::REQUEST_RESPONSE["req"] = request
				# 	ExcelRubyEasy::REQUEST_RESPONSE["res"] = response 					
				# 	if response.kind_of? Net::HTTPSuccess                                 
	   #      			response	
	   #      		else	        	
	   #      			check_status(response.code.to_i, response.body) 
	   #      		end
				# else
	        		check_status(response.code.to_i, response.body) 
	        	# end
	        end

	    end

	    def self.http_sync_with_body(uri, request, body,isRetry=true)
	    	ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"
	        
	        if body != nil
	            if body.is_a?(Hash)
	                request.set_form_data(HttpAction::clean_params(body))
	            elsif body.respond_to?(:read)
	                if body.respond_to?(:length)
	                    request["Content-Length"] = body.length.to_s
	                elsif body.respond_to?(:stat) && body.stat.respond_to?(:size)
	                    request["Content-Length"] = body.stat.size.to_s
	                else
	                	ExcelRubyEasy::logger.fatal "F, Error while processing the HTTP body"
	                    raise ArgumentError, "Don't know how to handle 'body' (responds to 'read' but not to 'length' or 'stat.size')."
	                end
	                request.body_stream = body
	            else
	                s = body.to_s
	                request["Content-Length"] = s.length
	                request.body = s
	            end
	        end	        
  	    	ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"

	        do_http(uri, request, isRetry)
	    end

	    ## 
	    # HTTP Error validation
	    #

	    def self.check_status(status, body = nil, header=nil, message = nil)

	        case status
	        when 200...300
	          nil
	        when 301, 302, 303, 307
	          message ||= 'Warning: The Excel file is open in edit mode elsewhere. You may experience issues. HTTP code: ' + status.to_s
	          raise ExcelRubyEasy::RedirectError.new(message, status_code: status, header: header, body: body)
	        when 401
	          message ||= 'Unauthorized, HTTP code: ' + status.to_s
	          raise ExcelRubyEasy::AuthorizationError.new(message, status_code: status, header: header, body: body)
	        when 304, 400, 402...500
	          message ||= 'Invalid request, HTTP code: ' + status.to_s
	          raise ExcelRubyEasy::ClientError.new(message, status_code: status, header: header, body: body)
	        when 500...600
	          message ||= 'Server error, HTTP code: ' + status.to_s
	          raise ExcelRubyEasy::ServerError.new(message, status_code: status, header: header, body: body)
	        else
	          ExcelRubyEasy::logger.warn("Encountered unexpected status code #{status.to_s}")
	          message ||= 'Unknown error, HTTP code: ' + status.to_s
	          raise ExcelRubyEasy::TransmissionError.new(message, status_code: status, header: header, body: body)
	        end
	    end
	end
end