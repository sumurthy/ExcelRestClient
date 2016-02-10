require_relative '../lib/excelrubyeasy'
require 'logger'
require 'sinatra'
require 'uri'
require 'json'

logger = Logger.new(STDOUT)
logger.level = Logger::DEBUG  

SESSION = {
        client_id: "dad4b481-a6b7-4cfd-9117-32eed770d4b1",
        redirect_uri: "http://excelrest1.cloudapp.net/signon",
        secret: "AGCifRyMSOMNQr5n36Kb9Pzh0U4oR8cfKQwjXv39ip0=",
        auth_url: "login.windows.net/common/oauth2/",
        resource: "https://graph.microsoft.com",
        persist_changes: true
}



client = ExcelRubyEasy::Client.new (SESSION)
auth_url = client.get_authurl

##
# Main entry point for the application
#

set :bind, '0.0.0.0'
set :sessions, true
set :session_secret, 'Q1W2E3R4'

get '/' do 
	logger.debug "T inside Main Page"
	erb :index, :locals => {:status => params[:status], :htrace => "none"} 
end

get '/go' do 
	logger.debug "T inside Main Page"
	erb :go, :locals => {:authrUrl => auth_url} 
end

get '/signon' do 
	puts "----------> Sign on redirect completed......."	
	client.set_access_token(params[:code], true) 
	saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE
	files_list = client.search_for_excelfiles
	erb :selectfile, :locals => {:files => files_list, :rr => saveRR } 
end

get '/setexcelfile' do 
	client.set_excelfile_and_session(params[:id]) 
	erb :index, :locals => {:status => params[:status], :htrace => "none"} 
end

get '/listworksheets' do 
	ExcelRubyEasy::logger.debug "T Inside list worksheets"
	begin
		@@sheets = client.list_objects("Worksheets")
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE
		ExcelRubyEasy::logger.debug "T #{@@sheets.to_s}"
		erb :listworksheets, :locals => {:sheets => @@sheets, :rr => saveRR } 
	rescue Exception => e 
		ExcelRubyEasy::logger.warn "ERROR!! #{e.to_s}"
		redirect "/?htrace=none&status=Error-List-Worksheets"
	end
end

get '/listtables' do 
	logger.debug "T Inside list tables"
	@@tables = client.list_objects("Tables")
	saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE_LIST
	erb :listtables, :locals => {:tables => @@tables, :status => params[:status], :rr => saveRR } 
	# out = "<center> <h3> Click <a href=#{auth_url}> here </a> to access file browser </h3></center> "
end

get '/listworksheettables' do 
	logger.debug "T Inside list tables"
	@@tables = client.get_worksheet(nil, params[:sheetname]).tables
	saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE_LIST
	erb :listtables, :locals => {:tables => @@tables, :status => params[:status], :rr => saveRR } 
	# out = "<center> <h3> Click <a href=#{auth_url}> here </a> to access file browser </h3></center> "
end

get '/listtablerows' do 
	# begin
		@@table = client.get_table(nil, params[:tablename])
		@@rows = @@table.rows
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE_LIST
		erb :listtablerows, :locals => {
			:rows => @@rows,
			:table => @@table,
			:htrace => "ok",
			:rr => saveRR
		}

	# rescue
	# 	saveRR = ExcelRubyEasy::REQUEST_RESPONSE		
	# 	erb :index, :locals => {:status => "Error-Fetching-Table-Rows: #{params[:tablename]}",
	# 							:htrace => "ok",
	# 							:rr => saveRR
	# 						   }	
	# end	

end

get '/showtablecolumnrangeinfo' do 

	begin
		saveRR = []
		@@table = client.get_table(nil, params[:tablename])
		@@tableColumn = @@table.get_column(params[:id])
		@@range = @@tableColumn.get_range

		temp = {}
		temp["req"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["req"]
		temp["res"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["res"]
		saveRR.push temp
		temp = nil

		@@databodyrange = @@tableColumn.get_databodyrange
		temp = {}
		temp["req"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["req"]
		temp["res"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["res"]
		saveRR.push temp
		temp = nil
		@@headerrowrange = nil
		if @@table.showHeaders
			@@headerrowrange = @@tableColumn.get_headerrowrange 
			temp = {}
			temp["req"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["req"]
			temp["res"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["res"]
			saveRR.push temp
			temp = nil		
		end
		if @@table.showTotals
			@@totalrowrange = @@tableColumn.get_totalrowrange 
			temp = {}
			temp["req"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["req"]
			temp["res"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["res"]
			saveRR.push temp
			temp = nil				
		else
			@@totalrowrange = nil
		end
		erb :showtablecolumnrangeinfo, :locals => {
			:table => @@table,
			:tableColumn => @@tableColumn,		
			:range => @@range,  
			:databodyrange =>  @@databodyrange,
			:headerrowrange => @@headerrowrange,
			:totalrowrange => @@totalrowrange,
			:htrace => "ok",
			:rr => saveRR			
		}

	rescue
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE		
		erb :index, :locals => {:status => "Error-Fetching-Table-Col-Range-Info: #{params[:tablename]}",
								:htrace => "ok",
								:rr => saveRR
							   }	
	end	

end



get '/showtablerowrangeinfo' do 


	begin
		table = client.get_table(nil, params[:tablename])
		tableRow = table.get_row(params[:index].to_i)
		range = tableRow.get_range
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE 
		erb :showtablerowrangeinfo, :locals => {
			:table => table,
			:tableRow => tableRow,		
			:range => range,  
			:htrace => "ok",
			:rr => saveRR			
		}

	rescue
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE		
		erb :index, :locals => {:status => "Error-Fetching-Table-Row-Range-Info: #{params[:tablename]}",
								:htrace => "ok",
								:rr => saveRR
							   }	
	end	

end

get '/getnameditemrange' do 

	@@range = client.get_nameditem(params[:name]).get_range
	saveRR = ExcelRubyEasy::REQUEST_RESPONSE
	@@worksheet = @@range.worksheet
	erb :rangegridload, :locals => {
		:range => @@range,
		:worksheet => @@worksheet,
		:rr => saveRR
	}
end 

get '/gettablerangeinfo' do 
	begin
		saveRR = []
		temp = {}
		@@table = client.get_table(nil, params[:tablename])
		@@range = @@table.get_range
		temp = {}
		temp["req"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["req"]
		temp["res"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["res"]
		saveRR.push temp
		temp = nil
		@@databodyrange = @@table.get_databodyrange
		temp = {}
		temp["req"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["req"]
		temp["res"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["res"]
		saveRR.push temp
		temp = nil

		@@headerrowrange = nil
		if @@table.showHeaders
			@@headerrowrange = @@table.get_headerrowrange 		
			temp = {}
			temp["req"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["req"]
			temp["res"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["res"]
			saveRR.push temp
		end
		if @@table.showTotals
			@@totalrowrange = @@table.get_totalrowrange 
			temp = {}
			temp["req"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["req"]
			temp["res"] = ExcelRubyEasy::REQUEST_RESPONSE_SAVE["res"]
			saveRR.push temp
		else
			@@totalrowrange = nil
		end
		saveRR.each do |item| 
		  puts "HERE >> #{item} "
		end 
		erb :showtablerangeinfo, :locals => {
			:table => @@table,
			:range => @@range,
			:databodyrange =>  @@databodyrange,
			:headerrowrange => @@headerrowrange,
			:totalrowrange => @@totalrowrange,
			:htrace => "ok",
			:rr => saveRR			
		}

	rescue
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE		
		erb :index, :locals => {:status => "Error-Fetching-Table-Range-Info: #{params[:tablename]}",
								:htrace => "ok",
								:rr => saveRR
							   }	
	end	

end

get '/listtablecolumns' do 

	begin
		@@table = client.get_table(nil, params[:tablename])
		@@columns = @@table.columns
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE_LIST
		erb :listtablecolumns, :locals => {
			:columns => @@columns,
			:table => @@table,
			:htrace => "ok",
			:rr => saveRR
		}

	rescue
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE		
		erb :index, :locals => {:status => "Error-Fetching-Table-Columns: #{params[:tablename]}",
								:htrace => "ok",
								:rr => saveRR
							   }	
	end	
	
end

post '/tablerowdeleteroute' do 

	begin
		status = client.get_table(nil, params[:tablename]).get_row(params[:index].to_i).delete
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE
		erb :index, :locals => {:status => "Deleted-Table-Row-OK: #{params[:tablename]} at  #{params[:index]}", 
								:htrace => "ok",
								:rr => saveRR
							   }		 
	rescue
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE
		erb :index, :locals => {:status => "Error-Deleting-Table-Row: #{params[:tablename]} at  #{params[:index]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end
end

post '/tablecoldeleteroute' do 
	begin
		status = client.get_table(nil, params[:tablename]).get_column(params[:id].to_i).delete
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE
		erb :index, :locals => {:status => "Deleted-Table-Column-OK: #{params[:tablename]} id  #{params[:id]}", 
								:htrace => "ok",
								:rr => saveRR
							   }		 
	rescue
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE		
		erb :index, :locals => {:status => "Error-Deleting-Table-Column: #{params[:tablename]} id  #{params[:id]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end	
end


get '/listnames' do 
	logger.debug "T Inside list names"
	@@names = client.list_objects("names")

	erb :listnames, :locals => {:names => @@names,
								:rr => ExcelRubyEasy::REQUEST_RESPONSE
								} 
end

get '/listcharts' do 
	begin
		@@charts = client.get_worksheet(nil, params[:sheetname]).charts
		erb :listcharts, :locals => {:charts => @@charts, :sheetname => params[:sheetname],
								 :rr => ExcelRubyEasy::REQUEST_RESPONSE}   
	rescue Exception => e 
		ExcelRubyEasy::logger.warn "ERROR!! #{e.to_s}"
		redirect "/?htrace=none&status=Error-List-Charts"
	end
end

get '/rangemethodoptions' do 
	
	erb :rangemethodoptions  
end

get '/addworksheet' do 
	
	erb :addworksheet 
end

get '/updateworksheet' do 	
	erb :updateworksheet, :locals => {:id => params[:id]}  
end

get '/chartupdate' do 	
	erb :chartupdate, :locals => {:id => params[:id], :sheetname => params[:sheetname]}  
end

post '/chartsetpositionroute' do 
	begin
		chartid = URI.unescape params[:id]
		sheetname = URI.unescape params[:sheetname]		
		worksheet = client.get_worksheet(nil, sheetname)
		chart = worksheet.get_chart(chartid)
		chart.set_position(params[:startcell], params[:endcell])
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE
		erb :index, :locals => {:status => "Chart-Set-Position-OK: #{chart.name}", 
								:htrace => "ok",
								:rr => saveRR
								}		 
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Chart-Set-Position: #{params[:id]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end
end

post '/chartsetsourceroute' do 
	begin
		chartid = URI.unescape params[:id]
		sheetname = URI.unescape params[:sheetname]		
		worksheet = client.get_worksheet(nil, sheetname)
		chart = worksheet.get_chart(chartid)
		chart.set_source(params[:sourcerange], params[:seriesby])
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE
		erb :index, :locals => {:status => "Chart-Set-Position-OK: #{chart.name}", 
								:htrace => "ok",
								:rr => saveRR
								}		 
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Chart-Set-Position: #{params[:id]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end

end

post '/chartsetpropupdateroute' do 
	begin
		id = URI.unescape params[:id]
		chartid = URI.unescape params[:id]
		sheetname = URI.unescape params[:sheetname]		
		worksheet = client.get_worksheet(nil, sheetname)
		chart = worksheet.get_chart(chartid)

		if !params[:name].nil? && !params[:name].empty?
			chart.name = params[:name]
		end
		if !params[:top].nil? && !params[:top].empty?
			chart.top = params[:top].to_i
		end
		if !params[:left].nil? && !params[:left].empty?
			chart.left = params[:left].to_i
		end
		if !params[:height].nil? && !params[:height].empty?
			chart.height = params[:height].to_i
		end
		if !params[:width].nil? && !params[:width].empty?
			chart.width = params[:width].to_i
		end

		chart.update
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE
		erb :index, :locals => {:status => "Updated-Worksheet-OK: #{worksheet.name}", 
								:htrace => "ok",
								:rr => saveRR
								}		 
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Updating-Worksheet: #{params[:id]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end
end

get '/updatetable' do 
	
	@@table = client.get_table(params[:id].to_i)
	erb :updatetable, :locals => {:table => @@table}  
end

get '/downloadimage' do 
	chart = client.get_worksheet(nil, params[:sheetname]).get_chart(params[:id])
	image = chart.get_chart_image
	erb :chartimage, :locals => {:chartimage => image, :chartname => chart.name,
								 :rr => ExcelRubyEasy::REQUEST_RESPONSE}  
end

post '/updatetableroute' do 
	begin
		id = URI.unescape params[:id]
		table = client.get_table(id)
		table.name = params[:name]
		table.style = params[:style].empty? ? nil : params[:style]
		table.showHeaders = (params[:showHeaders] == 'true') ? true : false
		table.showTotals = (params[:showTotals] == 'true') ? true : false
		table.update
		#redirect "/?htrace=none&status=Updated-Table%20" + table.id.to_s
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE
		erb :index, :locals => {:status => "Updated-Table-OK: #{table.name}", 
								:htrace => "ok",
								:rr => saveRR
							   }		 
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Updating-Table: #{params[:id]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end

end

get '/tablerowupdate' do 
	
	@@table = client.get_table(params[:id].to_i)
	@@tableRow = @@table.rows[params[:index].to_i]
	erb :tablerowupdate, :locals => {
									:tableRow => @@tableRow,
									:id => params[:id],
									:index => params[:index],
									:table => @@table
								  }  

end

post '/tablerowupdateroute' do 
	begin
		table = client.get_table(params[:id].to_i)
		tableRow = table.rows[params[:index].to_i]
		values1d = []   
		# params[:cols].to_i.times do |i|
		#	values1d << params["cell#{i.to_s}".to_sym]
		# end
		params.each do |key, item| 
			if key.to_s.include?('cell')
				item = nil if item.empty?
				values1d << item
			end
		end
		values2d = [values1d]	   

		tableRow.values = values2d
		tableRow.update
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE					  
#		redirect "/?htrace=none&status=Updated-Table-Row%20" + table.id.to_s + "%20" +  tableRow.index.to_s
		erb :index, :locals => {:status => "Updated-Table-Row-OK: #{params[:id].to_s}, #{tableRow.index.to_s}", 
								:htrace => "ok",
								:rr => saveRR
							   }   
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Updating-Table-Row: #{params[:id]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end
end

get '/tablecolumnupdate' do 
	@@table = client.get_table(params[:id].to_i)
	@@tableColumn = @@table.get_column(params[:colid].to_i)
	erb :tablecolumnupdate, :locals => {
									:tableColumn => @@tableColumn,
									:id => params[:id],
									:index => params[:index],
									:table => @@table
								  }  
end

post '/tablecolumnupdateroute' do 
	begin
		table = client.get_table(params[:id].to_i)
		tableColumn = table.get_column(params[:colid].to_i)
		values2d = []   
		params.each do |key, item| 
			if key.to_s.include?('cell')
				item = nil if item.empty?			   
				values2d << [item]
			end
		end   
		tableColumn.values = values2d
		tableColumn.update
		#redirect "/?htrace=none&status=Updated-Table-Column%20" + table.id.to_s + "%20" +  tableColumn.index.to_s

		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Updated-Table-Column-OK: #{params[:id].to_s}, #{params[:colid]}", 
								:htrace => "ok",
								:rr => saveRR
							   }   
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Updating-Table-Column: #{params[:id]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end
end


get '/addrows' do 
	
	erb :addrows, :locals => {:tablename => params[:tablename], 
							  :cols => params[:cols],
							  :rows => params[:rows]
							 }  
end

get '/addcolumns' do 
	
	erb :addcolumns, :locals => {:tablename => params[:tablename],
								 :rows => params[:rows],
								 :cols => params[:cols]							  
								 }  
end

get '/addtable' do 
	
	erb :addtable, :locals => {:worksheetid => params[:id]}
end

get '/addchart' do 
	
	erb :addchart, :locals => {:worksheetid => params[:id]}
end

post '/addworksheetroute' do 

	begin
		objectX = client.add_worksheet(params[:sheetnameopt])
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE		
	#	redirect '/?htrace=none&status=Added-Worksheet:%20' + objectX.name
		erb :index, :locals => {:status => "Added-Worksheet: #{objectX.name}", 
								:htrace => "ok",
								:rr => saveRR
								} 
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Adding-Worksheet: #{params[:sheetnameopt]}", 
								:htrace => "ok",
								:rr => saveRR
							   } 
	end

end

post '/addtableroute' do 
	begin
		hasheader = false
		if params[:hasheader] == 'true'
			hasheader = true
		end
		if !params[:id].to_s.empty?						
			table = client.get_worksheet(params[:id]).add_table(params[:rangeaddress], hasheader)
		else
			table = client.add_table(params[:rangeaddress], hasheader)
		end
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE		
		erb :index, :locals => {:status => "Added-Table: #{table.name}", 
								:htrace => "ok",
								:rr => saveRR
								} 
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Adding-Table", 
								:htrace => "ok",
								:rr => saveRR
							   } 
	end	
end


post '/addchartroute' do 
	begin
		chart = client.get_worksheet(params[:id]).add_chart(params[:type],params[:sourcedata],params[:seriesby])
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE		
		erb :index, :locals => {:status => "Added-Chart: #{chart.name}", 
								:htrace => "ok",
								:rr => saveRR
								} 
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Adding-Table", 
								:htrace => "ok",
								:rr => saveRR
							   } 
	end	
end



post '/addrowroute' do 
		# values2d = [params[:rowdata].split(',')]
	begin	
		values1d = []   
		params.each do |key, item| 
			if key.to_s.include?('cell')
				values1d << item
			end
		end
		values2d = [values1d]	   
		indexval = params[:index] == 'lastrow' ?  nil : params[:index].to_i 
		tablerow = client.get_table(nil, params[:tablename]).add_row(values2d, indexval)
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE
		erb :index, :locals => {:status => "Added-Table-Row-OK: #{params[:tablename].to_s}, #{tablerow.index.to_s}", 
								:htrace => "ok",
								:rr => saveRR
							   }   
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Adding-Table-Row: #{params[:tablename]}, #{params[:index]}", 
								:htrace => "ok",
								:rr => saveRR
							   } 
	end   
end

post '/addcolumnroute' do 
	begin
		values2d = []   

		params.each do |key, item| 
			if key.to_s.include?('cell')
				values2d << [item]
			end
		end
		indexval = params[:index].to_i 
		tablecolumn = client.get_table(nil, params[:tablename]).add_column(values2d, indexval)
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE
		erb :index, :locals => {:status => "Added-Table-Column-OK: #{params[:tablename].to_s}, #{params[:index]}", 
								:htrace => "ok",
								:rr => saveRR
							   }   
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Adding-Table-Column: #{params[:tablename]}, #{params[:index]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end		
end

post '/updateworksheetroute' do 
	begin
		id = URI.unescape params[:id]
		worksheet = client.get_worksheet(id)
		if !params[:sheetname].nil? && !params[:sheetname].empty?
			worksheet.name = params[:sheetname]
		end
		if !params[:position].nil? && !params[:position].empty?
			worksheet.position = params[:position].to_i
		end
		worksheet.update
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE
		erb :index, :locals => {:status => "Updated-Worksheet-OK: #{worksheet.name}", 
								:htrace => "ok",
								:rr => saveRR
								}		 
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Updating-Worksheet: #{params[:id]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end
end

get '/deleteworksheetroute' do 
	begin
		id = URI.unescape params[:id]
		worksheet = client.get_worksheet(id)
		worksheet.delete
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			 
		erb :index, :locals => {:status => "Deleted-Worksheet: #{params[:id]}", 
								:htrace => "ok",
								:rr => saveRR
								} 
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Deleting-Worksheet: #{params[:id]}", 
								:htrace => "ok",
								:rr => saveRR
							   }		 
	end 
end


post '/rangegridroute' do 
	erb :rangegrid, :locals => {
							:sheetname => params[:sheetname],
							:addr => params[:addr], 
							:rows => params[:rows], 
							:cols => params[:cols] 
							}  
end

#loc1 Read all format values here and pass to locals than reading in the ERB
post '/rangeformatroute' do 
	if (params[:sheetname].nil? || params[:sheetname].empty?)
		@@range = client.get_range(params[:addr]) 
	else
		@@range = client.get_range(nil, params[:sheetname], params[:addr])
	end
	@@worksheet = @@range.worksheet
	erb :rangeformatload, :locals => {
		:range => @@range,
		:worksheet => @@worksheet
	}

end


post '/rangesyncpreproute' do 
	if (params[:sheetname].nil? || params[:sheetname].empty?)
		@@range = client.get_range(params[:addr]) 
	else
		@@range = client.get_range(nil, params[:sheetname], params[:addr])
	end
	saveRR = {}
	saveRR["req"] = ExcelRubyEasy::REQUEST_RESPONSE["req"]
	saveRR["res"] = ExcelRubyEasy::REQUEST_RESPONSE["res"]	

	@@worksheet = @@range.worksheet
	erb :rangegridload, :locals => {
		:range => @@range,
		:worksheet => @@worksheet,
		:rr => saveRR
	}
end

get '/rangesyncpreproute' do 
	@@range = client.get_range(nil, params[:sheetname], params[:addr])
	saveRR = {}
	saveRR["req"] = ExcelRubyEasy::REQUEST_RESPONSE["req"]
	saveRR["res"] = ExcelRubyEasy::REQUEST_RESPONSE["res"]	
	@@worksheet = @@range.worksheet
	erb :rangegridload, :locals => {
		:range => @@range,
		:worksheet => @@worksheet,
		:rr => saveRR
	}
end

post '/rangedeleteroute' do 
	begin
		if (params[:sheetname].nil? || params[:sheetname].empty?)
			@@range = client.get_range(params[:addr]) 
		else
			@@range = client.get_range(nil, params[:sheetname], params[:addr])
		end

		status = @@range.delete(params[:shift])
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  

		erb :index, :locals => {
				:status => "Deleted-Range: #{params[:addr]}", 
									:htrace => "ok",				
									:rr => saveRR
									}
	rescue								
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  

		erb :index, :locals => {:status => "Error-Deleting-Range: #{params[:addr]}", 
									:htrace => "ok",
									:rr => saveRR
								   }	
	end
end

post '/rangeclearroute' do 
	begin
		if (params[:sheetname].nil? || params[:sheetname].empty?)
			@@range = client.get_range(params[:addr]) 
		else
			@@range = client.get_range(nil, params[:sheetname], params[:addr])
		end

		status = @@range.clear(params[:applyto])
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {
								:status => "Cleared-Range: #{params[:addr]}", 
								:htrace => "ok",
								:rr => saveRR
							   }
	rescue
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  

		erb :index, :locals => {:status => "Error-Clearing-Range: #{params[:addr]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end	
end

post '/tableconverttorange' do 
	begin
		id = URI.unescape params[:id]
		table = client.get_table(id)
		puts "convert to range2"
		table.convert_to_range
		puts "convert to range3"
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {
								:status => "Convert-to-Range-Ok: #{table.name}", 
								:htrace => "ok",
								:rr => saveRR
								}
	rescue
		erb :index, :locals => {:status => "Error-Convert-to-Range", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end 
end

post '/rangeinsertroute' do 
	begin
		if (params[:sheetname].nil? || params[:sheetname].empty?)
			@@range = client.get_range(params[:addr]) 
		else
			@@range = client.get_range(nil, params[:sheetname], params[:addr])
		end
		@@newrange = @@range.insert(params[:shift])
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {
								:status => "Inserted-Range: #{@@newrange.address}", 
								:htrace => "ok",
								:rr => saveRR
								}
	rescue
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-Inserting-Range", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end 
end


post '/worksheetrangeopsroute' do 

	begin
		case params[:op]
		when 'getcell'
			@@worksheet = client.get_worksheet(params[:id])
			@@range = @@worksheet.get_cell(params[:row], params[:column])
			saveRR = ExcelRubyEasy::REQUEST_RESPONSE		
			erb :rangegridload, :locals => {
					 :range => @@range,
					 :worksheet => @@worksheet,
					 :rr => saveRR
			}
		when 'usedrange'
			@@worksheet = client.get_worksheet(params[:id])
			@@range = @@worksheet.get_usedrange
			saveRR = ExcelRubyEasy::REQUEST_RESPONSE		
			erb :rangegridload, :locals => {
					 :range => @@range,
					 :worksheet => @@worksheet,
					 :rr => saveRR
			}	   
		when 'getrange'
			@@worksheet = client.get_worksheet(params[:id])
			@@range = @@worksheet.get_range(params[:addr])
			saveRR = ExcelRubyEasy::REQUEST_RESPONSE		
			erb :rangegridload, :locals => {
					 :range => @@range,
					 :worksheet => @@worksheet,
					 :rr => saveRR
			}	   
		end

	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-With-Range-Read: #{params[:method]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end

end


get '/worksheetrangeops' do 
	erb :worksheetrangeops, :locals => {:id => params[:id]}  

end

post '/rangecellroute' do 
	if (params[:sheetname].nil? || params[:sheetname].empty?)
		@@range = client.get_range(params[:addr]) 
	else
		@@range = client.get_range(nil, params[:sheetname], params[:addr])
	end

	if params[:method] == 'getcell'
		@@out = @@range.get_cell(params[:row], params[:column])
	else
		@@out = @@range.get_offsetrange(params[:row], params[:column])
	end
	saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE
	@@worksheet = @@range.worksheet
	erb :rangegridload, :locals => {
				 :range => @@out,
				 :worksheet => @@worksheet,
				 :rr => saveRR
	}   
end


post '/rangeboundingroute' do
	
end


post '/rangerowcolumnroute' do 
	if (params[:sheetname].nil? || params[:sheetname].empty?)
		@@range = client.get_range(params[:addr]) 
	else
		@@range = client.get_range(nil, params[:sheetname], params[:addr])
	end
	@@worksheet = @@range.worksheet


	case params[:method]
	when 'row'
		@@out = @@range.get_row(params[:number])
	when 'column'						 
		@@out = @@range.get_column(params[:number]) 
	end 
	saveRR = ExcelRubyEasy::REQUEST_RESPONSE_SAVE
	erb :rangegridload, :locals => {
				 :range => @@out,
				 :worksheet => @@worksheet,
				 :rr => saveRR
	}   
end

post '/rangesimplegetroute' do 
	begin
		if (params[:sheetname].nil? || params[:sheetname].empty?)
			@@range = client.get_range(params[:addr]) 
		else
			@@range = client.get_range(nil, params[:sheetname], params[:addr])
		end
		@@worksheet = @@range.worksheet

		case params[:method]
			when 'entirecolumn'
				@@responseRange = @@range.get_entirecolumn
			when 'entirerow'						 
				@@responseRange = @@range.get_entirerow
			when 'lastrow'				
				@@responseRange = @@range.get_lastrow
			when 'lastcolumn'				
				@@responseRange = @@range.get_lastcolumn								  
			when 'lastcell'										 
				@@responseRange = @@range.get_lastcell
			when 'usedrange'
				@@responseRange = @@range.get_usedrange
		end
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE	

		erb :rangegridload, :locals => {
			:range => @@responseRange,
			:worksheet => @@worksheet,
			:rr => saveRR
		}
	rescue Exception => e 
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE			  
		erb :index, :locals => {:status => "Error-With-Range-Read: #{params[:method]}", 
								:htrace => "ok",
								:rr => saveRR
							   }	
	end

end

#loc1		
post '/rangesyncroute' do 
	begin
		keyarr = []
		params.each do |key, item| 
			if key.to_s.include?('cell')
				keyarr << key
			end
		end
		values1d, values2d = [], []
		keyarr.sort.each_with_index do |key, i|
			val = params[key.to_sym].empty? ? nil : params[key.to_sym]
			values1d << val
			if ((i+1) % params[:colCount].to_i) == 0				
				values2d << values1d
				values1d = []
			end
		end
		range = client.get_worksheet(nil, params[:sheetname]).get_range(params[:addr])
		range.values = values2d
		range = range.sync(true)
		saveRR = ExcelRubyEasy::REQUEST_RESPONSE
		#redirect '/?htrace=none&status=Sync-Range-Success:%20' + range.address


		erb :index, :locals => {:status => "Sync-Range-OK: #{range.address}", 
								:htrace => "ok",
								:rr => saveRR
							   }		 
	rescue Exception => e 
		erb :index, :locals => {:status => "Error-Synching-Range",
								:htrace => "ok",
								:rr => saveRR
							   }	
	end		

end

#loc1
post '/updaterangeformatroute' do 

	range = client.get_worksheet(nil, params[:sheetname]).get_range(params[:addr])
	puts "000 #{params[:sheetname]}, #{params[:horizontalalignment]}"
	worksheet = range.worksheet	
	#replace null and blanks to nil

	params.each do |k, v|
		if v == 'null' || v == ""
			params[k] = nil
		end
	end
	format = range.format
	format.horizontalAlignment = params[:horizontalalignment]
	format.verticalAlignment = params[:verticalalignment]

	if params[:wraptext]
		format.wrapText = (params[:wraptext] == 'true') ? true : false
	end
	format.update

	fill = range.format.fill
	fill.color = params[:fillcolor]
	fill.update

	font = range.format.font
	if params[:fontbold]
		font.bold = (params[:fontbold] == 'true') ? true : false
	end
	
	font.color = params[:fontcolor]
	if params[:fontitalic]
		font.italic = (params[:fontitalic] == 'true') ? true : false
	end
	font.name = params[:fontname]
	if params[:fontsize]
		font.size = params[:fontsize].to_i
	end
	font.underline = params[:fontunderline]
	font.update

	erb :rangeformatload, :locals => {
		:range => @@range,
		:worksheet => @@worksheet
	}
end

