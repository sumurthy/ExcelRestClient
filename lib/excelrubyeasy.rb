
require_relative 'excelrubyeasy/logging'
require_relative 'excelrubyeasy/client'

module ExcelRubyEasy

	#EXCEL_SERVER = "https://graph.microsoft.com/testexcel/me/drive/items/01KIA3ZGYNASNDXHVNVBGJYTXYOLMUOHRN/workbook/"
	EXCEL_BASE = "https://graph.microsoft.com/testexcel/me/drive/items/"
	#EXCEL_SERVER = "https://graph.microsoft.com/testexcel/me/drive/items/01353TX3BVODPICVWIBJD3UXNY6VC3JVGD/workbook/"
	RESOURCE_PATH = {"path" => ""}
	ONEDRIVE_SEARCH = "https://graph.microsoft.com/testexcel/me/drive/root/microsoft.graph.search(q='.xlsx')?$select=id,name,size"
	TOKEN = {"token" => ""}

   	HEADERS_PATCH = {
			"Content-Type" => "Application/Json", 
	   		"Accept" => "application/Json"
			}
   	HEADERS_DELETE = {
			"Content-Type" => "Application/Json", 
	   		"Accept" => "application/Json"
			}
   	HEADERS_POST = {
			"Content-Type" => "Application/Json", 
	   		"Accept" => "application/Json"
			}
   	HEADERS_POST_BASIC = {
			"Content-Type" => "Application/Json", 
	   		"Accept" => "Application/Json"
			}
   	HEADERS_GET = {   					
	   		"Accept" => "Application/Json"
			}
   	HEADERS_GET_ALL = {   					
	   		"Accept" => "*/*"
			}			
	SESSION_TYPE = nil   


	REQUEST_RESPONSE = {
			"req" => nil,
			"res" => nil
			}

	REQUEST_RESPONSE_SAVE = {
			"req" => nil,
			"res" => nil
			}
	REQUEST_RESPONSE_SAVE_LIST = {
			"req" => nil,
			"res" => nil
			}
	   	  
end

##
# TO-DOs:
# 1. Data validation based on published specs for creation and updates. 
# 2. Range data as HTML 
# 3. Search in HTML
# 4. Range or table - highlight duplicates in a column 
# 5. Logging , HTTP logging
# 6. Error handling 
# 7. Range update based on what has changed (not have to update range format, formulas, etc. all the time)
# 8. Private public methods
# 9. Handle 404's. Especially Table.totalrowrange
# 10. Headers for GET requests
# 11. Add testing 
# 12. Session handling
# 13. End to end OAuth pieces, OneDrive stuff.
# 15. List charts from worksheet
# 16. relationship handling from objects (Expand)??
# 17. Attr_reader and writer clear separation 
# 18. SEPARATE ALL PARAMS LOAD OPERATIONS TO STATIC CLIENT METHODS. e.g., load_rangeparms(j=nil)
# 19. Functionality: range-grid: numformat, formulas, formulaslocal update, 
# 20. Chart update, delete, relations, and other connected objects. 		
# 21. change Range key to be worksheet Id in all places.
# 22. Qry params: OData ones
# 23. Pagination for all collections 
# 24. HTTP trace for Range format read and update 
# 25. Separate table row range operation to a link
# 26. Remove @@ classvariables from test
# 27. Proper exception handling...