require 'uri'
require 'json'

module ExcelRubyEasy
  module Model
	class NamedItem 
		attr_accessor :name, :type, :visible, :value

		def initialize (parms={}) 
			@name = parms[:name]
			@type = parms[:type]
			@visible = parms[:visible]
			@value = parms[:value]			
		end

		def get_range(address=nil)
			ExcelRubyEasy::logger.debug "D, #{__method__.to_s}"		 
			j = ExcelRubyEasy::HttpAction::doGetRequest (ExcelRubyEasy::Client.excelserver + "Names('#{URI.escape name}')/Range")
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