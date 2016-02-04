require 'logger'

module ExcelRubyEasy
	
	## 
    # Logging location
    #

    LOG_FOLDER = "Logs"
    LOG_FILE = "#{LOG_FOLDER}/#{$PROGRAM_NAME.chomp('.rb')}.txt"
    Dir.mkdir(LOG_FOLDER) unless File.exists?(LOG_FOLDER)
    if File.exists?(LOG_FILE)
       File.delete(LOG_FILE)
    end
	
	class << self
	  ##
	  # Logger for the API client
	  #
	  # @return [Logger] logger instance.
	  ## todo: change to relative folder or make it part of options.
	  
	  attr_accessor :logger
	end

	self.logger = Logger.new(ExcelRubyEasy::LOG_FILE)
	self.logger.level = Logger::DEBUG  


	##
	# Module to make accessing the logger simpler
	module Logging
	  ##
	  # Logger for the API client
	  #
	  # @return [Logger] logger instance.
	  def logger
		ExcelRubyEasy::logger
	  end
	end
  
end