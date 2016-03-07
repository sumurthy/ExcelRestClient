require_relative '../lib/excelrubyeasy'
require 'logger'
require 'uri'
require 'json'

@master_error_count = 0
@sub_error_count = 0

session_config = {}

session_config = JSON.parse(File.read('../../secret.json'), {:symbolize_names => true})

@client = ExcelRubyEasy::Client.new session_config
@fileid = "01353TX3HLZMV7SFS77FE3564GR7DWDPFP"
#@fileid = "01353TX3C2RASTFWIAQRC2QTTDAJZBLZHX"
begin
	@client.set_access_token('AAABAAAAiL9Kn2Z27UubvWFPbm0gLa3SJsxM_Q3I8YQ1_zgme3ktnU-G50Ttzci8U-gfkx8TUmzNvgFnnJVqfzRQ0nZwNGzHR__bOmNtRP2FO-c8wtEM24lRDPebrdsIFoFAcNXh-RcviRsYH5AYYZUArpPFvZLLWifI499xToxJ7S5l4RmQcw0KiISdH5G_erc7A028_6Qb8sOvgqwB3nmgFRLjPrhpGhQ1bfqNoaXjDcI6A4M2VSXhjiPAmVrZRYjNBDC1tBMKFV-XlGtm1jIRJs0nWth7fF2mzdgMxzxXLPwdNPzfZ3Ud9fj98DR4cpBY2E3YlECL51-20EPPRssKtApndIqzISqJSUna73MdVZdFLEdheL8ppTO5eMUc2HuRTfigDSLeC1f7BITHxvtv86rVRvRib4oN7_iR5uphC5OW7ePI3vvAVAP9LSjiOX8omFzE_11CN_q7NVTASQ7Eb3ne41gQt-GxMyGdRs1Xba_onKsdMdKc0x_kupR6EfQaRB1F6Pa3pgLOzCY6N5rlMf3_MCAA', true) 
	puts "Created access token"
	@client.set_excelfile_and_session(@fileid) 
	puts "Created workbook session"
	@logtable = @client.add_table('Sheet1!A1:D1', false)	 
rescue Exception => e 
	puts " Abort stage-1"
	puts e.status_code
	exit 
end

puts "Created Table - Waiting to begin insert ops."
sleep 10 

puts "Starting insert ops."

@master_req_count = 0
@sub_req_count = 0

600.times do |n|
	puts "Insert op: #{n}"

	begin
		values2d = [["#{@master_req_count}", "#{@sub_req_count}", "#{@sub_error_count}", "#{Time.now}"]]
		tablerow = @logtable.add_row(values2d)
		@master_req_count = @master_req_count + 1
		@sub_req_count = @sub_req_count + 1
	rescue Exception => e 
		puts "!! Error: #{e.status_code}"
		@master_error_count = @master_error_count + 1
		if e.status_code == 401
			@client.refresh_access_token
		else
			@sub_error_count = @sub_error_count + 1
			if @sub_error_count > 2
				puts "Reset session id"
				@client.set_excelfile_and_session(@fileid) 
				@sub_req_count = 0
				@sub_error_count = 0
			end
		end		
		if @master_error_count > 50
			puts
			puts "!!!!! Too many errors. Abort."
			puts
			exit 99
		end		
	end
	puts "sleeping..."
	sleep 55
end

puts "--- completed ---"

puts ""
puts "@master_req_count, #{@master_req_count}"
puts "@master_error_count, #{@master_error_count} "
puts ""
exit 0