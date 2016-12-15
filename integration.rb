
require 'rubygems'
require 'roo'
require 'roo-xls'
require 'csv'
require 'fileutils'
require 'date'
require 'json'
require 'unirest'
require 'zip'



#some variables to make it do stuff and things
basefolder = "/Users/YourUser/somefolder/Integration" 
access_token = ""
domain = "" #subdomain only
env = nil 
#dont touch this stuff. 


time = Time.now.to_i
source_folder = "upload"
archive_folder = "archive"
output_filename = "upload/enrolment_upload_#{time}.csv"

def deep_remove!(text, array)
  array.delete_if do |value|
    case value
    when String
      value.include? text
    when Array
      deep_remove!(text, value)
      false
    else
      false
    end
  end
end


Dir.glob("#{basefolder}/data/*.xlsx") do |file|
	file_path = "#{file}"  
	file_basename = File.basename(file, ".xlsx") 
	xls = Roo::Excelx.new(file_path)
	xls.to_csv("#{basefolder}/csv/#{file_basename}.csv")
end


CSV.open("#{basefolder}/student/student_output.csv", "w") do |f|
	f << ["section_id","user_id","role","status"]
end

	
Dir.glob("#{basefolder}/csv/*.csv") do |csvfile|
    
    col_data = Array.new

    CSV.foreach(csvfile,'r') {|row| $num = row.size}
		csv = CSV.read(csvfile, :headers=>false).drop(2)		
		raw_enrolment_data = csv.transpose.drop(1) 
		deep_remove!('Teacher:', raw_enrolment_data)
		deep_remove!('Room:', raw_enrolment_data)
		deep_remove!('Students: ', raw_enrolment_data)
		clean_data = raw_enrolment_data.map { |e| e.instance_of?(Array) ? e.compact : e }.compact
		
		   	CSV.open("#{basefolder}/student/student_output.csv", "a") do |f|
				clean_data.each do |x|
				  	$i = 1
				  	$num = x.count - 1

			  	while $i < $num
				   	f << [x[0].to_s,x[$i].to_s,"student","active"]
				   	$i +=1
			   	end

		    end
	end
			
end


teachers_data = File.read('teacher/teacher_output.csv')
students_data = File.read('student/student_output.csv')
File.write(output_filename, students_data)


FileUtils.rm_rf(Dir.glob('student/*'))
FileUtils.rm_rf(Dir.glob('csv/*'))

#push
env ? env << "." : env
test_url = "https://#{domain}.#{env}instructure.com/api/v1/accounts/self"
endpoint_url = "#{test_url}/sis_imports.json?import_type=instructure_csv"

test = Unirest.get(test_url, headers: { "Authorization" => "Bearer #{access_token}" })

unless test.code == 200
	raise "Error: The token, domain, or env variables are not set correctly"
end

unless Dir.exists?(source_folder)
	raise "Error: source_folder isn't a directory, or can't be located."
end

unless Dir.entries(source_folder).detect {|f| f.match /.*(.csv)/}
	raise "Error: There are no CSV's in the source directory"
end

unless Dir.exists?(archive_folder)
	Dir.mkdir archive_folder
	puts "Created archive folder at #{archive_folder}"
end


files_to_zip = []

Dir.foreach(source_folder) { |file| files_to_zip.push(file) }


zipfile_name = "#{source_folder}/archive.zip"


Zip::File.open(zipfile_name, Zip::File::CREATE) do |zipfile|
	files_to_zip.each do |file|
		zipfile.add(file, "#{source_folder}/#{file}")

	end

end






upload = Unirest.post(endpoint_url,
	headers: {
		"Authorization" => "Bearer #{access_token}"
  },
	parameters: {
		attachment: File.new(zipfile_name, "r")
	}
)
job = upload.body

FileUtils.mv(zipfile_name, "#{archive_folder}/archive-#{time}.zip")
FileUtils.rm(output_filename)

import_status_url = "#{test_url}/sis_imports/#{job['id']}"

while job["workflow_state"] == "created" || job["workflow_state"] == "importing"
	puts "importing"
	sleep(3)

  import_status = Unirest.get(import_status_url,
  	headers: {
  	  "Authorization" => "Bearer #{access_token}"
    }
  )
	job = import_status.body
end

if job["processing_errors"]
	File.delete(zipfile_name)
	raise "An error occurred uploading this file. \n #{job}"
end

if job["processing_warnings"]
	puts "Processing Errors: #{job["processing_errors"]}"
end

puts "Successfully uploaded files"





