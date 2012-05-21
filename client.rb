require File.join(File.dirname(__FILE__), 'requester')
require 'spreadsheet'
require 'yaml'
require 'yajl'

URL = {
  :hasoffers => "http://sponsorpaynetwork.api.hasoffers.com/Api/json"
}

config = YAML.load_file("config.yaml")
@network_id = config["config"]["network_id"]
@network_token = config["config"]["network_token"]

#Create result file

result = Spreadsheet::Workbook.new
sheet1 = result.create_worksheet
sheet1.name = 'Result'
sheet1.row(0).concat %w{LandingPageID HasOffersID Country Status}

row_counter = 1

#Read Excel File
Spreadsheet.open('non.xls') do |book|
  book.worksheet('Sheet1').each do |row|
    break if row[0].nil?
      next if row[1] == "AFFILIATE_OFFER_ID"
        response = Requester.make_request(
        URL[:hasoffers],
        {
          "Format"             => "json",
          "Service"            => "HasOffers",
          "Version"            => "2",
          "NetworkId"          => "#{@network_id}",
          "NetworkToken"       => "#{@network_token}",
          "Target"             => "Offer",
          "Method"             => "findById",
          "id"                 => "#{row[1]}"
        },
        :get 
      )
      
      json = StringIO.new("#{response}")
      parser = Yajl::Parser.new
      hash = parser.parse(json)
      if hash["response"]["data"].nil?
        excel_row = sheet1.row(row_counter)
        excel_row[0] = row[0]
        excel_row[1] = row[1]
        excel_row[2] = row[2]
        excel_row[3] = "No data"
      else
        excel_row = sheet1.row(row_counter)
        excel_row[0] = row[0]
        excel_row[1] = row[1]
        excel_row[2] = row[2]
        excel_row[3] = hash["response"]["data"]["Offer"]["status"]
      end
      row_counter += 1
  end
end

result.write 'result.xls'
