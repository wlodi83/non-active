require File.join(File.dirname(__FILE__), 'exasol')
require 'spreadsheet'
require 'yaml'

config = YAML.load_file("config.yaml")
@login = config["config"]["login"]
@password = config["config"]["password"]

#Create result file
result_excel = Spreadsheet::Workbook.new
sheet1 = result_excel.create_worksheet
sheet1.name = 'Result'
sheet1.row(0).concat %w{ProgramName LandingPageID HasOffersID Country Status HasOffers_offer_ID_used_by_other_offer_live_in_all_AMS_levels}

row_counter = 1

@connection = Exasol.new(@login, @password)
@connection.connect

Spreadsheet.open('result.xls') do |book|
  book.worksheet('Result').each do |row|
    next if row[1] == "LandingPageID" && row[2] == "HasOffersID"
    lp = row[1]
    hi = row[2]
    puts lp
    puts hi
    
    query_1 = "select prog.name from cms.advertisers as adv join cms.programs as prog on adv.id = prog.advertiser_id join cms.program_regions as pr on prog.id = pr.program_id join cms.countries as co on pr.country_id = co.id join cms.landing_pages as lp on lp.program_region_id = pr.id where lp.id = '#{lp}'"
    @connection.do_query(query_1)
    result_1 = @connection.print_result_array
    
    query_2 = "select lp.id from cms.affiliate_networks as an join cms.advertisers as adv on an.id = adv.affiliate_network_id join cms.programs as prog on adv.id = prog.advertiser_id join cms.program_regions as pr on prog.id = pr.program_id join cms.countries as co on pr.country_id = co.id join cms.landing_pages as lp on lp.program_region_id = pr.id where lp.\"enabled\" = 1 and adv.\"enabled\" = 1 and prog.\"enabled\" = 1 and lp.id != '#{lp}' and lp.affiliate_offer_id = '#{hi}' and an.id = '52'"
    @connection.do_query(query_2)
    result_2 = @connection.print_result_array

      if result_1.empty? && result_2.empty?
        excel_row = sheet1.row(row_counter)
        excel_row[0] = "null"
        excel_row[1] = row[1]
        excel_row[2] = row[2]
        excel_row[3] = row[3]
        excel_row[4] = row[4]
        excel_row[5] = "NO"
      elsif result_2.empty?
        excel_row = sheet1.row(row_counter)
        excel_row[0] = result_1[0][0]
        excel_row[1] = row[1]
        excel_row[2] = row[2]
        excel_row[3] = row[3]
        excel_row[4] = row[4]
        excel_row[5] = "NO"
      else
        excel_row = sheet1.row(row_counter)
        excel_row[0] = result_1[0][0]
        excel_row[1] = row[1]
        excel_row[2] = row[2]
        excel_row[3] = row[3]
        excel_row[4] = row[4]

        if result_2.length > 1
          multiple_landing_pages = result_2.flatten.join(', ')
          excel_row[5] = multiple_landing_pages
        else
          excel_row[5] = result_2[0][0]
        end

      end

      row_counter += 1

  end

end

@connection.disconnect
result_excel.write 'result_with_program_name.xls'
