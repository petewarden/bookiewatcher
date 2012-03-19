#***********************************************************************************
#
# Watches the results of the bookies awards poll
#
# (C) Pete Warden <pete@petewarden.com>
#
#***********************************************************************************

require 'rubygems' if RUBY_VERSION < '1.9'

require 'google_spreadsheet'
require 'net/https'

# Configure these to match the poll you want to watch
BOOKIE_POLL_ID='6011539'
BOOKIE_SPREADSHEET_KEY='0AjL2XrwdNUq7dFg1cFA2d3hGd2RTaTJXLTMxNm5xSHc'
MINUTES_BETWEEN_CHECKS=2
BOOKIE_POLL_PREFIX='http://polls.polldaddy.com/vote-js.php?p='

def get_http(url)
  
  uri = URI.parse(url)
  http = Net::HTTP.new(uri.host, uri.port)
  
  headers = { 'User-Agent' => ENV['CRAWLER_USER_AGENT'] }
  request = Net::HTTP::Get.new(uri.request_uri, headers)
  
  begin
    response = http.request(request)  
    
    result = nil
    if response.code != '200'
      log "Bad response code #{response.status} for '#{url}'"
      result = nil
      else
      result = response.body
    end
  rescue NoMethodError
    result = nil
  end
  
  result
end

def update_bookie_spreadsheet()
  
  retry_count = 0
  
  while retry_count < 3

    begin
    
      bookie_url = BOOKIE_POLL_PREFIX+BOOKIE_POLL_ID
      bookie_content = get_http(bookie_url)
      
      matches = bookie_content.scan(/<span class=\"pds-answer-text\">([^<]+)<\/span><span class=\"pds-feedback-result\"><span class=\"pds-feedback-per\">&nbsp;([0-9\.]+)%<\/span>/)

      time_string = Time.now.strftime("%Y-%m-%d %H:%M")
    
      session = GoogleSpreadsheet.login(ENV['GOOGLE_DOCS_ACCOUNT'], ENV['GOOGLE_DOCS_PASSWORD'])
      ws = session.spreadsheet_by_key(BOOKIE_SPREADSHEET_KEY).worksheets[0]
      
      header_row = 1
      header_map = {}
      header_column = 1
      first_empty_header = nil
      while (header_column <= ws.num_cols)
        header_name = ws[header_row, header_column]
        if header_name == ''
          if !first_empty_header
            first_empty_header = header_column
          end
        else
          header_map[header_name] = header_column
        end
        header_column += 1
      end
      if !first_empty_header
        first_empty_header = (ws.num_cols + 1)
      end
      
      if !header_map['Time']
        ws[header_row, first_empty_header] = 'Time'
        header_map['Time'] = first_empty_header
        first_empty_header += 1
      end
      time_column = header_map['Time']
      
      insertion_row = 2
      while insertion_row <= ws.num_rows
        current_time = ws[insertion_row, time_column]
        if current_time != time_string and current_time != ''
          insertion_row += 1
        else
          break
        end
      end

      ws[insertion_row, time_column] = time_string
      
      matches.each do |title, percent|
        title = title.strip
        if !header_map[title]
          ws[header_row, first_empty_header] = title
          header_map[title] = first_empty_header
          first_empty_header += 1
        end
        title_column = header_map[title]
        ws[insertion_row, title_column] = percent+'%'
      end
        
      ws.save()
      $stderr.puts "bookie spreadsheet updated"
      break
    rescue Exception => e
      retry_count += 1
      $stderr.puts "Error updating bookie spreadsheet on try #{retry_count}"
      $stderr.puts "Error: #{e}: #{e.message}"
    end
  end

end

if __FILE__ == $0

  while true
    update_bookie_spreadsheet
    sleep MINUTES_BETWEEN_CHECKS * 60
  end

end
