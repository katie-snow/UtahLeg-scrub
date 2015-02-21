require 'rubygems'
require "google/api_client"
require "google_drive"
require "watir-webdriver"

# Collects the data from the reading lists
# Reading Calendars we want to look at
# House 3rd Reading
# Senate Bills
# Senate 2nd Reading
# Senate 3rd Reading
# Table on Second
# Table on Third
def getReadingList (browser, *args)
  puts "--- Processing Reading Calendars ---"
  bills = {}
    
  args.each do |idLoc|
    result = false

    puts "Table '" + idLoc + "'..."
    # Try/Catch (Rescue) because of AJAX page timing problems
    begin
      holdHash = {}
 
      begin
        browser.div(:id, idLoc).ol.lis.each.with_index do |_, idx|
          arr = []
    
          arr << idx + 1 << browser.div(:id, idLoc).ol.li(index: idx).font.attribute_value('class')
        
          # In the reading table a substituted bill has a number before the bill number
          # this regex will exclude the number.
          # Any number of characters that are capital A through capital Z
          # and any number of numbers after.
          # Example: 1HB 23 becomes HB 23
          holdHash[browser.div(:id, idLoc).ol.li(index: idx).text.partition(/[A-Z]+.\d+/)[1]] = arr
        end
        bills.merge!(holdHash)
        result = true
      rescue
        #puts  "Exception occurred, Trying again!"
        result = false
      end
    end unless result
  end

  return bills
end

# The second column(Col B) has the long name for the committees
# The first column (Col A) has the abbreviation for the committees
def searchCommittees(ws, longName)
  str = ''
  for row in 1..ws.num_rows
    if ws[row, 2] == longName
      str =  ws[row, 1]
      break
    end
  end

  if str.empty?
    str = longName
  end

  return str
end

def connect(browser)
  puts "--- Google Login ---"
  puts "Please login with your Google credentials on the Firefox browser"
  
  # Fixes SSL Connection Error in Windows execution of Ruby
  # Based on fix described at: https://gist.github.com/fnichol/867550
  ENV['SSL_CERT_FILE'] = File.expand_path(File.dirname(__FILE__)) + "/cacert.pem"
  
  # Authorizes with OAuth and gets an access token.
  client = Google::APIClient.new(
      application_name: 'Utah Legislature Bill Extraction',
      application_version: '1.0.0')
      
  auth = client.authorization
  auth.client_id = "548733547850-u2bd7pnsi4deou7el50skmoj05uih25a.apps.googleusercontent.com"
  auth.client_secret = "_dW9KyMxZ02i2uhgYbhj0zxy"
  
  # Need both of these scope items to be able to edit google spreadsheets
  auth.scope =
      "https://www.googleapis.com/auth/drive " +
      "https://spreadsheets.google.com/feeds/"
  auth.redirect_uri = "urn:ietf:wg:oauth:2.0:oob:auto"
  
  # Open the Login authorization page
  browser.goto auth.authorization_uri.to_str
  
  # Give the user time to enter their credentials
  Watir::Wait.until(60) { browser.title.include? 'Success code' }
  
  auth.code = browser.title.partition('=')[2]
  auth.fetch_access_token!
  access_token = auth.access_token
  
  # Creates a session.
  session = GoogleDrive.login_with_oauth(access_token)
  return session
end

if __FILE__ == $0
  billYear = '2015'
  puts "Running against " + billYear + " Utah Legislative Session"
  
  # Output FireFox web browser
  browser = Watir::Browser.new :ff

  session = connect(browser)
 
  # Returns nil if not found can do a check here
  lowvData = session.spreadsheet_by_title(billYear + ' Leg Session - LoWV Data')
  if lowvData.nil?
    puts billYear + ' Leg Session - LoWV Data spreadsheet not found!'
    exit
  end

  # Same here change to spreadsheet_by_title
  template = session.spreadsheet_by_title('Legislative Session - Template')
  if template.nil?
    puts 'Legislative Session - Template spreadsheet not found!'
   exit
  end

  output = template.duplicate(billYear + " Legislative Session - Output").worksheets[0]

  committee = lowvData.worksheet_by_title('committees')
  billData = lowvData.worksheet_by_title('bills')

  baseLink = "http://le.utah.gov"
  
  # Reading Calendar Information
  ##############################
  browser.goto baseLink + ':443/FloorCalendars/'
  browser.div(:id, "sexpandbtn").click
  browser.div(:id, "hexpandbtn").click

  # collect reading calendar data, return value is a hash object {bill, circled}
  readingBills = getReadingList(browser, 'dt6', 'dt7', 'dt16', 
                                         'dt17', 'dt18', 'dt19')

  # Get Bill information
  ######################
  # Don't look at the first row it has header information that is not needed by us
  bills = billData.rows.drop(1)

  # the number of extra rows before the bill information starts
  # buffering the work-sheet by 10 rows
  # This is so the last bill doesn't run out of rows when entering sponsor or vote information into the work-sheet
  # the max rows is only updated and save after each bill not during
  output.max_rows = output.max_rows + bills.length + 10
  output.save

  # Track where the current bill is located in the worksheet
  # start at the first empty row
  outputPos = output.num_rows + 1

  puts "--- Processing Bills ---"
  # for loop through all the bills
  bills.each do |bill|
    # Open bill in browser
    browser.goto baseLink + "/~" + billYear + "/bills/static/" + bill[0] + ".html"

    # Bill #
    # Place bill number in output work-sheet
    output[outputPos,1] = bill[0]
    
    # if this text element is not located on the web page, then we are not at the right site
    if !browser.h3.exists?
      puts "WARNING: Bill '" + bill[0] + "' website was not found, check bill number."
      # Place warning in output work-sheet
      output[outputPos, 2] = "WARNING: Bill '" + bill[0] + "' website was not found, check bill number."
      outputPos = outputPos + 1

      # save work-sheet
      output.save
      next
    end

    # Bill Name and Link #
    # Starting at front of the line with any number of characters, then partition on B. or R. 
    # then any number of numbers then space.  
    # The \. are making sure ruby doesn't try to do a function call
    partition = browser.h3.text.partition(/^\S*(B|R)\. [0-9]* /)
    billNum = partition[1].delete('.').chop
    billTitle = partition[2]
    puts "Scrubbing " + bill[0] + " " + billTitle + "..."

    # Bill Link - Assumption Introduced link is at position [1] and Amended link is at position [2]
    billTmp = browser.div(:id, 'billinfo').ul(:id, 'billTextDiv')
    if billTmp.text.include? "Amended"
      billLink = billTmp.a(:class => "nlink", index: 2).attribute_value("href")
    else
      billLink = billTmp.a(:class => "nlink", index: 1).attribute_value("href")
    end

    # Place bill title and link to text in output worksheet
    output[outputPos,2] = '=hyperlink("' + billLink + '";"' + billTitle + '")'

    # Sponsors and Links #
    # Regex Positive Lookbehind functionality.  Search for this if you need help adjusting this regex
    # Search backwards through the string looking for 'Leg=' but don't include in the resulting string
    # Useful web tool for regression is Rubular.com
    
    sponsorOutputPos = outputPos
    sponsor = browser.div(:id, 'billsponsordiv').a
    sponsorTmp = sponsor.attribute_value("href").scan(/(?<=Leg=).*/)[0]
    # the HREF has %20 as the space.  Replace %20 with ' '
    sponsorTmp = sponsorTmp.gsub("%20", ' ')
    output[sponsorOutputPos,3] = '=hyperlink("' + sponsor.attribute_value("href") + '";"' + sponsorTmp + '")'
    
    
    # TODO: This div exists all the time.  Need to see if the floor Sponsor is empty text
    if browser.div(:id, 'floorsponsordiv').exists?
      sponsorOutputPos = sponsorOutputPos + 1
      sponsor = browser.div(:id, 'floorsponsordiv').a
      sponsorTmp = sponsor.attribute_value("href").scan(/(?<=Leg=).*/)[0]
      # the HREF has %20 as the space.  Replace %20 with ' '
      sponsorTmp = sponsorTmp.gsub("%20", ' ')
      output[sponsorOutputPos,3] = '=hyperlink("' + sponsor.attribute_value("href") + '";"' + sponsorTmp + '")'   
    end
    
    # Short Description #
    # Place description in output work-sheet
    output[outputPos,4] = bill[1]
    # Classification #
    # Place classification in output work-sheet
    output[outputPos,5] = bill[2]

    # Click on the History tab to have the Votes be displays so we can collect the data.
    # Silly HTML behavior, not visible by human the program can't find the text either
    browser.a(:id, "activator-billStatus").click
    
    # Vote Locations #
    locations = browser.div(:id, 'billStatus').table
    votesText = ''
    tmpLocationPos = outputPos
    for i in 2..locations.rows.length-1
      if !locations.rows[i][3].text.empty?
        if locations.rows[i][3].text != "Voice vote"
          cm = searchCommittees(committee,locations.rows[i][2].text)
          # Using link.text to fix when one cell has duplicate entries
          # Place vote location, text and link in output work-sheet
          output[tmpLocationPos,6] = '=hyperlink("' +
              locations.rows[i][3].link.attribute_value("href") + '";"' + cm + ' ' + locations.rows[i][3].link.text + '")'
          tmpLocationPos = tmpLocationPos + 1
        end
      end
    end

    # Current location #
    # Known Error - if current location is the same as the last vote then
    # there will be duplicate entries in the vote location column

    # Place current location in output work-sheet
    currentLoc = searchCommittees(committee, locations.rows[i][2].text)
    
    # Place reading calendar information in the output work-sheet
    if (readingBills.has_key?(billNum))
      readingArray = readingBills[billNum]
      currentLoc = currentLoc + ' (' + readingArray[0].to_s + ')'
      
      if (readingArray[1] == "circled")
        str = readingArray[1]
        currentLoc = currentLoc + '(' + str + ')'  
      end
    end
    
    # BUG - This is a terrible search, the Clerk of the House is used frequently and for many different actions
    # NOTE - Search again if the action for the clerk location was a veto
    # NOTE - If the senate ever has a veto then their location information can be added here
    #if (currentLoc == 'Clerk of the House')
    #  currentLoc = searchCommittees(committee, 'Governor Vetoed')
    #end

    # Place current location in output work-sheet
    output[tmpLocationPos,6] = currentLoc
     
    # LoWV Position #
    # Place position text in output work-sheet
    output[outputPos,7] = bill[3]
    # save the work-sheet so we can call num_rows and get the current populated row count
    output.save
    
    #Increment position counters
    output.max_rows = output.max_rows + (output.num_rows - outputPos)
    outputPos = output.num_rows + 1

    # End of bills loop
  end
    
    output.max_rows = output.num_rows
    output.save
    browser.quit
    
    puts "Complete"
end
