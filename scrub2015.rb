require 'rubygems'
require "google_drive"
require "watir-webdriver"

# Check the following CheckBoxes
# House 3rd Reading ('c4' => 'm1')
# Senate Bills ('c5' => 'm1')
# Senate 2nd Reading ('c14' => 'm2')
# Senate 3rd Reading ('c15' => 'm2')
# Tab on 2rd  ('c16' => 'm2')
# Tab on 3rd  ('c17' => 'm2')
# UnCheck all Display section CheckBoxes
def setupReadingCalPage (browser, baseLink, checkboxes)
  browser.goto baseLink + ':443/FloorCalendars/'
  
  checkboxes.each do |checkbox, divLoc|
    cb = browser.div(:id, divLoc).checkbox(:name, checkbox)
    cb.set
  end

  displayOpts = browser.div(:id, 'm3').checkboxes()
  displayOpts.each do |option|
    option.clear
  end  
end

# Collects the data from the reading lists
# reading list number portion of the 'class' name matches the checkbox number
def getReadingList (browser, *args)
  bills = {}
  
  args.each do |idLoc|
    list = browser.div(:id, idLoc).ol
    list.links.each_with_index do |entry, index|
      arr = []
 
      arr << index + 1 << entry.font.attribute_value('class')
      
      # In the reading table a substituded bill has a number before the bill number
      # this regex will exclude the number.
      # Any number of characters that are capital A through capital Z
      # and any number of characters after.
      # Example: 1HB 23 becomes HB 23
      bills[entry.text.partition(/[A-Z]+.*/)[1]] = arr
    end
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

if __FILE__ == $0
  # quit unless our script gets the correct number of command line arguments
  unless ARGV.length == 2
    puts "Incorrect Arguments.."
    puts "Usage: ruby scrub2014.rb <gmail user> <password>\n"
    exit
  end
  billYear = '2014'
  puts "Running program for Legislative Session " + billYear

  session = GoogleDrive.login(ARGV[0], ARGV[1])

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

  # Output FireFox web browser
  browser = Watir::Browser.new :ff

  baseLink = "http://le.utah.gov"
  
  # Get Reading Calendar information
  # If you want more reading calendars you need to 
  # add arguments to setupReadingCalPage and getReadingList functions
  ##################################
  # The hash table information
  # Key: checkbox HTML name
  # Value: The div the checkbox resides under
  setupReadingCalPage(browser, baseLink, { 'c4' => 'm1', 'c5' => 'm1', 
                                           'c14' => 'm2', 'c15' => 'm2', 'c16' => 'm2', 'c17' => 'm2'})
  
  # collect reading calendar data, return value is a hash object {bill, circled}
  readingBills = getReadingList(browser, 'divScroll4', 'divScroll5', 'divScroll14', 
                                         'divScroll15', 'divScroll16', 'divScroll17')
  
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
    sponsorTmp["%20"] = ' '
    output[sponsorOutputPos,3] = '=hyperlink("' + sponsor.attribute_value("href") + '";"' + sponsorTmp + '")'
    
    
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
    
    puts "Program complete"
end