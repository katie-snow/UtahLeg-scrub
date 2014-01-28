require 'rubygems'
require "google_drive"
require "watir-webdriver"

# The first column(Col A) has the bill abbreviation name
# The third column(Col C) has the link to website
def searchLegRoster(ws, sponsorName)
  link = ''
  for row in 1..ws.num_rows
    if ws[row, 1] == sponsorName
      link =  ws[row, 3]
      break
    end
  end

  return link
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
  unless ARGV.length == 3
    puts "Incorrect Arguments.."
    puts "Usage: ruby scrub.rb <bill year> <gmail user> <password>\n"
    exit
  end

  puts "Running program for Legislative Session " + ARGV[0]

  session = GoogleDrive.login(ARGV[1], ARGV[2])

  # I can change this to spreadsheet_by_title.  Returns nil if not found can do a check here
  # This change would make it more human readable
  #session.spreadsheet_by_key("0AiCOGqCaWW_DdGRYZkRPWjlaeVAwQnhUTFdFcjZxZHc")
  lowvData = session.spreadsheet_by_title(ARGV[0] + ' Leg Session - LoWV Data')
  if lowvData.nil?
    puts ARGV[0] + ' Leg Session - LoWV Data spreadsheet not found!'
    exit -1
  end

  # Same here change to spreadsheet_by_title
  template = session.spreadsheet_by_title('Legislative Session - Template')
  if template.nil?
    puts 'Legislative Session - Template spreadsheet not found!'
    exit -1
  end

  output = template.duplicate(ARGV[0] + " Legislative Session - Output").worksheets[0]

  roster = lowvData.worksheet_by_title('legislators')
  committee = lowvData.worksheet_by_title('committees')
  billData = lowvData.worksheet_by_title('bills')

  # Output firefox web browser
  browser = Watir::Browser.new :ff

  baseLink = "http://le.utah.gov"

  # Don't look at the first row it has header information that is not needed by us
  bills = billData.rows.drop(1)

  # the number of extra rows before the bill information starts
  # buffering the worksheet by 10 rows
  # This is so the last bill doesn't run out of rows when entering sponsor or vote information into the worksheet
  # the max rows is only updated and save after each bill not during
  output.max_rows = output.max_rows + bills.length + 10
  output.save

  # Track where the current bill is located in the worksheet
  # start at the first empty row
  outputPos = output.num_rows + 1

  # Keeps track of where the next bill should start in the worksheet
  nextBillPos = outputPos

  # for loop through all the bills
  bills.each do |bill|
    # Open bill in broswer
    browser.goto baseLink + "/~" + ARGV[0] + "/bills/static/" + bill[0] + ".html"

    # Bill #
    # Place bill number in output worksheet
    output[outputPos,1] = bill[0]

    # if this image is not located on the webpage, then we are not at the right site
    if !browser.img(:name, "img0").exists?
      puts "WARNING: Bill '" + bill[0] + "' website was not found, check bill number."
      # Place warning in output worksheet
      output[outputPos, 2] = "WARNING: Bill '" + bill[0] + "' website was not found, check bill number."
      nextBillPos = nextBillPos + 1
      outputPos = nextBillPos

      # save worksheet
      output.save
      next
    end

    # Click on the image to have the Votes be displays so we can collect the data.
    # Silly html behavior, not visible by human the program can't find the text either
    browser.img(:name, "img0").click

    # Bill Name and Link #
    # Partition on B. or R. then any number of numbers then space.  The \. is making sure ruby doesn't try to do "#{look
    # a function call
    billTitle = browser.h3.text.partition(/(B|R)\. [0-9]* /)[2]
    puts "Scrubbing " + bill[0] + " " + billTitle + "..."

    # Bill Link - Assumption Introduced link is at position [1] and Amended link is at position [3]
    billTmp = browser.div(:id, 'content').table[0][0]
    if billTmp.text.include? "Amended"
      billLink = billTmp.a(:class => "nlink", index: 3).attribute_value("href")
    else
      billLink = billTmp.a(:class => "nlink", index: 1).attribute_value("href")
    end

    # Place bill title and link to text in output worksheet
    output[outputPos,2] = '=hyperlink("' + billLink + '";"' + billTitle + '")'

    # Sponsors and Links #
    # Regex Positive Lookbehind functionality.  Search for this if you need help adjusting this regex
    # Search backwards through the string looking for ': ' but don't include in the resulting string
    # Useful web tool for regression is Rubular.com
    # uniq removes duplicate entries
    sponsorTmp = browser.div(:id, 'content').table[0][1].text.scan(/(?<=: ).*/).uniq
    tmpSponsorOutputPos = outputPos
    sponsorTmp.each do |sponsor|
      sp = searchLegRoster(roster, sponsor)
      # Place sponsor and link to sponsor in output worksheet
      output[tmpSponsorOutputPos,3] = '=hyperlink("' + sp + '";"' + sponsor + '")'
      tmpSponsorOutputPos = tmpSponsorOutputPos + 1
    end
    nextBillPos = tmpSponsorOutputPos

    # Short Description #
    # Place description in output worksheet
    output[outputPos,4] = bill[1]
    # Classification #
    # Place classification in output worksheet
    output[outputPos,5] = bill[2]

    # Vote Locations #
    locations = browser.div(:id, 'item0').table
    votesText = ''
    tmpLocationPos = outputPos
    for i in 2..locations.rows.length-1
      if !locations.rows[i][3].text.empty?
        if locations.rows[i][3].text != "Voice vote"
          cm = searchCommittees(committee,locations.rows[i][2].text)
          # Using link.text to fix when one cell has duplicate entries
          # Place vote location, text and link in output worksheet
          output[tmpLocationPos,6] = '=hyperlink("' +
              locations.rows[i][3].link.attribute_value("href") + '";"' + cm + ' ' + locations.rows[i][3].link.text + '")'
          tmpLocationPos = tmpLocationPos + 1
        end
      end
    end

    # Current location #
    # Known Error - if current location is the same as the last vote then
    # there will be duplicate entries in the vote location column

    # Place current location in output worksheet
    currentLoc = searchCommittees(committee, locations.rows[i][2].text)
    # NOTE - Search again if the action for the clerk location was a veto
    # NOTE - If the senate ever has a veto then their location information can be added here
    if (currentLoc == 'Clerk of the House')
      currentLoc = searchCommittees(committee, 'Governor Vetoed')
    end

    # Place current location in output worksheet
    output[tmpLocationPos,6] = currentLoc

    # If there are more sponsors than votes stick with the sponsors position
    if (tmpLocationPos > tmpSponsorOutputPos)
      nextBillPos = tmpLocationPos
    end

    # LoWV Position #
    # Place position text in output worksheet
    output[outputPos,7] = bill[3]

    #Increment position counters
    output.max_rows = output.max_rows + (nextBillPos - outputPos)
    nextBillPos = nextBillPos + 1
    outputPos = nextBillPos

    # save worksheet
    output.save
    # End of bills loop
  end
    # TODO clean up the extra rows here, uncomment code
    # output.max_rows = output.num_rows
    # output.save
    browser.quit
end
