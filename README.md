UtahLeg-scrub
=============

Tool for the League of Women Voters of Utah to collect data from the Utah Legistature web site

Installation
============

Windows

To install Ruby go to rubyinstaller.org and use the windows installer Ruby 1.9.3-p448.  All the of the Ruby gems are up-to-date with this install
Do the default install

To install Ruby DevKit go to rubyinstaller.org and install DevKit-tdm-32-4.5.2-20111229-1559-sfx.exe.

From the Ruby command line run the following

gem install selenium-webdriver
gem install watir-webdriver
gem install google_drive

That is everything that is needed to run the ruby script.

Execution
=========

This scrub executable is only good for the 2014 legislative year.  The Utah Legislative web developers changed the design
of the pages this year.
Place the Template and LoWV data spreadsheets in the base directory of the google drive.  
Convert them to google docs.
The template and LoWV Data spreadsheets need to be accessable through the gmail account provided

To run the script

ruby scrub2014.rb <gmail name> <gmail pass>


If the script fails with GoogleDrive::AuthenticationError then the gmail username/password failed
