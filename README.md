# DMARC reporter V1.5.0

This script pulls DMARC reports out of an Outlook mailbox (Which supports shared mailboxes) and generates a visual report of it.
Reports are organized per domain and show how many emails are successful, failed on SPF, or failed on DKIM.
To see the generated report of a domain, click the JSON report button.
To see the reports for a domain click the Show reports button (each report can be clicked to open it in notepad).

If no reports show up and no error is returned then no reports were found in the set mailbox folder for the given domains.

_Note: I'm not a developer by profession and the code is far from best practice, this is just something I use personally that I think might be useful for someone else, if you see any improvements I'm open to suggestions._

## Requirements

* Python 3.10 or greater (Older version not tested).
* Python Modules pywin32, xmltodict, and PySimpleGUI.
* Windows 10/ 11.
* Outlook Client.
* (Shared) mailbox folder that receives DMARC reports.
* Microsoft Notepad (Not required to run).

## Set up

Install required modules

`py -m pip install pywin32 xmltodict PySimpleGUI`

Setting your domains and mailbox

* Open the script in an editor.
* At the bottom of the script replace the domains the scripts need to check.
  * Line: `domains = ["mydomain.com", "mydomain.co.uk", "anotherdomain.eu"]`
* At the bottom of the script, replace the name "DMARC\\\\Inbox" with the name of the mailbox and path to the folder that will receive DMARC reports.
  * Line: `outlook.saveAttachments(outlook.getInboxMessages("DMARC\\Inbox"))`
* Save the script and it's ready to run.

## Updates

### V1.5.0

* Reworked and organized most of the code to be easier to navigate (Yes again. . .)
* Optimized the way items are firstly cached to make this noticeable faster.
* Optimized most of the code to make launching with a filled cache a bit faster.
* Reworked the loading splash screen and it now shows the progress of caching items.
* Now the show reports option only shows the 100 most recent reports as to many reports could slow down the window.
* Note that this doesn't effect the per domain summary.
* Up to 3 domains should always run without issue, up to 6 domains with small slowdowns, and up to 8 domains while workable.
* This only applies after all reports have been shown (Hiding them doesn't help).
* Fixing this would take too much processing for every hide/ show action.

### V1.4.0

* Reworked all of the code to make everything more readable.
* Reworked most of the code that imports attachments from Outlook.
* Now no longer appends an index number.
* Now it appends the date and time of when the email was created which contained the report.
* This should fix duplicate emails from being counted 2 times.
* Now no longer removes the cache on exit.
* This is because if the order changes the emails this will no longer cause reports to be counted 2 times.
* This should speed up startup times a lot (especially when there are a lot of reports).
* Added loading screen instead of printing to the CLI.

### V1.3.1

* Added a warning message for when the script can't connect to the Outlook client.

### V1.3.0

* Reworked most of the code that imports attachments from Outlook.
* Now supports the listing of reports which have the same name.
* This is done by appending an index number at the end of the XML file names.
* Note that if duplicate emails with the same reports are present that both reports will be counted.
* Now removes cached data on exit as the implementation can list the same report 2 times if the order of mails changes.

### V1.2.0

* Small tweaks to the GUI layout.
* Small tweaks to the code variable naming.
* Reworked some code for better organization.
* Readable dates now get stored in the JSON file.
* Added right-click menu to Success, SPF Failed and DKIM Failed to open reports that contain the associated items.
* Added show/ hide all reports button.

### V1.1.2

* Reports will now be sorted based on the start date (newer on top).

### V1.1.1

* Tweaked GUI, now has a scrollbar for the reports with a fixed height.

### V1.1.0

* Reworked half the code and the main application should be more responsive.
* Small tweaks to the GUI layout.
* Added a reload function, this regenerates all the cached data.
* Added functionality to easily disclose in which Outlook folder the DMARC reports land.

### V1.0.0

* Original release.
