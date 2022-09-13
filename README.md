# DMARC reporter V1.6.0

This script pulls DMARC reports out of an Outlook mailbox (Which supports shared mailboxes) and generates a visual report of it.
Reports are organized per domain and show how many emails are successful, failed on SPF, or failed on DKIM.
To see the generated report of a domain, click the JSON report button.
To see the reports for a domain click the Show reports button (each report can be clicked to open it in notepad).

If no reports show up and no error is returned then no reports were found in the set mailbox folder for the given domains.

_Note: I'm not a developer by profession and the code is far from best practice, this is just something I use personally that I think might be useful for someone else, if you see any improvements I'm open to suggestions._

## Requirements

* Python 3.10 or greater (Older version not tested).
* Python Modules pywin32, xmltodict, and PySimpleGUI.
* Outlook Client.
* (Shared) mailbox folder that receives DMARC reports.

## Getting started

Install required modules

`py -m pip install pywin32 xmltodict PySimpleGUI`

Setting your domains and mailbox

* These are the default settings:
  * Domains: mydomain.com, mydomain.co.uk, anotherdomain.eu
    * Specify which domains to look for in the DMARC reports.
    * Change with the -d or -domain argument.
    * For eq: `py dmarcAnalyzer.py -d mydomain.com,mydomain.co.uk,anotherdomain.eu`
  * Mailbox: DMARC\\\\Inbox
    * Specify in which mailbox and (sub)folders to look for DMARC reports.
    * Change with the -m or -mailbox argument.
    * For eq: `py dmarcAnalyzer.py -m DMARC\\Inbox`
  * Age: 0
    * Specify how old in days the reports may be, based on the date the email was received (already cached reports are not removed).
    * 0 or lower will disable this filter.
    * Change with the -a or -age argument.
    * For eq: `py dmarcAnalyzer.py -a 0`
  * Unread: False
    * If the argument is present only reports of unread emails will be allowed.
    * Apply this filter with -ur or -unread argument.
    * For eq: `py dmarcAnalyzer.py -ur`
  * Remove: False
    * If the argument is present this will remove all cached reports on startup.
    * Do this action with the -r or -remove argument.
    * For eq: `py dmarcAnalyzer.py -r`
* The default settings can be changed by modifying this code block present at the top of the script:

  ```python
  parser.add_argument("-d", "-domains", default="mydomain.com,mydomain.co.uk,anotherdomain.eu", type=str, help="Specify domains to be checked, split with \',\'")
  parser.add_argument("-m", "-mailbox", default="DMARC\\Inbox", type=str, help="Specify mailbox where DMARC reports land in, folders can be specified with '\\'")
  parser.add_argument("-a", "-age", default=0, type=int, help="Specify how old in days reports may be, based on email receive date (already cashed reports are not removed)")
  parser.add_argument("-ur", "-unread", action="store_true", help="Only cache unread mails.")
  parser.add_argument("-r", "-remove", action="store_true", help="Remove already cached files")
  ```

  * To modify the default settings for domains, mailbox, and age you can change the value after `default=`
  * To modify the default settings for unread and remove you can change the value of `action=` to `store_false` (default is `store_true`)

## Updates

### V1.6.0

* Stopped using system calls, now when opening reports the system's default text editor will be used.
* Made the naming of saved attachments more robust in cases illegal characters are used for filenames.
* Refactor most code to make everything more readable, now uses classes better instead of declaring too many globals.
* Improved checks that happen when interacting with the GUI (And more readable code).
* Removed docstrings as they didn't add much value.
* Removed the reload option as this doesn't handle arguments well.
* Added error message in case the script is unable to open the specified Outlook folder.
* Added support for reports that list multiple SPF and DKIM checks under the same record.
* Added support for startup arguments.
* Added support for age filtering.
* Added support for unread filtering.
* Added support for removing cached files on startup.

### V1.5.2

* Added percentages of success, SPF failed and DKIM failed.
* Reworked the way attachments are imported and the naming of saved attachments.
* Now tries to append the report ID included in the subject instead of the creation date and time.
* This prevents emails created at the same time with the same details from raising an error.
* The date and time will only be appended if the report ID couldn't be resolved as a fallback.

### V1.5.1

* Optimized the way items are firstly cached a little bit more.

### V1.5.0

* Reworked and organized most of the code to be easier to navigate (Yes again. . .).
* Optimized the way items are firstly cached to make this noticeable faster.
* Optimized most of the code to make launching with a filled cache a bit faster.
* Reworked the loading splash screen and it now shows the progress of caching items.
* Now the show reports option only shows the 100 most recent reports as to many reports could slow down the window.
* Note that this doesn't affect the per-domain summary.
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
* Now removes cached data on exit as the implementation can list the same report 2 times if the order of emails changes.

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
