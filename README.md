# DMARC reporter V0.8.0

This script pulls DMARC reports out of an Outlook mailbox (Which supports shared mailboxes) and generates a visual report of it.
Reports are organized per domain and show how many emails are successful, failed on SPF, or failed on DKIM.
To see the generated report of a domain, click the JSON report button.
To see the reports for a domain click the Show reports button (each report can be clicked to open it in notepad).

If no reports show up and no error is returned then no reports were found in the set mailbox folder for the given domains.

_Note: I'm not a developer by profession and the code is far from best practice, this is just something I use personally that I think might be useful for someone else, if you see any improvements I'm open to suggestions._

## Requirements

* Python 3.10 or greater (Older version not tested).
* Python Modules xmltodict, PySimpleGUI and pywin32 (Windows only).
* Outlook Client (Windows only).
* Thunderbird Client (Linux only).
* Mailbox folder that contains DMARC reports.

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
  * Cache: False
    * If the argument is present this will allow the use of cached reports on startup.
    * Do this action with the -c or -cache argument.
    * For eq: `py dmarcAnalyzer.py -c`
  * Visable Reports: 0
    * Specify how many reports may show up in the GUI per domain (To many will cause the GUI to lag).
    * Change with the -vr or -visablereports argument.
    * For eq: `py dmarcAnalyzer.py -vr 100`
* The default settings can be changed by modifying this code block present at the bottom of the script:

  ```python
    parser.add_argument("-d", "-domains", default="mydomain.com,mydomain.co.uk,anotherdomain.eu", type=str, help="Specify domains to be checked, split with ','.")
    parser.add_argument("-m", "-mailbox", default="DMARC/Inbox", type=str, help="Specify mailbox where dmarc reports land in, folders can be specified with '/'.")
    parser.add_argument("-a", "-age", default=31, type=int, help="Specify how old in days reports may be, based on email receive date (31 is default; 0 to disable age filtering).")
    parser.add_argument("-ur", "-unread", action="store_true", help="Only cache unread mails (Windows only).")
    parser.add_argument("-c", "-cache", action="store_true", help="Use already cached files, note that if cached reports are outside any applied filters there still counted.")
    parser.add_argument("-vr", "-visablereports", default=100, type=int, help="Specify how many reports may show up in the GUI per domain (To many will cause the GUI to lag).")
  ```

  * To modify the default settings for domains, mailbox, and age you can change the value after `default=`
  * To modify the default settings for unread and remove you can change the value of `action=` to `store_false` (default is `store_true`)

## Making source executable

Install required modules:

`py -m pip install pyinstaller`

Making executable (Linux):

1. Modify default settings in source (Refer to "[Getting started](#getting-started)")
2. Run pyInstaller: `pyinstaller -F -w --clean --distpath ./ ./dmarcAnalyzer.py`
3. Clean up temporary files: `rm -r -f ./build/ && rm -f ./dmarcAnalyzer.spec`

Making executable (Windows):

1. Modify default settings in source (Refer to "[Getting started](#getting-started)")
2. Run pyInstaller: `pyinstaller -F -w --clean --distpath .\ .\dmarcAnalyzer.py`
3. Clean up temporary files: `Remove-Item -Path ".\build\" -Recurse -Force ; Remove-Item -Path ".\dmarcAnalyzer.spec" -Force`

_Note: Startup arguments are not supported in executable format!_

## Limitations

* On Windows only Outlook can be used as mailclient.
* On Linux only Thunderbird can be used as mailclient.
* The Outlook client won't deal succesfully with .msg email attachments, will give an warning in the CLI.
* The Thunderbird client can't deal with all email from enterprise.protection.outlook.com, will not give an warning (An fix will likely be implemented sometime).
