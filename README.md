# DMARC reporter V1.6.1

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
