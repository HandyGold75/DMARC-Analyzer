# DMARC reporter

This scripts pulls dmarc reports out of an Outlook mailbox (Supports shared mailboxes) and generates an visual report of it.
Reports are orgenized per domains and shows how many mails are successfull, failed on SPF or failed on DKIM.
To see the genereted report of an domain, click the Json report button.
To see the reports for an domain click the Show reports button (each report can be clicked to open it in notepad).

If no reports show up and no error is returned then no reports where found in the set mailbox for the given domains.

_Note: I'm not an developer by profession and the code is far from best practice, this is just something I use personaly that I think might be usefull for someone else, if you see any improvements I'm open for suggestions._

## Requirements

* Python 3.10 or greater (Older version not tested).
* Python Modules PySimpleGUI and xmltodict
* Windows 10/ 11.
* Outlook Client
* Microsoft Notepad (Not required to run).

## Set up

Install required modules

`py -m pip install PySimpleGUI xmltodict`

Setting you domains

* Open the script in a editor.
* At the bottom of the script replace the domains the scripts need to check.
  * Line: `domains = ["mydomain.com", "mydomain.co.uk", "anotherdomain.eu"]`
* At the bottom of the script replace the name "DMARC" with the name of the mailbox that will receive dmarc reports.
  * Line: `outlook.saveAttachments(outlook.getInboxMessages("DMARC"))`
* Save the script and its ready to run.
