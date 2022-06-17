# DMARC reporter V1.3.1

This scripts pulls dmarc reports out of an Outlook mailbox (Supports shared mailboxes) and generates an visual report of it.
Reports are orgenized per domains and shows how many mails are successfull, failed on SPF or failed on DKIM.
To see the genereted report of an domain, click the Json report button.
To see the reports for an domain click the Show reports button (each report can be clicked to open it in notepad).

If no reports show up and no error is returned then no reports where found in the set mailbox folder for the given domains.

_Note: I'm not an developer by profession and the code is far from best practice, this is just something I use personaly that I think might be usefull for someone else, if you see any improvements I'm open for suggestions._

## Requirements

* Python 3.10 or greater (Older version not tested).
* Python Modules pywin32, xmltodict and PySimpleGUI.
* Windows 10/ 11.
* Outlook Client
* (Shared) mailbox folder that receives dmarc reports.
* Microsoft Notepad (Not required to run).

## Set up

Install required modules

`py -m pip install pywin32 xmltodict PySimpleGUI`

Setting you domains and mailbox

* Open the script in a editor.
* At the bottom of the script replace the domains the scripts need to check.
  * Line: `domains = ["mydomain.com", "mydomain.co.uk", "anotherdomain.eu"]`
* At the bottom of the script replace the name "DMARC\\\\Inbox" with the name of the mailbox and path to the folder that will receive dmarc reports.
  * Line: `outlook.saveAttachments(outlook.getInboxMessages("DMARC\\Inbox"))`
* Save the script and its ready to run.

## Updates

### V1.3.1

* Added a warning message for when the script can't connect to the Outlook client.

### V1.3.0

* Reworked most of the code wich imports attachments from Outlook.
* Now supports listing of reports wich have the same name.
* This is done by appending an index number at the end of the xml file names.
* Note that if dublicate emails with the same reports are present that both reports will be counted.
* Now removes cached data on exit as the implementation can list the same report 2 times if the order of mails changes.

### V1.2.0

* Small tweaks to the GUI layout.
* Small tweaks to the code variable naming.
* Reworked some code for better orginazation.
* Readable dates now get stored in the Json file.
* Added right click menu to Success, SPF Failed and DKIM Failed to open reports that contain the assosiated items.
* Added show/ hide all reports button.

### V1.1.2

* Reports will now be sorted based on start date (newer on top).

### V1.1.1

* Tweaked GUI, now has an scrollbar for the reports with an fixed heigt.

### V1.1.0

* Reworked half the code and main application should be more responsive.
* Small tweaks to the GUI layout.
* Added an reload function, this regenerates all the cached data.
* Added functionality to easly disclose in wich Outlook folder the dmarc reports land in.

### V1.0.0

* Original release.
