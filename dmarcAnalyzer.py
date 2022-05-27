from os import listdir, mkdir, path, system
from shutil import unpack_archive, move
from win32com.client import Dispatch, CDispatch
from xmltodict import parse as xmlParse
from gzip import open as gopen
from json import dumps, load
from datetime import datetime
import PySimpleGUI as sg
from subprocess import Popen


class outlook:
    def perpFolderStructure(folder: str, subfolders: list):
        """Preps the folder structure for the rest of the script.

        Args:
            folder (str): Path to the main folder the structre gets put under.
            subfolders (list): Subfolders that need to be preped.
        """
        if not path.exists(folder):
            mkdir(folder)
        for domain in subfolders:
            if not path.exists(folder + "\\" + domain):
                mkdir(folder + "\\" + domain)
            if not path.exists(folder + "\\" + domain + "\\Comp"):
                mkdir(folder + "\\" + domain + "\\Comp")
            if not path.exists(folder + "\\" + domain + "\\Xml"):
                mkdir(folder + "\\" + domain + "\\Xml")
            if not path.exists(folder + "\\" + domain + "\\Done"):
                mkdir(folder + "\\" + domain + "\\Done")
            if not path.exists(folder + "\\" + domain + "\\" + domain + "_Report.json"):
                jsonFile = open(folder + "\\" + domain + "\\" + domain + "_report.json", "w")
                jsonFile.write("{}")
                jsonFile.close()

    def getInboxMessages(target: str):
        """Get messages in the inbox of the target

        Args:
            target (str): Name of the target mailbox (NOT MAILADRESS).

        Returns:
            win32com.client.CDispatch: Object of inbox items
        """
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.Folders(target).Folders("Inbox")
        return inbox.Items

    def saveAttachments(messages: CDispatch):
        """Save all attachements of x messages

        Args:
            messages (win32com.client.CDispatch): Object of inbox items genereated by win32com.client
            workFolder (str, Global): Main outputfolder.
            domains (list, Global): Domains to check for in the subject and sort into subfolders.
        """
        outlook.perpFolderStructure(workFolder, domains)
        for message in messages:
            subject = message.Subject
            attachments = message.Attachments
            for attachment in attachments:
                for domain in domains:
                    if "report domain: " + domain in subject.lower() and not path.exists(workFolder + "\\" + domain + "\\Done\\" + str(attachment).replace(".zip", ".xml").replace(".xml.gz", ".xml")):
                        attachment.SaveAsFile(workFolder + "\\" + domain + "\\Comp\\" + str(attachment))
                        if str(attachment).endswith(".gz"):
                            fileIn = gopen(workFolder + "\\" + domain + "\\Comp\\" + str(attachment), "rt")
                            fileOut = open(workFolder + "\\" + domain + "\\Xml\\" + str(attachment).replace(".gz", ""), "w")
                            fileOut.write(fileIn.read())
                            fileIn.close()
                            fileOut.close()
                        else:
                            unpack_archive(workFolder + "\\" + domain + "\\Comp\\" + str(attachment), workFolder + "\\" + domain + "\\Xml\\")


class reportHandel:
    def readXmlFiles():
        """Reads xml files in tasks list.

        Args:
            tasks (list): List of Xml files that need to be read.
            workFolder (str, Global): Path to the main work folder. 
            domains (list, Global): List of domains being used.

        Returns:
            Dict: Dict of feedback items.
        """
        tasks = {}
        for domain in domains:
            xmlFiles = listdir(workFolder + "\\" + domain + "\\Xml")
            for file in xmlFiles:
                tasks[file] = workFolder + "\\" + domain + "\\Xml\\" + file

        reports = []
        for task in tasks:
            currentFile = open(tasks[task], "r")
            currentFile_Dict = xmlParse(currentFile.read())
            currentFile_Dict["feedback"]["report_metadata"]["filename"] = task
            currentFile.close()
            reports.append(currentFile_Dict)

        return reports

    def formatReports(reports):
        """Format reports.

        Args:
            reports (dict): Dict of reports to log.
            workFolder (str, Global): Path to the main work folder. 
            domains (list, Global): List of domains being used.

        Returns:
            Dict: Dict of formated reports.
        """
        finalReports = {}
        for report in reports:
            metaData = report["feedback"]["report_metadata"]
            policy = report["feedback"]["policy_published"]
            records = report["feedback"]["record"]
            if type(records) is dict:
                records = [records]
            currentRecords = []
            for record in records:
                currentRecord = {
                    "source_ip": str(record["row"]["source_ip"]),
                    "source_domain": str(record["identifiers"]["header_from"]),
                    "amount": int(record["row"]["count"]),
                    "evaluated_spf": str(record["row"]["policy_evaluated"]["spf"]),
                    "evaluated_dkim": str(record["row"]["policy_evaluated"]["dkim"]),
                    "spf_domain": None,
                    "spf_check": None,
                    "dkim_domain": None,
                    "dkim_check": None
                }
                if "spf" in record["auth_results"]:
                    currentRecord["spf_domain"] = record["auth_results"]["spf"]["domain"]
                    currentRecord["spf_check"] = record["auth_results"]["spf"]["result"]
                if "dkim" in record["auth_results"]:
                    currentRecord["dkim_domain"] = record["auth_results"]["dkim"]["domain"]
                    currentRecord["dkim_check"] = record["auth_results"]["dkim"]["result"]
                currentRecords.append(currentRecord)

            #yapf: Disable
            finalReports[metaData["filename"]] = {
                "id": metaData["report_id"],
                "domain": policy["domain"],
                "receiver": metaData["org_name"],
                "date_range": metaData["date_range"]["begin"] + "/" + metaData["date_range"]["end"],
                "records": currentRecords
                }
            #yapf: Enable

        return finalReports

    def logData(reports):
        """Log reports.

        Args:
            reports (dict): Dict of reports to log.
            workFolder (str, Global): Path to the main work folder. 
            domains (list, Global): List of domains being used.

        Returns:
            Dict: Dict of reports including already logged items.
        """
        for report in reports:
            reportData = reports[report]
            mainLogFile = open(workFolder + "\\" + reportData["domain"] + "\\" + reportData["domain"] + "_report.json", "r")
            mainLog = load(mainLogFile)
            mainLogFile.close()
            mainLogFile = open(workFolder + "\\" + reportData["domain"] + "\\" + reportData["domain"] + "_report.json", "w")
            if not report in mainLog:
                mainLog[report] = reportData
                move(workFolder + "\\" + reportData["domain"] + "\\Xml\\" + report, workFolder + "\\" + reportData["domain"] + "\\Done")
            mainLogFile.write(str(dumps(mainLog, indent=4)))
            mainLogFile.close()
        mainLog = {}
        for domain in domains:
            mainLogFile = open(workFolder + "\\" + domain + "\\" + domain + "_report.json", "r")
            mainLogTmp = load(mainLogFile)
            mainLogFile.close()
            for report in mainLogTmp:
                reportData = mainLogTmp[report]
                mainLog[report] = reportData
        return mainLog


class gui:
    class func:
        def openDir(folder: bool = None, file: bool = None):
            """Open folder (explorer) or file (notepad).

            Args:
                workFolder (str, Global): Path to the main work folder. 
                folder (bool, optional): Folder to open. Defaults to None.
                file (bool, optional): File to open. Defaults to None.
            """
            if not folder is None:
                Popen(r"explorer " + workFolder + "\\" + folder)
            elif not file is None:
                system("notepad \"" + workFolder + "\\" + file)

    def summaryReports():
        """Makes an summary out of all the reports

        Args:
            allReports (list, Global): List of Xml files that need to be read.

        Returns:
            List: List of items for the GUI layout.
        """
        summaryData = {}
        for report in allReports:
            domain = allReports[report]["domain"]
            if not domain in summaryData:
                summaryData[domain] = {"count": 0, "success": 0, "spf_failed": 0, "dkim_failed": 0, "reports": []}
            for recordData in allReports[report]["records"]:
                summaryData[domain]["count"] += recordData["amount"]
                if recordData["spf_check"] != "pass":
                    summaryData[domain]["spf_failed"] += recordData["amount"]
                elif recordData["dkim_check"] != "pass":
                    summaryData[domain]["dkim_failed"] += recordData["amount"]
                else:
                    summaryData[domain]["success"] += recordData["amount"]
            summaryData[domain]["reports"].append(report)

        returnListTmp = []
        for domain in summaryData:
            countSgText = sg.Text("Count: " + str(summaryData[domain]["count"]))
            successSgText = sg.Text("Success: " + str(summaryData[domain]["success"]))
            spfFailedSgText = sg.Text("SPF Failed: " + str(summaryData[domain]["spf_failed"]))
            dkimFailedSgText = sg.Text("DKIM Failed: " + str(summaryData[domain]["dkim_failed"]))

            reportSgList = []
            for report in summaryData[domain]["reports"]:
                reportSgList.append([sg.Text(report, key="Report_" + domain + "_" + report, enable_events=True, pad=(1, 1), tooltip="Open this report in notepad.")])

            #yapf: Disable
            returnListTmp.append(
                sg.Frame(domain, [
                        [countSgText, sg.Push(), successSgText, sg.Push(), spfFailedSgText, sg.Push(), dkimFailedSgText],
                        [sg.Button("Show reports", key="ShowHide_Reports_" + domain, tooltip="Show/ hide the reports list."), sg.Button("Json report", key="JsonReport_" + domain, tooltip="Open the Json report in notepad."), sg.Button("Open dir", key="OpenDir_" + domain, tooltip="Open the domain's directory.")],
                        [sg.Frame("Reports", reportSgList, title_location="n", visible=False, key="Reports_" + domain, expand_y=True, expand_x=True, pad=(1, 1), vertical_alignment="top"), sg.Image(size=(0, 0))]
                    ],
                    expand_y=True,
                    expand_x=True,
                    pad=(0, 5)
                    )
                )
            #yapf: Enable

        returnList = []
        tempList = []
        for i, item in enumerate(returnListTmp):
            tempList.append([item])
            if (i + 1) % 2 == 0:
                returnList.append(sg.Column(tempList, expand_y=True, expand_x=True))
                tempList = []
        returnList.append(sg.Column(tempList))

        return returnList

    def layout():

        sg.theme("DarkGrey13")

        layout = [[sg.Text("--- Report ---", justification="center", expand_x=True)], gui.summaryReports()]

        footer = []
        footer.append(sg.Button("Show help", key="ShowHide_Help", pad=(10, 10), tooltip="Show/ hide the help menu."))
        footer.append(sg.Button("Open dir", key="OpenDir_", pad=(10, 10), tooltip="Open the main work directory."))
        layout.append(footer)

        #yapf: Disable
        helpMenu = ("Before the GUI is launched dmarc reports will be pulled from the Outlook client and processed into an report per domain.\n" +
                    "An summary of every domain's report is visable in this GUI.\n\n" +
                    "The files used by this program are orginized as followed:\n\n"  +
                    "DMARC                                       -->     Main work directory.\n" +
                    "└───<Domain>                           -->     Folder per domain.\n" +
                    "    ├───Comp                             -->     Folder for pulled compiled files.\n" +
                    "    ├───Done                              -->     Folder for uncompiled files after being processed.\n" +
                    "    ├───Xml                                -->     Folder for uncompiled files.\n" +
                    "    └───<Domain>_report.json     -->     Json report of all processed files.\n\n" +
                    "Supported domains: \n\n" + str(domains).replace("[", "").replace("]", "").replace("'", "").replace(", ", "\n")
                    )
        #yapf: Enable
        # for domain in domains:
        #     helpMenu += domain + "\n"
        layout.append([sg.Text(helpMenu, visible=False, key="Help_Menu"), sg.Image(size=(0, 0))])

        return layout

    def main(layout):
        window = sg.Window("DMARC Analazer", layout, finalize=True)

        showHelp = False
        showHideButtons = {}
        for domain in domains:
            showHideButtons[domain] = False

        while True:
            event = window.read()[0]
            if event == sg.WIN_CLOSED:
                break
            elif event == "ShowHide_Help":
                showHelp = not showHelp
                if showHelp:
                    window["ShowHide_Help"].update(text="Hide help")
                    window["Help_Menu"].update(visible=True)
                else:
                    window["ShowHide_Help"].update(text="Show help")
                    window["Help_Menu"].update(visible=False)
            elif event.startswith("ShowHide_Reports_"):
                for domain in domains:
                    if event == "ShowHide_Reports_" + domain:
                        showHideButtons[domain] = not showHideButtons[domain]
                        if showHideButtons[domain]:
                            window["Reports_" + domain].update(visible=True)
                            window["ShowHide_Reports_" + domain].update(text="Hide reports")
                        else:
                            window["Reports_" + domain].update(visible=False)
                            window["ShowHide_Reports_" + domain].update(text="Show reports")
            elif event.startswith("Report_"):
                gui.func.openDir(file=event.replace("Report_", "").replace("_", "\\Done\\"))
            elif event.startswith("JsonReport_"):
                gui.func.openDir(file=event.replace("JsonReport_", "") + "\\" + event.replace("JsonReport_", "") + "_report.json")
            elif event.startswith("OpenDir_"):
                gui.func.openDir(event.replace("OpenDir_", ""))

        window.close()


if __name__ == "__main__":

    workFolder = path.split(__file__)[0] + "\\DMARC"
    domains = ["mydomain.com", "mydomain.co.uk", "anotherdomain.eu"]

    outlook.saveAttachments(outlook.getInboxMessages("DMARC"))
    allReports = reportHandel.logData(reportHandel.formatReports(reportHandel.readXmlFiles()))
    gui.main(gui.layout())
