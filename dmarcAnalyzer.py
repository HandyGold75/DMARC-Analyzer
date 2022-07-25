from os import listdir, mkdir, path, rename, system
from shutil import rmtree, unpack_archive, move, ReadError
from win32com.client import Dispatch, pywintypes
from xmltodict import parse as xmlParse
from gzip import BadGzipFile, open as gopen
from json import dumps, load
from datetime import datetime
import PySimpleGUI as sg
from subprocess import Popen


class outlook:
    def perpFolderStructure():
        """Preps the folder structure for the rest of the script.

        Args:
            workFolder (str, Global): Main outputfolder.
            domains (list, Global): Domains to check for in the subject and sort into subfolders.
        """
        if not path.exists(workFolder):
            mkdir(workFolder)

        for domain in domains:
            if not path.exists(workFolder + "\\" + domain):
                mkdir(workFolder + "\\" + domain)

            if not path.exists(workFolder + "\\" + domain + "\\Comp"):
                mkdir(workFolder + "\\" + domain + "\\Comp")

            if not path.exists(workFolder + "\\" + domain + "\\Xml"):
                mkdir(workFolder + "\\" + domain + "\\Xml")

            if not path.exists(workFolder + "\\" + domain + "\\Done"):
                mkdir(workFolder + "\\" + domain + "\\Done")

            if not path.exists(workFolder + "\\" + domain + "\\" + domain + "-report.json"):
                jsonFile = open(workFolder + "\\" + domain + "\\" + domain + "-report.json", "w")
                jsonFile.write("{}")
                jsonFile.close()

    def saveAttachments(mailbox: str):
        """Save all attachements of x messages

        Args:
            messages (win32com.client.CDispatch): Object of inbox items genereated by win32com.client
            workFolder (str, Global): Main outputfolder.
            domains (list, Global): Domains to check for in the subject and sort into subfolders.
        """
        outlook.perpFolderStructure()

        try:
            folder = Dispatch("Outlook.Application").GetNamespace("MAPI")
        except pywintypes.com_error:
            print("\nCan't connect to the Outlook client!\nMake sure the script and Outlook are not running with elevated privalages or try restarting Outlook!\n")
            exit()

        for i in mailbox.split("\\"):
            folder = folder.Folders(i)

        loadingCount = 0

        for message in folder.Items:
            creationTime = str(message.CreationTime)[:19].replace(" ", "!").replace(":", "").replace("-", "")

            for attachment in message.Attachments:
                xmlFileName = str(attachment).replace(".xml", "").replace(".zip", ".xml").replace(".gztar", ".xml").replace(".bztar", ".xml").replace(".tar", ".xml").replace(".gz", ".xml")

                for domain in domains:
                    compFolder = workFolder + "\\" + domain + "\\Comp"
                    xmlFolder = workFolder + "\\" + domain + "\\Xml"

                    if "report domain: " in message.Subject.lower() and domain.lower() in message.Subject.lower() and not path.exists(workFolder + "\\" + domain + "\\Done\\" + xmlFileName.replace(".xml", "") + "!" + str(creationTime) + ".xml"):
                        attachment.SaveAsFile(compFolder + "\\" + str(attachment))
                        attachment = str(attachment)

                        loadingCount += 1
                        gui_splash.update("Saving/ unpacking attachments: ", loadingCount)

                        if attachment.endswith(".rar") or attachment.endswith(".7z"):
                            print("\nWARNING! Rar or 7z file found, unpacking these archives are not yet supported!\n")

                        elif attachment.endswith(".zip") or attachment.endswith(".tar"):
                            unpack_archive(compFolder + "\\" + attachment, workFolder + "\\" + domain + "\\Xml\\")
                            rename(xmlFolder + "\\" + xmlFileName, xmlFolder + "\\" + xmlFileName.replace(".xml", "") + "!" + str(creationTime) + ".xml")

                        elif attachment.endswith(".gz"):
                            fileIn = gopen(compFolder + "\\" + attachment, "rt")
                            fileOut = open(xmlFolder + "\\" + xmlFileName, "w")

                            fileOut.write(fileIn.read())

                            fileIn.close()
                            fileOut.close()

                            rename(xmlFolder + "\\" + xmlFileName, xmlFolder + "\\" + xmlFileName.replace(".xml", "") + "!" + str(creationTime) + ".xml")

                        else:
                            try:
                                unpack_archive(compFolder + "\\" + attachment, workFolder + "\\" + domain + "\\Xml\\")
                                rename(xmlFolder + "\\" + xmlFileName, xmlFolder + "\\" + xmlFileName.replace(".xml", "") + "!" + str(creationTime) + ".xml")
                            except ReadError:
                                fileIn = gopen(compFolder + "\\" + attachment, "rt")
                                fileOut = open(xmlFolder + "\\" + xmlFileName, "w")

                                try:
                                    fileOut.write(fileIn.read())
                                    fileOut.close()

                                    rename(xmlFolder + "\\" + xmlFileName, xmlFolder + "\\" + xmlFileName.replace(".xml", "") + "!" + str(creationTime) + ".xml")
                                except BadGzipFile:
                                    print("\nWARNING! Unknown archive, unable to unpack archive!!\nFile: " + compFolder + "\\" + attachment)
                                    fileOut.close()

                                fileIn.close()


class reportHandel:
    def readXmlFiles():
        """Reads xml files.

        Args:
            workFolder (str, Global): Path to the main work folder. 
            domains (list, Global): List of domains being used.

        Returns:
            Dict: Dict of feedback items.
        """
        reports = []

        for domain in domains:
            xmlFiles = listdir(workFolder + "\\" + domain + "\\Xml")

            for file in xmlFiles:
                currentFile = open(workFolder + "\\" + domain + "\\Xml\\" + file, "r")
                currentFile_Dict = xmlParse(currentFile.read())
                currentFile_Dict["feedback"]["report_metadata"]["filename"] = file
                currentFile.close()

                reports.append(currentFile_Dict)

        return reports

    def formatReports(reports):
        """Format reports.

        Args:
            reports (dict): Dict of reports to log.

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

            finalReports[metaData["filename"]] = {
                "id": metaData["report_id"],
                "domain": policy["domain"],
                "receiver": metaData["org_name"],
                "date_range": metaData["date_range"]["begin"] + "/" + metaData["date_range"]["end"],
                "date_start": datetime.fromtimestamp(int(metaData["date_range"]["begin"])).strftime("%d/%m/%y %H:%M"),
                "date_end": datetime.fromtimestamp(int(metaData["date_range"]["end"])).strftime("%d/%m/%y %H:%M"),
                "records": currentRecords
            }

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
        for i, report in enumerate(reports):
            gui_splash.update("Updating main reports: ", i)

            reportData = reports[report]

            mainLogFile = open(workFolder + "\\" + reportData["domain"] + "\\" + reportData["domain"] + "-report.json", "r")
            mainLog = load(mainLogFile)
            mainLogFile.close()

            if not report in mainLog:
                mainLog[report] = reportData
                move(workFolder + "\\" + reportData["domain"] + "\\Xml\\" + report, workFolder + "\\" + reportData["domain"] + "\\Done")

                mainLogFile = open(workFolder + "\\" + reportData["domain"] + "\\" + reportData["domain"] + "-report.json", "w")
                mainLogFile.write(str(dumps(mainLog, indent=4)))
                mainLogFile.close()

        mainLog = {}

        for domain in domains:
            mainLogFile = open(workFolder + "\\" + domain + "\\" + domain + "-report.json", "r")
            mainLogTmp = load(mainLogFile)
            mainLogFile.close()

            for report in mainLogTmp:
                mainLog[report] = mainLogTmp[report]

        return mainLog

    def getSummary():
        """Makes an summary out of all the reports.

        Args:
            allReports (list, Global): List of Xml files that need to be read.

        Returns:
            Dict: Dict of reports including already logged items (Summary version).
            List: List of domains that where found in the reports.
        """
        summaryData = {}

        for report in allReports:
            domain = allReports[report]["domain"]

            if not domain in summaryData:
                summaryData[domain] = {"count": 0, "success": 0, "spf_failed": 0, "dkim_failed": 0, "reports": [], "success_files": [], "spf_files": [], "dkim_files": []}

            for recordData in allReports[report]["records"]:
                summaryData[domain]["count"] += recordData["amount"]

                if recordData["spf_check"] != "pass":
                    summaryData[domain]["spf_failed"] += recordData["amount"]

                    if not report in summaryData[domain]["spf_files"]:
                        summaryData[domain]["spf_files"].append(report)

                elif recordData["dkim_check"] != "pass":
                    summaryData[domain]["dkim_failed"] += recordData["amount"]

                    if not report in summaryData[domain]["dkim_files"]:
                        summaryData[domain]["dkim_files"].append(report)

                else:
                    summaryData[domain]["success"] += recordData["amount"]
                    summaryData[domain]["success_files"].append(report)

            summaryData[domain]["reports"].append({"file": report, "start": allReports[report]["date_start"], "end": allReports[report]["date_end"]})

        for dom in summaryData:
            summaryData[dom]["reports"].sort(key=lambda r: r["file"].split("!")[2], reverse=True)

        return summaryData


class gui_splash:
    sg.theme("DarkGrey13")

    window_loading = sg.Window("DMARC Analazer",
                               [[sg.Text("DMARC Analyzer", justification="center", font=("Helvetica", 14), pad=(0, 10), expand_x=True, expand_y=True)], [sg.Text("Loading. . .\n", justification="center", expand_x=True, expand_y=True, key="LoadingText")]],
                               size=(225, 100),
                               no_titlebar=True,
                               keep_on_top=True,
                               finalize=True)

    def setup():
        """Set up and finilize the splash loading screen.
        """
        gui_splash.window_loading.refresh()

    def update(message, count):
        """Update the splash loading screen.
        """
        gui_splash.window_loading["LoadingText"].update(value=message + "\n" + str(count))
        gui_splash.window_loading.refresh()

    def close():
        """Closes the splash loading screen.
        """
        gui_splash.window_loading.close()


class gui_func:
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

    def reloadData():
        """Removes the workfolder and reruns the script.
        
        Args:
            workFolder (str, Global): Path to the main work folder. 
        """
        while path.exists(workFolder):
            rmtree(workFolder)

        system("py \"" + __file__ + "\"")
        exit()

    def columnize(list: list):
        """Puts GUI items in GUI columns.

        Args:
            list (list): GUI items to put in columns.
            
        Returns:
            list: List of Colums for the GUI.
        """
        returnList = []

        tempList = []

        for i, item in enumerate(list):
            tempList.append([item])

            if (i + 1) % 2 == 0:
                returnList.append(sg.Column(tempList, expand_y=True, expand_x=True))
                tempList = []

        if tempList != []:
            returnList.append(sg.Column(tempList))

        return returnList


class gui_main:
    def header():
        return [sg.Text("--- Report ---", justification="center", font=("Helvetica", 14), expand_x=True)]

    def reports():
        """Makes an summary out of all the reports.

        Args:
            summaryData (Dict, Global): Dict of reports including already logged items (Summary version).
            
        Returns:
            List: List of report frames for the GUI layout.
        """
        reportFrameList = []

        for domain in summaryData:
            success_keys = []

            for file in summaryData[domain]["success_files"]:
                success_keys.append(file + "::OpenDir_" + domain + "\\Done\\" + file)

            spf_keys = []

            for file in summaryData[domain]["spf_files"]:
                spf_keys.append(file + "::OpenDir_" + domain + "\\Done\\" + file)

            dkim_keys = []

            for file in summaryData[domain]["dkim_files"]:
                dkim_keys.append(file + "::OpenDir_" + domain + "\\Done\\" + file)

            summary = [
                sg.Text("Count: " + str(summaryData[domain]["count"]), pad=(5, 10)),
                sg.Push(),
                sg.Text("Success: " + str(summaryData[domain]["success"]), right_click_menu=["", success_keys], pad=(5, 10)),
                sg.Push(),
                sg.Text("SPF Failed: " + str(summaryData[domain]["spf_failed"]), right_click_menu=["", spf_keys], pad=(5, 10)),
                sg.Push(),
                sg.Text("DKIM Failed: " + str(summaryData[domain]["dkim_failed"]), right_click_menu=["", dkim_keys], pad=(5, 10))
            ]

            buttons = [
                sg.Button("Show reports", key="ShowHide_" + domain + "_reports", tooltip="Show/ hide the reports list.", pad=(5, 10)),
                sg.Button("Json report", key="OpenFile_" + domain + "\\" + domain + "-report.json", tooltip="Open the Json report in notepad.", pad=(5, 10)),
                sg.Button("Open dir", key="OpenDir_" + domain, tooltip="Open the domain's directory.", pad=(5, 10))
            ]

            reportsList = []

            for i, report in enumerate(summaryData[domain]["reports"]):
                if i >= 99:
                    break

                reportsList.append([sg.Text("Start: " + report["start"], pad=(0, 0)), sg.Push(), sg.Text("End: " + report["end"], pad=(0, 0))])
                reportsList.append([sg.Text(report["file"], key="OpenDir_" + domain + "\\Done\\" + report["file"], enable_events=True, tooltip="Open this report in notepad.")])

            reports = [
                sg.Frame("Reports", [[sg.Column(reportsList, scrollable=True, vertical_scroll_only=True, expand_x=True, size=(None, 200), pad=(0, 0))]],
                         key="ShowHide_Item_" + domain,
                         visible=False,
                         title_location="n",
                         vertical_alignment="top",
                         expand_x=True,
                         expand_y=True,
                         pad=(1, 1)),
                sg.Image(size=(0, 0), pad=(0, 0))
            ]

            reportFrameList.append(sg.Frame(domain, [
                summary,
                buttons,
                reports,
            ], expand_x=True, expand_y=True, pad=(0, 10)))

        return gui_func.columnize(reportFrameList)

    def footer():
        """Get footer for the GUI.

        Returns:
            List: List of footer items for the GUI.
        """
        return [
            sg.Button("Show help", key="ShowHide_Help_help", pad=(10, 10), tooltip="Show/ hide the help menu."),
            sg.Button("Open dir", key="OpenDir_", pad=(10, 10), tooltip="Open the main work directory."),
            sg.Button("Show all", key="ShowHide_All_all", pad=(10, 10), tooltip="Show/ hide all reports."),
            sg.Button("Reload", key="Action_Reload", pad=(10, 10), tooltip="Reloads the data fresh from Outlook.")
        ]

    def helpmenu():
        """Get helpmenu in text format for the GUI.

        Args:
            domains (list, Global): Domains to check for in the subject and sort into subfolders.

        Returns:
            List: List with the helmenu for the GUI.
        """
        #yapf: Disable
        helpMenu = ("Before the GUI is launched dmarc reports will be pulled from the Outlook client and processed into an report per domain.\n" +
                    "An summary of every domain's report is visable in this GUI.\n\n" +
                    "The files used by this program are orginized as followed:\n\n"  +
                    "DMARC                                       -->     Main work directory.\n" +
                    "└───<Domain>                           -->     Folder per domain.\n" +
                    "    ├───Comp                             -->     Folder for pulled compiled files.\n" +
                    "    ├───Done                              -->     Folder for uncompiled files after being processed.\n" +
                    "    ├───Xml                                -->     Folder for uncompiled files.\n" +
                    "    └───<Domain>-report.json     -->     Json report of all processed files.\n\n" +
                    "Supported domains: \n\n" + str(domains).replace("[", "").replace("]", "").replace("'", "").replace(", ", "\n")
                    )
        #yapf: Enable

        return [sg.Text(helpMenu, visible=False, key="ShowHide_Item_Help"), sg.Image(size=(0, 0), pad=(0, 0))]

    def layout():
        """Set up layout of GUI.

        Returns:
            List: List with the leyout items.
        """
        sg.theme("DarkGrey13")

        layout = [gui_main.header(), gui_main.reports(), gui_main.footer(), gui_main.helpmenu()]

        return layout

    def loop():
        """Main loop, responisble for setting up, updating and keeping track (parts) of the GUI.

            Args:
                activeDomains (list, Global): All active domains that hve reports.
                workFolder (str, Global): Path to the main work folder. 
        """
        window = sg.Window("DMARC Analazer", gui_main.layout(), finalize=True)

        while True:
            event = window.read()[0]

            if event == sg.WIN_CLOSED:
                break

            if "::" in event:
                event = event.split("::")[1]

            if event.startswith("ShowHide_"):
                button, text = event.split("_")[1:]
                showHideStates[button] = not showHideStates[button]

                if showHideStates[button]:
                    window["ShowHide_" + button + "_" + text].update(text="Hide " + text)
                else:
                    window["ShowHide_" + button + "_" + text].update(text="Show " + text)

                if button != "All":
                    window["ShowHide_Item_" + button].update(visible=showHideStates[button])
                else:
                    for domain in activeDomains:
                        showHideStates[domain] = showHideStates["All"]

                        if showHideStates[domain]:
                            window["ShowHide_" + domain + "_reports"].update(text="Hide reports")
                        else:
                            window["ShowHide_" + domain + "_reports"].update(text="Show reports")

                        window["ShowHide_Item_" + domain].update(visible=showHideStates[domain])

            elif event.startswith("OpenFile_"):
                gui_func.openDir(file=event.split("_")[1])

            elif event.startswith("OpenDir_"):
                gui_func.openDir(event.split("_")[1])

            elif event.startswith("Action_"):
                action = event.split("_")[1]

                if action == "Reload":
                    window.close()
                    gui_func.reloadData()

        window.close()


if __name__ == "__main__":
    gui_splash.setup()

    workFolder = path.split(__file__)[0] + "\\DMARC"

    domains = ["mydomain.com", "mydomain.co.uk", "anotherdomain.eu"]
    outlook.saveAttachments("DMARC\\Inbox")

    allReports = reportHandel.logData(reportHandel.formatReports(reportHandel.readXmlFiles()))
    summaryData = reportHandel.getSummary()

    activeDomains = []
    for activeDomain in summaryData:
        activeDomains.append(activeDomain)

    showHideStates = {"Help": False, "All": False}
    for domain in activeDomains:
        showHideStates[domain] = False

    gui_splash.close()

    gui_main.loop()
