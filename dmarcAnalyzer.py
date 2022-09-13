from os import listdir, mkdir, path, rename, remove, startfile
from shutil import rmtree, unpack_archive, move, ReadError
from win32com.client import Dispatch, pywintypes
from xmltodict import parse as xmlParse
from gzip import BadGzipFile, open as gopen
from json import dumps, load
from datetime import datetime, timedelta
from subprocess import Popen
import PySimpleGUI as sg
from argparse import ArgumentParser


class setup:
    workFolder = path.split(__file__)[0] + "\\DMARC"

    domains = []
    mailbox = ""
    ignoreRead = False
    age = 0

    loaded = 0
    skipped = 0

    def arg():
        parser = ArgumentParser(description="Generates an interactive dmarc report pulled from dmarc reports in an Outlook mailbox.")
        parser.add_argument("-d", "-domains", default="mydomain.com,mydomain.co.uk,anotherdomain.eu", type=str, help="Specify domains to be checked, split with \',\'")
        parser.add_argument("-m", "-mailbox", default="DMARC\\Inbox", type=str, help="Specify mailbox where dmarc reports land in, folders can be specified with '\\'")
        parser.add_argument("-a", "-age", default=0, type=int, help="Specify how old in days reports may be, based on email receive date (already cashed reports are not removed)")
        parser.add_argument("-ur", "-unread", action="store_true", help="Only cache unread mails.")
        parser.add_argument("-r", "-remove", action="store_true", help="Remove already cached files")
        args = parser.parse_args()

        setup.domains = args.d.replace(" ", "").split(",")
        setup.mailbox = args.m
        setup.ignoreRead = args.ur
        setup.age = args.a

        while args.r and path.exists(setup.workFolder):
            rmtree(setup.workFolder)

    def perpFolderStructure():
        if not path.exists(setup.workFolder):
            mkdir(setup.workFolder)

        for domain in setup.domains:
            if not path.exists(setup.workFolder + "\\" + domain):
                mkdir(setup.workFolder + "\\" + domain)

            if not path.exists(setup.workFolder + "\\" + domain + "\\Comp"):
                mkdir(setup.workFolder + "\\" + domain + "\\Comp")

            if not path.exists(setup.workFolder + "\\" + domain + "\\Xml"):
                mkdir(setup.workFolder + "\\" + domain + "\\Xml")

            if not path.exists(setup.workFolder + "\\" + domain + "\\Done"):
                mkdir(setup.workFolder + "\\" + domain + "\\Done")

            if not path.exists(setup.workFolder + "\\" + domain + "\\" + domain + "-report.json"):
                jsonFile = open(setup.workFolder + "\\" + domain + "\\" + domain + "-report.json", "w")
                jsonFile.write("{}")
                jsonFile.close()

    def saveAttachments():
        try:
            folder = Dispatch("Outlook.Application").GetNamespace("MAPI")
        except pywintypes.com_error:
            print("\nCan't connect to the Outlook client!\nMake sure the script and Outlook are not running with elevated privalages or try restarting Outlook!\n")
            exit()

        for i in setup.mailbox.split("\\"):
            try:
                folder = folder.Folders(i)
            except pywintypes.com_error:
                print("\nCan't open Outlook folder " + i + "!\nMake sure the script and Outlook are not running with elevated privalages, the target folder is not open in Outlook or try restarting Outlook!\n")
                exit()

        for message in folder.Items:
            if not message.UnRead and setup.ignoreRead:
                setup.skipped += 1
                gui_splash.update("Saving attachments: " + str(setup.loaded) + "\nSkipped: " + str(setup.skipped))
                continue

            if setup.age > 0 and message.ReceivedTime.timestamp() < (datetime.today() - timedelta(days=setup.age)).timestamp():
                setup.skipped += 1
                gui_splash.update("Saving attachments: " + str(setup.loaded) + "\nSkipped: " + str(setup.skipped))
                continue

            subjectSplited = str(message.Subject).replace(":", "").split(" ")
            subjectIncludeReportID = False

            for i, item in enumerate(subjectSplited):
                if item == "Report-ID":
                    nameAppend = subjectSplited[i + 1].replace("<", "").replace(">", "").replace("=", "")
                    subjectIncludeReportID = True
                elif "Report-ID" in item:
                    nameAppend = subjectSplited[i].replace("Report-ID", "").replace("<", "").replace(">", "")
                    subjectIncludeReportID = True

            if not subjectIncludeReportID:
                nameAppend = str(message.CreationTime)[:19].replace(" ", "!").replace(":", "").replace("-", "")

            for attachment in message.Attachments:
                xmlFileName = str(attachment).replace(".xml", "").replace(".zip", ".xml").replace(".gztar", ".xml").replace(".bztar", ".xml").replace(".tar", ".xml").replace(".gz", ".xml")

                for domain in setup.domains:
                    compFolder = setup.workFolder + "\\" + domain + "\\Comp"
                    xmlFolder = setup.workFolder + "\\" + domain + "\\Xml"

                    if "report domain: " in message.Subject.lower() and domain.lower() in message.Subject.lower() and not path.exists(setup.workFolder + "\\" + domain + "\\Done\\" + xmlFileName.replace(".xml", "") + "!" + str(nameAppend) + ".xml"):
                        attachment.SaveAsFile(compFolder + "\\" + str(attachment))
                        attachment = str(attachment)

                        setup.loaded += 1
                        gui_splash.update("Saving attachments: " + str(setup.loaded) + "\nSkipped: " + str(setup.skipped))

                        if attachment.endswith(".rar") or attachment.endswith(".7z"):
                            print("\nWARNING! Rar or 7z file found, unpacking these archives are not yet supported!\n")
                            continue

                        elif attachment.endswith(".zip") or attachment.endswith(".tar"):
                            unpack_archive(compFolder + "\\" + attachment, setup.workFolder + "\\" + domain + "\\Xml\\")

                        elif attachment.endswith(".gz"):
                            fileIn = gopen(compFolder + "\\" + attachment, "rt")
                            fileOut = open(xmlFolder + "\\" + xmlFileName, "w")

                            fileOut.write(fileIn.read())

                            fileIn.close()
                            fileOut.close()

                        else:
                            try:
                                unpack_archive(compFolder + "\\" + attachment, setup.workFolder + "\\" + domain + "\\Xml\\")

                            except ReadError:
                                fileIn = gopen(compFolder + "\\" + attachment, "rt")
                                fileOut = open(xmlFolder + "\\" + xmlFileName, "w")

                                try:
                                    fileOut.write(fileIn.read())

                                except BadGzipFile:
                                    print("\nWARNING! Unknown archive, unable to unpack archive!!\nFile: " + compFolder + "\\" + attachment)

                                    fileIn.close()
                                    fileOut.close()

                                    continue

                                fileIn.close()
                                fileOut.close()

                        rename(xmlFolder + "\\" + xmlFileName, xmlFolder + "\\" + xmlFileName.replace(".xml", "") + "!" + str(nameAppend) + ".xml")
                        remove(compFolder + "\\" + str(attachment))

    def main():
        setup.arg()
        setup.perpFolderStructure()
        setup.saveAttachments()


class reportHandel:
    reportsAll = None
    reportsSummary = None

    def getReports():
        reports = []

        for domain in setup.domains:
            xmlFiles = listdir(setup.workFolder + "\\" + domain + "\\Xml")

            for file in xmlFiles:
                currentFile = open(setup.workFolder + "\\" + domain + "\\Xml\\" + file, "r")
                currentFile_Dict = xmlParse(currentFile.read())
                currentFile_Dict["feedback"]["report_metadata"]["filename"] = file
                currentFile.close()

                reports.append(currentFile_Dict)

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
                    "spf_domain": [],
                    "spf_check": [],
                    "dkim_domain": [],
                    "dkim_check": []
                }

                if "spf" in record["auth_results"]:
                    if not type(record["auth_results"]["spf"]) is list:
                        currentRecord["spf_domain"] = [record["auth_results"]["spf"]["domain"]]
                        currentRecord["spf_check"] = [record["auth_results"]["spf"]["result"]]
                    else:
                        for check in record["auth_results"]["spf"]:
                            currentRecord["spf_domain"].append(record["auth_results"]["spf"]["domain"])
                            currentRecord["spf_check"].append(record["auth_results"]["spf"]["result"])

                if "dkim" in record["auth_results"]:
                    if not type(record["auth_results"]["dkim"]) is list:
                        currentRecord["dkim_domain"] = [record["auth_results"]["dkim"]["domain"]]
                        currentRecord["dkim_check"] = [record["auth_results"]["dkim"]["result"]]
                    else:
                        for check in record["auth_results"]["dkim"]:
                            currentRecord["dkim_domain"].append(check["domain"])
                            currentRecord["dkim_check"].append(check["result"])

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

    def logReports(reports):
        for i, report in enumerate(reports):
            gui_splash.update("Updating main reports:\n" + str(i))

            reportData = reports[report]

            mainLogFile = open(setup.workFolder + "\\" + reportData["domain"] + "\\" + reportData["domain"] + "-report.json", "r")
            mainLog = load(mainLogFile)
            mainLogFile.close()

            if not report in mainLog:
                mainLog[report] = reportData
                move(setup.workFolder + "\\" + reportData["domain"] + "\\Xml\\" + report, setup.workFolder + "\\" + reportData["domain"] + "\\Done")

                mainLogFile = open(setup.workFolder + "\\" + reportData["domain"] + "\\" + reportData["domain"] + "-report.json", "w")
                mainLogFile.write(str(dumps(mainLog, indent=4)))
                mainLogFile.close()

        mainLog = {}

        for domain in setup.domains:
            mainLogFile = open(setup.workFolder + "\\" + domain + "\\" + domain + "-report.json", "r")
            mainLogTmp = load(mainLogFile)
            mainLogFile.close()

            for report in mainLogTmp:
                mainLog[report] = mainLogTmp[report]

        return mainLog

    def summarize(reports):
        reportsSummary = {}

        for report in reports:
            domain = reports[report]["domain"]

            if not domain in reportsSummary:
                reportsSummary[domain] = {"count": 0, "success": 0, "spf_failed": 0, "dkim_failed": 0, "reports": [], "success_files": [], "spf_files": [], "dkim_files": []}

            for recordData in reports[report]["records"]:
                reportsSummary[domain]["count"] += recordData["amount"]

                try:
                    if recordData["spf_check"] == [] and recordData["evaluated_spf"] != "pass":
                        reportsSummary[domain]["spf_failed"] += recordData["amount"]

                        if not report in reportsSummary[domain]["spf_files"]:
                            reportsSummary[domain]["spf_files"].append(report)

                        continue

                    for spfCheck in recordData["spf_check"]:
                        if spfCheck != "pass":
                            reportsSummary[domain]["spf_failed"] += recordData["amount"]

                            if not report in reportsSummary[domain]["spf_files"]:
                                reportsSummary[domain]["spf_files"].append(report)

                            raise Exception

                    if recordData["dkim_check"] == [] and recordData["evaluated_dkim"] != "pass":
                        reportsSummary[domain]["dkim_failed"] += recordData["amount"]

                        if not report in reportsSummary[domain]["dkim_files"]:
                            reportsSummary[domain]["dkim_files"].append(report)

                        continue

                    for dkimCheck in recordData["dkim_check"]:
                        if dkimCheck != "pass":
                            reportsSummary[domain]["dkim_failed"] += recordData["amount"]

                            if not report in reportsSummary[domain]["dkim_files"]:
                                reportsSummary[domain]["dkim_files"].append(report)

                            raise Exception

                except Exception:
                    continue

                reportsSummary[domain]["success"] += recordData["amount"]
                reportsSummary[domain]["success_files"].append(report)

            reportsSummary[domain]["reports"].append({"file": report, "start": reports[report]["date_start"], "end": reports[report]["date_end"]})

        for dom in reportsSummary:
            reportsSummary[dom]["reports"].sort(key=lambda r: r["file"].split("!")[2], reverse=True)

        return reportsSummary

    def main():
        reportHandel.reportsAll = reportHandel.logReports(reportHandel.getReports())
        reportHandel.reportsSummary = reportHandel.summarize(reportHandel.reportsAll)


class gui_splash:
    sg.theme("DarkGrey13")

    window_loading = sg.Window("DMARC Analazer",
                               [[sg.Text("DMARC Analyzer", justification="center", font=("Helvetica", 14), pad=(0, 10), expand_x=True, expand_y=True)], [sg.Text("Loading. . .\n", justification="center", expand_x=True, expand_y=True, key="LoadingText")]],
                               size=(200, 100),
                               no_titlebar=True,
                               keep_on_top=True,
                               finalize=True)

    def setup():
        gui_splash.window_loading.refresh()

    def update(message):
        gui_splash.window_loading["LoadingText"].update(value=message)
        gui_splash.window_loading.refresh()

    def close():
        gui_splash.window_loading.close()


class gui_main:
    sg.theme("DarkGrey13")

    showHideStates = {"Help": False, "All": False}
    activeDomains = []

    window = None

    def lay_reports():
        reportFrameList = []

        for domain in reportHandel.reportsSummary:
            success_keys = []

            for file in reportHandel.reportsSummary[domain]["success_files"]:
                success_keys.append(file + "::Action_OpenDir_" + domain + "\\Done\\" + file)

            spf_keys = []

            for file in reportHandel.reportsSummary[domain]["spf_files"]:
                spf_keys.append(file + "::Action_OpenDir_" + domain + "\\Done\\" + file)

            dkim_keys = []

            for file in reportHandel.reportsSummary[domain]["dkim_files"]:
                dkim_keys.append(file + "::Action_OpenDir_" + domain + "\\Done\\" + file)

            summary = [
                sg.Text("Count: " + str(reportHandel.reportsSummary[domain]["count"]), pad=(5, 10)),
                sg.Push(),
                sg.Text("Success: " + str(reportHandel.reportsSummary[domain]["success"]) + " (" + str(round((reportHandel.reportsSummary[domain]["success"] * 100) / reportHandel.reportsSummary[domain]["count"], 2)) + "%)",
                        right_click_menu=["", success_keys],
                        pad=(5, 10)),
                sg.Push(),
                sg.Text("SPF Failed: " + str(reportHandel.reportsSummary[domain]["spf_failed"]) + " (" + str(round((reportHandel.reportsSummary[domain]["spf_failed"] * 100) / reportHandel.reportsSummary[domain]["count"], 2)) + "%)",
                        right_click_menu=["", spf_keys],
                        pad=(5, 10)),
                sg.Push(),
                sg.Text("DKIM Failed: " + str(reportHandel.reportsSummary[domain]["dkim_failed"]) + " (" + str(round((reportHandel.reportsSummary[domain]["dkim_failed"] * 100) / reportHandel.reportsSummary[domain]["count"], 2)) + "%)",
                        right_click_menu=["", dkim_keys],
                        pad=(5, 10))
            ]

            buttons = [
                sg.Button("Show reports", key="Action_ShowHide_" + domain + "_reports", tooltip="Show/ hide the reports list.", pad=(5, 10)),
                sg.Button("Json report", key="Action_OpenFile_" + domain + "\\" + domain + "-report.json", tooltip="Open the Json report in notepad.", pad=(5, 10)),
                sg.Button("Open dir", key="Action_OpenDir_" + domain, tooltip="Open the domain's directory.", pad=(5, 10))
            ]

            reportsList = []

            for i, report in enumerate(reportHandel.reportsSummary[domain]["reports"]):
                if i >= 99:
                    break

                reportsList.append([sg.Text("Start: " + report["start"], pad=(0, 0)), sg.Push(), sg.Text("End: " + report["end"], pad=(0, 0))])
                reportsList.append([sg.Text(report["file"], key="Action_OpenDir_" + domain + "\\Done\\" + report["file"], enable_events=True, tooltip="Open this report in notepad.")])

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

        returnList = []
        tempList = []

        for i, item in enumerate(reportFrameList):
            tempList.append([item])

            if (i + 1) % 2 == 0:
                returnList.append(sg.Column(tempList, expand_y=True, expand_x=True))
                tempList = []

        if tempList != []:
            returnList.append(sg.Column(tempList))

        return returnList

    def layout():
        header = [sg.Text("--- DMARC Report ---", justification="center", font=("Helvetica", 14), expand_x=True)]

        footer = [
            sg.Button("Show help", key="Action_ShowHide_Help_help", pad=(10, 10), tooltip="Show/ hide the help menu."),
            sg.Button("Open dir", key="Action_OpenDir_", pad=(10, 10), tooltip="Open the main work directory."),
            sg.Button("Show all", key="Action_ShowHide_All_all", pad=(10, 10), tooltip="Show/ hide all reports."),
        ]

        #yapf: Disable
        helpmenu = [sg.Text(("Before the GUI is launched dmarc reports will be pulled from the Outlook client and processed into an report per domain.\n" +
                            "An summary of every domain's report is visable in this GUI.\n\n" +
                            "The files used by this program are orginized as followed:\n\n"  +
                            "DMARC                                       -->     Main work directory.\n" +
                            "└───<Domain>                           -->     Folder per domain.\n" +
                            "    ├───Comp                             -->     Folder for pulled compiled files.\n" +
                            "    ├───Done                              -->     Folder for uncompiled files after being processed.\n" +
                            "    ├───Xml                                -->     Folder for uncompiled files.\n" +
                            "    └───<Domain>-report.json     -->     Json report of all processed files.\n\n"
                            ),visible=False, key="ShowHide_Item_Help"), sg.Image(size=(0, 0), pad=(0, 0))]
        #yapf: Enable

        return [header, gui_main.lay_reports(), footer, helpmenu]

    def act(event):
        if event.split("_")[1] == "OpenDir":
            gui_main.act_open(event.split("_")[2])

        elif event.split("_")[1] == ("OpenFile"):
            gui_main.act_open(file=event.split("_")[2])

        elif event.split("_")[1] == ("ShowHide"):
            gui_main.act_showHide(event.split("_")[2], event.split("_")[3])

    def act_open(folder=None, file=None):
        if not folder is None:
            Popen(r"explorer " + setup.workFolder + "\\" + folder)

        elif not file is None:
            startfile(setup.workFolder + "\\" + file)

    def act_showHide(button, text):
        gui_main.showHideStates[button] = not gui_main.showHideStates[button]

        if gui_main.showHideStates[button]:
            gui_main.window["Action_ShowHide_" + button + "_" + text].update(text="Hide " + text)
        else:
            gui_main.window["Action_ShowHide_" + button + "_" + text].update(text="Show " + text)

        if button == "All":
            for domain in gui_main.activeDomains:
                gui_main.showHideStates[domain] = gui_main.showHideStates["All"]

                if gui_main.showHideStates[domain]:
                    gui_main.window["Action_ShowHide_" + domain + "_reports"].update(text="Hide reports")
                else:
                    gui_main.window["Action_ShowHide_" + domain + "_reports"].update(text="Show reports")

                gui_main.window["ShowHide_Item_" + domain].update(visible=gui_main.showHideStates[domain])

            return None

        gui_main.window["ShowHide_Item_" + button].update(visible=gui_main.showHideStates[button])

    def loop():
        while True:
            event = gui_main.window.read()[0]

            if event == sg.WIN_CLOSED:
                break

            if "::" in event:
                event = event.split("::")[1]

            elif event.startswith("Action_"):
                gui_main.act(event)

    def main():
        for activeDomain in reportHandel.reportsSummary:
            gui_main.activeDomains.append(activeDomain)

        for domain in gui_main.activeDomains:
            gui_main.showHideStates[domain] = False

        gui_main.window = sg.Window("DMARC Analazer", gui_main.layout(), finalize=True)

        gui_main.loop()
        gui_main.window.close()


if __name__ == "__main__":
    gui_splash.setup()
    setup.main()
    reportHandel.main()
    gui_splash.close()

    gui_main.main()
