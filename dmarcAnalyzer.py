from argparse import ArgumentParser
from base64 import b64decode
from datetime import datetime, timedelta
from gzip import BadGzipFile, open as gopen
from json import dumps, load
from os import listdir, mkdir, path as osPath, remove, rename
from quopri import decodestring
from re import match
from shutil import ReadError, move, rmtree, unpack_archive
from sys import platform
from webbrowser import open as wbOpen

import PySimpleGUI as sg
from xmltodict import parse as xmlParse

if platform == "win32":
    from os import startfile

    from win32com.client import Dispatch, pywintypes  # type: ignore
else:
    from subprocess import call


class splashGUI:
    def __init__(self, template: str = ""):
        sg.theme("DarkGrey13")

        self.template = template

        layout = [
            [sg.Text("DMARC Analyzer", justification="center", font=("Helvetica", 14), pad=(0, 10), expand_x=True, expand_y=True)],
            [sg.Text("Loading. . .\n", justification="center", expand_x=True, expand_y=True, key="LoadingText")],
        ]

        self.window = sg.Window(
            "DMARC Analazer",
            layout,
            size=(300, 150),
            no_titlebar=True,
            keep_on_top=True,
            finalize=True,
        )

    def open(self):
        self.window.refresh()

    def update(self, *args, msg=None):
        if msg is None:
            msg = str(self.template)
            for i, arg in enumerate(args):
                msg = msg.replace(f"%{i}", str(arg))

        self.window["LoadingText"].update(value=msg)

        self.window.refresh()

    def close(self):
        self.window.close()


class dmarcGUI:
    def __init__(self, args, reports):
        sg.theme("DarkGrey13")

        self.workFolder = f"{osPath.split(__file__)[0]}/DMARC"
        self.reports = reports

        self.activeDomains = []
        for activeDomain in self.reports:
            self.activeDomains.append(activeDomain)

        self.visableReportsLimit = args.vr

        self.window = sg.Window("DMARC Analazer", self.layout(), finalize=True)

    def layout(self):
        def body():
            reportFrameList = []
            for domain in self.reports:
                success_keys = []
                for file in self.reports[domain]["success_files"]:
                    success_keys.append(f"{file}::Action_OpenDir_{domain}/Done/{file}")

                spf_keys = []
                for file in self.reports[domain]["spf_files"]:
                    spf_keys.append(f"{file}::Action_OpenDir_{domain}/Done/{file}")

                dkim_keys = []
                for file in self.reports[domain]["dkim_files"]:
                    dkim_keys.append(f"{file}::Action_OpenDir_{domain}/Done/{file}")

                reportsList = []
                for i, report in enumerate(self.reports[domain]["reports"]):
                    if i >= self.visableReportsLimit:
                        break
                    reportsList.append([sg.Text(f'Start: {report["start"]}', pad=(0, 0)), sg.Push(), sg.Text(f'End: {report["end"]}', pad=(0, 0))])
                    reportsList.append([sg.Text(report["file"], key=f'Action_OpenDir_{domain}/Done/{report["file"]}', enable_events=True, tooltip="Open this report in notepad.")])

                summary = [
                    sg.Text(f'Count: {self.reports[domain]["count"]}', pad=(5, 10)),
                    sg.Push(),
                    sg.Text(f'Success: {self.reports[domain]["success"]} ({((self.reports[domain]["success"] * 100) / self.reports[domain]["count"]):.2f}%)', right_click_menu=["", success_keys], pad=(5, 10)),
                    sg.Push(),
                    sg.Text(f'SPF Failed: {self.reports[domain]["spf_failed"]} ({((self.reports[domain]["spf_failed"] * 100) / self.reports[domain]["count"]):.2f}%)', right_click_menu=["", spf_keys], pad=(5, 10)),
                    sg.Push(),
                    sg.Text(f'DKIM Failed: {self.reports[domain]["dkim_failed"]} ({((self.reports[domain]["dkim_failed"] * 100) / self.reports[domain]["count"]):.2f}%)', right_click_menu=["", dkim_keys], pad=(5, 10)),
                ]
                buttons = [
                    sg.Button("Show reports", key=f"Action_ShowHide_{domain}_reports", tooltip="Show/ hide the reports list.", pad=(5, 10)),
                    sg.Button("Json report", key=f"Action_OpenFile_{domain}/{domain}-report.json", tooltip="Open the Json report in notepad.", pad=(5, 10)),
                    sg.Button("Open dir", key=f"Action_OpenDir_{domain}", tooltip="Open the domain's directory.", pad=(5, 10)),
                ]
                reports = [
                    sg.Frame(
                        "Reports",
                        [[sg.Column(reportsList, scrollable=True, vertical_scroll_only=True, expand_x=True, size=(None, 200), pad=(0, 0))]],
                        key=f"ShowHide_Item_{domain}",
                        visible=False,
                        title_location="n",
                        vertical_alignment="top",
                        expand_x=True,
                        expand_y=True,
                        pad=(1, 1),
                    ),
                    sg.Image(size=(0, 0), pad=(0, 0)),
                ]

                reportFrameList.append(sg.Frame(domain, [summary, buttons, reports], expand_x=True, expand_y=True, pad=(0, 10)))

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

        self.showHideStates = {"Help": False, "All": False}
        for domain in self.activeDomains:
            self.showHideStates[domain] = False

        header = [sg.Text("--- DMARC Report ---", justification="center", font=("Helvetica", 14), expand_x=True)]
        footer = [
            sg.Button("Show help", key="Action_ShowHide_Help_help", pad=(10, 10), tooltip="Show/ hide the help menu."),
            sg.Button("Open dir", key="Action_OpenDir_", pad=(10, 10), tooltip="Open the main work directory."),
            sg.Button("Show all", key="Action_ShowHide_All_all", pad=(10, 10), tooltip="Show/ hide all reports."),
        ]
        helpmenu = [
            sg.Text(
                (
                    """
Before the GUI is launched dmarc reports will be pulled from the Outlook client and processed into an report per domain.
An summary of every domain's report is visable in this GUI.\n
The files used by this program are orginized as followed:\n
DMARC					-->		Main work directory.
└───<Domain>				-->		Folder per domain.
    ├───Comp				-->		Folder for pulled compiled files.
    ├───Done				-->		Folder for uncompiled files after being processed.
    ├───Xml				-->		Folder for uncompiled files.
    └───<Domain>-report.json		-->		Json report of all processed files.\n"""
                ),
                visible=False,
                key="ShowHide_Item_Help",
            ),
            sg.Image(size=(0, 0), pad=(0, 0)),
        ]
        return [header, body(), footer, helpmenu]

    def openFolder(self, folder):
        wbOpen(f"{self.workFolder}/{folder}")

    def openFile(self, file):
        if platform == "win32":
            startfile(f"{self.workFolder}/{file}")
        else:
            call(("xdg-open", f"{self.workFolder}/{file}"))

    def showHide(self, button, text):
        self.showHideStates[button] = not self.showHideStates[button]
        if self.showHideStates[button]:
            self.window[f"Action_ShowHide_{button}_{text}"].update(text=f"Hide {text}")
        else:
            self.window[f"Action_ShowHide_{button}_{text}"].update(text=f"Show {text}")

        if button == "All":
            for domain in self.activeDomains:
                self.showHideStates[domain] = self.showHideStates["All"]
                if self.showHideStates[domain]:
                    self.window[f"Action_ShowHide_{domain}_reports"].update(text=f"Hide reports")
                else:
                    self.window[f"Action_ShowHide_{domain}_reports"].update(text=f"Show reports")

                self.window[f"ShowHide_Item_{domain}"].update(visible=self.showHideStates[domain])

            return None

        self.window[f"ShowHide_Item_{button}"].update(visible=self.showHideStates[button])

    def action(self, event):
        if event.split("_")[1] == "OpenDir":
            self.openFolder(event.split("_")[2])
        elif event.split("_")[1] == ("OpenFile"):
            self.openFile(event.split("_")[2])
        elif event.split("_")[1] == ("ShowHide"):
            self.showHide(event.split("_")[2], event.split("_")[3])

    def loop(self):
        while True:
            event = self.window.read()[0]
            if event == sg.WIN_CLOSED:
                break

            if "::" in event:
                event = event.split("::")[1]
            elif event.startswith("Action_"):
                self.action(event)

    def main(self):
        self.loop()
        self.window.close()


class outlookClient:
    def __init__(self):
        try:
            self.client = Dispatch("Outlook.Application").GetNamespace("MAPI")
        except pywintypes.com_error:
            raise ConnectionError("\nCan't connect to the Outlook client!\nMake sure the script and Outlook are not running with elevated privalages or try restarting Outlook!\n")

    def fetchFolder(self, mailbox):
        folder = self.client
        for dir in mailbox.split("/"):
            try:
                folder = folder.Folders(dir)
            except pywintypes.com_error:
                exit(f"\nCan't open Outlook folder {dir}!\nMake sure the script and Outlook are not running with elevated privalages, the target folder is not open in Outlook or try restarting Outlook!\n")

        return folder

    def fetchMessages(self, folder):
        return folder.Items

    def fetchAttachments(self, msg):
        return msg.Attachments

    def saveAttachment(self, att, path):
        try:
            att.SaveAsFile(path)
        except pywintypes.com_error as err:
            raise Exception(err)

    def isRead(self, msg):
        return not msg.UnRead

    def getSubject(self, msg):
        return msg.Subject

    def getDate(self, msg):
        return msg.CreationTime.timestamp()

    def getAttachmentFileName(self, att):
        return str(att.FileName)


class thunderbirdClient:
    def fetchFolder(self, mailbox):
        tbPath = osPath.expanduser("~/.thunderbird")
        for profile in listdir(tbPath):
            if osPath.isfile(f"~/.thunderbird/{profile}") or not osPath.exists(f"{tbPath}/{profile}/addons.json"):
                continue
            if not osPath.exists(f"{tbPath}/{profile}/webaccountMail/outlook.office365.com"):
                exit(f"Outlook {tbPath}/{profile}/webaccountMail/outlook.office365.com folder not found!\nOnly Outlook mailboxes trough IMAP are supported!")

            folderFile = f"{tbPath}/{profile}/webaccountMail/outlook.office365.com"
            mailboxFolders = mailbox.split("/")
            for i, dir in enumerate(mailboxFolders):
                if i + 1 < len(mailboxFolders):
                    dir = f"{dir}.sbd"
                folderFile += f"/{dir}"

                if not osPath.exists(folderFile):
                    exit(f"\nCan't open Thundebird folder {folderFile}!\nMake sure you have entered the right folder and Thunderbird is setup and synced!\n")

            if not osPath.isfile(folderFile):
                exit(f"\nCan't open Thundebird folder {folderFile}!\nThe folder seems to be improparly formated!\n")
            break
        else:
            raise exit(f"Unable to locate profile folder in {tbPath}")

        return folderFile

    def fetchMessages(self, folder):
        msgList = []
        with open(folder, "r", encoding="UTF-8") as fileR:
            msg = {}
            skipReport = False
            foundAttachment = False
            commitAttachment = False
            lastLineKey = None
            for line in fileR.readlines():
                if line.startswith("From - "):
                    if not msg == {} and not skipReport:
                        for check in ("Subject", "Date", "From", "To", "MessageId", "Attachment", "AttType", "AttDisposition"):
                            if not check in msg:
                                if "Subject" in msg:
                                    print(f"Can't reconstruct email ({check}): {msg['Subject']}")
                                else:
                                    print(f"Can't reconstruct email ({check}): <Missing Subject>")
                                break
                        else:
                            msgList.append(msg)

                    msg = {}
                    skipReport = False
                    foundAttachment = False
                    commitAttachment = False

                if "This is a copy of the headers that were received before the error" in line:
                    skipReport = True
                    msg = {}
                if skipReport:
                    continue

                line = line.replace("\n", "")
                for check in ("Content-Type: application/", "Content-Disposition: attachment;"):
                    if line.startswith(check):
                        foundAttachment = True
                        msg["Attachment"] = []

                if foundAttachment:
                    if line.startswith("--"):
                        foundAttachment = False
                        commitAttachment = False
                        continue

                    for check, checkName in (("Content-Type: ", "AttType"), ("Content-Disposition: ", "AttDisposition"), ("Content-Transfer-Encoding: ", "AttEnc")):
                        if checkName in msg:
                            continue

                        if line.startswith(check):
                            msg[checkName] = line.replace(check, "", 1)
                            break
                    else:
                        if line.startswith("\tfilename") or line.startswith(" filename"):
                            msg["AttDisposition"] += line.replace("\t", " ").replace(" ", " ")
                            continue

                        if line == "":
                            commitAttachment = not commitAttachment
                            continue
                        if commitAttachment:
                            msg["Attachment"].append(line)
                        continue

                for check, checkName in (("Date: ", "Date"), ("From: ", "From"), ("To: ", "To"), ("Subject: ", "Subject"), ("Message-ID: ", "MessageId"), ("Message-Id: ", "MessageId")):
                    if checkName in msg:
                        continue

                    if line.startswith(check):
                        line = line.replace(check, "", 1)
                        if line.startswith("=?"):
                            char, enc, txt = match(r"=\?{1}(.+)\?{1}([B|b|Q|q])\?{1}(.+)\?{1}=", line).groups()
                            line = (decodestring(txt) if enc.lower() == "q" else b64decode(txt)).decode(char)

                        lastLineKey = checkName
                        msg[checkName] = line
                        break
                else:
                    if lastLineKey is None:
                        continue
                    if line.startswith(" "):
                        msg[lastLineKey] += line
                        continue
                    lastLineKey = None

            for check in ("Date", "From", "To", "Subject", "MessageId", "Attachment", "AttType", "AttDisposition"):
                if not check in msg:
                    break
            else:
                msgList.append(msg)

        return msgList

    def fetchAttachments(self, msg):
        att = {}
        for checkVar in ("AttDisposition", "AttType"):
            for checkKey in ("filename=", "name="):
                if checkKey in msg[checkVar]:
                    att["Filename"] = msg[checkVar].split(checkKey)[-1]
                    if att["Filename"].count('"') > 1:
                        att["Filename"] = att["Filename"].split('"')[1]
                    break
            else:
                for checkKeyP1, checkKeyP2 in (("filename*0=", "filename*1="), ("filename*0*=", "filename*1*=")):
                    if checkKeyP1 in msg[checkVar] and checkKeyP2 in msg[checkVar]:
                        filenameP1 = msg[checkVar].split(checkKeyP1)[-1].split(";")[0]
                        if filenameP1.count('"') > 1:
                            filenameP1 = filenameP1.split('"')[1]

                        filenameP2 = msg[checkVar].split(checkKeyP2)[-1].split(";")[0]
                        if filenameP2.count('"') > 1:
                            filenameP2 = filenameP2.split('"')[1]

                        att["Filename"] = f"{filenameP1}{filenameP2}"
                        break

        att["Data"] = "".join(msg["Attachment"])

        return (att,)

    def saveAttachment(self, att, path):
        try:
            with open(path, "wb") as fileWB:
                fileWB.write(b64decode(att["Data"]))
        except Exception as err:
            raise Exception(err)

    def isRead(self, msg):
        return False

    def getSubject(self, msg):
        return msg["Subject"]

    def getDate(self, msg):
        date = msg["Date"].replace(" GMT", "")
        date = "".join(date.split(" (")[0])
        date = "".join(date.split(" +")[0])
        date = "".join(date.split(" -")[0])
        date = "".join(date.split(", ")[-1])
        return datetime.strptime(date, "%d %b %Y %H:%M:%S").timestamp()

    def getAttachmentFileName(self, att):
        return att["Filename"]


class reportCache:
    def __init__(self, args):
        self.workFolder = f"{osPath.split(__file__)[0]}/DMARC"
        self.gui = splashGUI(template="Saving attachments: %0%1Skipped: %2\nUnpacked attachments: %3%4Skipped: %5\nLoaded reports: %6%7Skipped: %8")

        self.domains = args.d.replace(" ", "").split(",")
        self.mailbox = args.m
        self.age = args.a
        self.ignoreRead = args.ur

        self.savedAtt = 0
        self.skippedSavedAtt = 0
        self.unpackedAtt = 0
        self.skippedUnpackedAtt = 0
        self.loadedRep = 0
        self.skippedLoadedRep = 0

        while not args.c and osPath.exists(self.workFolder):
            rmtree(self.workFolder)
        if not osPath.exists(self.workFolder):
            mkdir(self.workFolder)

        for domain in self.domains:
            for path in (f"{self.workFolder}/{domain}", f"{self.workFolder}/{domain}/Comp", f"{self.workFolder}/{domain}/Xml", f"{self.workFolder}/{domain}/Done"):
                if not osPath.exists(path):
                    mkdir(path)
            if not osPath.exists(f"{self.workFolder}/{domain}/{domain}-report.json"):
                with open(f"{self.workFolder}/{domain}/{domain}-report.json", "w") as file_W:
                    file_W.write("{}")

    def updateGUI(self):
        self.gui.update(
            self.savedAtt,
            f'{" " * (10 - len(str(self.savedAtt)))}',
            self.skippedSavedAtt,
            self.unpackedAtt,
            f'{" " * (10 - len(str(self.unpackedAtt)))}',
            self.skippedUnpackedAtt,
            self.loadedRep,
            f'{" " * (10 - len(str(self.loadedRep)))}',
            self.skippedLoadedRep,
        )

    def getUniqueName(self, subject: str, creationTime: int):
        def sanitize(obj: str):
            return obj.replace("<", "").replace(">", "").replace("=", "").replace(" ", "").replace("	", "")

        subjectSplited = str(subject).replace(":", "").split(" ")
        name = None
        for i, item in enumerate(subjectSplited):
            if item == "Report-ID":
                name = sanitize(subjectSplited[i + 1])
                break
            elif "Report-ID" in item:
                name = sanitize(subjectSplited[i].replace("Report-ID", ""))
                break
        if name is None:
            name = sanitize(str(creationTime)[:19].replace(" ", "!").replace(":", "").replace("-", ""))

        return name

    def saveAttachments(self):
        if platform == "win32":
            cl = outlookClient()
        else:
            cl = thunderbirdClient()

        folder = cl.fetchFolder(self.mailbox)
        for msg in cl.fetchMessages(folder):
            if cl.isRead(msg) and self.ignoreRead:
                self.skippedSavedAtt += 1
                self.updateGUI()
                continue
            if self.age > 0 and cl.getDate(msg) < (datetime.today() - timedelta(days=self.age)).timestamp():
                self.skippedSavedAtt += 1
                self.updateGUI()
                continue

            uniqueName = self.getUniqueName(cl.getSubject(msg), cl.getDate(msg))
            for att in cl.fetchAttachments(msg):
                attName = cl.getAttachmentFileName(att)
                if attName.count("!") < 3 and attName.endswith(".msg"):
                    print(f"Can't save attachment in email: {cl.getSubject(msg)}")
                    continue

                xmlFileName = attName.replace(".xml", "").replace(".zip", ".xml").replace(".gztar", ".xml").replace(".bztar", ".xml").replace(".tar", ".xml").replace(".gz", ".xml")
                for domain in self.domains:
                    compFolder = f"{self.workFolder}/{domain}/Comp"
                    xmlFolder = f"{self.workFolder}/{domain}/Xml"

                    if not "report domain: " in cl.getSubject(msg).lower() or not domain.lower() in cl.getSubject(msg).lower():
                        self.skippedSavedAtt += 1
                        self.updateGUI()
                        if not "report domain: " in cl.getSubject(msg).lower():
                            print(f"Can't verify if email contains a DMARC report: {cl.getSubject(msg)}")
                        continue
                    elif osPath.exists(f'{self.workFolder}/{domain}/Done/{xmlFileName.replace(".xml", "")}!{uniqueName}.xml'):
                        self.updateGUI()
                        continue

                    try:
                        cl.saveAttachment(att, f"{compFolder}/{attName}")
                    except Exception:
                        self.skippedSavedAtt += 1
                        self.updateGUI()
                        print(f'Can\'t save attachment "{attName}" in email: {cl.getSubject(msg)}')
                        continue
                    self.savedAtt += 1
                    self.updateGUI()

                    if attName.endswith(".rar") or attName.endswith(".7z"):
                        self.skippedUnpackedAtt += 1
                        self.updateGUI()
                        print(f"Rar or 7z file found, unpacking these archives is not yet supported!")
                        continue

                    elif attName.endswith(".zip") or attName.endswith(".tar"):
                        try:
                            unpack_archive(f"{compFolder}/{attName}", xmlFolder)
                        except ReadError:
                            print(f"Invalid archive, unable to unpack archive! File: {compFolder}/{attName}")
                            continue

                    elif attName.endswith(".gz"):
                        try:
                            with gopen(f"{compFolder}/{attName}", "rt") as fileIn:
                                with open(f"{xmlFolder}/{xmlFileName}", "w") as fileOut:
                                    fileOut.write(fileIn.read())
                        except BadGzipFile:
                            print(f"Invalid archive, unable to unpack archive! File: {compFolder}/{attName}")
                            continue

                    else:
                        try:
                            unpack_archive(f"{compFolder}/{attName}", xmlFolder)
                        except ReadError:
                            with gopen(f"{compFolder}/{attName}", "rt") as fileIn:
                                with open(f"{xmlFolder}/{xmlFileName}", "w") as fileOut:
                                    try:
                                        fileOut.write(fileIn.read())
                                    except BadGzipFile:
                                        print(f"Unknown archive, unable to unpack archive! File: {compFolder}/{attName}")
                                        self.skippedUnpackedAtt += 1
                                        self.updateGUI()
                                        continue

                    try:
                        rename(f"{xmlFolder}/{xmlFileName}", f'{xmlFolder}/{xmlFileName.replace(".xml", "")}!{uniqueName}.xml')
                    except FileExistsError:
                        remove(f'{xmlFolder}/{xmlFileName.replace(".xml", "")}!{uniqueName}.xml')
                        rename(f"{xmlFolder}/{xmlFileName}", f'{xmlFolder}/{xmlFileName.replace(".xml", "")}!{uniqueName}.xml')
                    remove(f"{compFolder}/{attName}")
                    self.unpackedAtt += 1
                    self.updateGUI()

    def getReports(self):
        reports = {}
        for domain in self.domains:
            xmlFiles = listdir(f"{self.workFolder}/{domain}/Xml")
            for file in xmlFiles:
                with open(f"{self.workFolder}/{domain}/Xml/{file}", "r") as fileR:
                    try:
                        fileData = xmlParse(fileR.read())
                    except Exception:
                        self.skippedLoadedRep += 1
                        self.updateGUI()
                        print(f"Unable to read file! File: {self.workFolder}/{domain}/Xml/{file}")
                        continue
                    fileData["feedback"]["report_metadata"]["filename"] = file

                metaData = fileData["feedback"]["report_metadata"]
                policy = fileData["feedback"]["policy_published"]
                records = fileData["feedback"]["record"]
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
                        "dkim_check": [],
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

                reports[metaData["filename"]] = {
                    "id": metaData["report_id"],
                    "domain": ".".join(policy["domain"].replace("co.uk", "co-uk").split(".")[-2:]).replace("co-uk", "co.uk"),
                    "receiver": metaData["org_name"],
                    "date_range": f'{metaData["date_range"]["begin"]}/{metaData["date_range"]["end"]}',
                    "date_start": datetime.fromtimestamp(int(metaData["date_range"]["begin"])).strftime("%d/%m/%y %H:%M"),
                    "date_end": datetime.fromtimestamp(int(metaData["date_range"]["end"])).strftime("%d/%m/%y %H:%M"),
                    "records": currentRecords,
                }
                self.loadedRep += 1
                self.updateGUI()

        return reports

    def saveReports(self):
        reports = self.getReports()
        for i, report in enumerate(reports):
            reportData = reports[report]
            with open(f'{self.workFolder}/{reportData["domain"]}/{reportData["domain"]}-report.json', "r") as mainLogFile:
                mainLog = load(mainLogFile)

            if not report in mainLog:
                mainLog[report] = reportData
                move(f'{self.workFolder}/{reportData["domain"]}/Xml/{report}', f'{self.workFolder}/{reportData["domain"]}/Done')
                with open(f'{self.workFolder}/{reportData["domain"]}/{reportData["domain"]}-report.json', "w") as mainLogFile:
                    mainLogFile.write(str(dumps(mainLog, indent=4)))

            self.gui.update(msg=f"Updating main reports:\n{i} / {len(reports)}")

    def getSavedReports(self):
        mainLog = {}
        i = 0
        for domain in self.domains:
            with open(f"{self.workFolder}/{domain}/{domain}-report.json", "r") as mainLogFile:
                mainLogTmp = load(mainLogFile)
            for report in mainLogTmp:
                i += 1
                mainLog[report] = mainLogTmp[report]
                self.gui.update(msg=f"Loading main reports:\n{i}")

        return mainLog

    def getSummary(self, reports):
        reportsSummary = {}
        for i, report in enumerate(reports):
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
            self.gui.update(msg=f"Generating summary:\n{i} / {len(reports)}")

        for dom in reportsSummary:
            reportsSummary[dom]["reports"].sort(key=lambda r: r["file"].split("!")[2], reverse=True)

        return reportsSummary

    def main(self):
        self.gui.open()
        self.saveAttachments()
        self.saveReports()
        reportsSummary = self.getSummary(self.getSavedReports())
        self.gui.close()
        return reportsSummary


if __name__ == "__main__":
    parser = ArgumentParser(description="Generates an interactive dmarc report pulled from dmarc reports in an Outlook mailbox.")
    parser.add_argument("-d", "-domains", default="mydomain.com,mydomain.co.uk,anotherdomain.eu", type=str, help="Specify domains to be checked, split with ','.")
    parser.add_argument("-m", "-mailbox", default="DMARC/Inbox", type=str, help="Specify mailbox where dmarc reports land in, folders can be specified with '/'.")
    parser.add_argument("-a", "-age", default=31, type=int, help="Specify how old in days reports may be, based on email receive date (31 is default; 0 to disable age filtering).")
    parser.add_argument("-ur", "-unread", action="store_true", help="Only cache unread mails (Windows only).")
    parser.add_argument("-c", "-cache", action="store_true", help="Use already cached files, note that if cached reports are outside any applied filters there still counted.")
    parser.add_argument("-vr", "-visablereports", default=100, type=int, help="Specify how many reports may show up in the GUI per domain (To many will cause the GUI to lag).")
    args = parser.parse_args()

    reports = reportCache(args).main()
    dmarcGUI(args, reports).main()
