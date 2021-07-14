# pip install pyqt5
from PyQt5.QtWidgets import *
from PyQt5 import uic, QtGui
from PyQt5.QtCore import Qt, QThread, pyqtSlot, pyqtSignal
import csv
from jira import JIRA
import logging
from datetime import datetime
import os
import sys


# PREFIX
URL = 'http://jira.lxsemicon.com/'
QT_UI = "./JIRA_uploader.ui"
project_file_name = "project.ini"
DISPLAY_LOG_IN_TERMNINAL = True

project_list = { "SW08009_dxtest_DV2_regression_PDM":"SWDXDRP"}
table_fiedls = ["project" , "Issue Type", "Label", "Summary", "Component/s", "Assignee", "expected", "Description", "Reporter"]
issue_fiedls = ["project" , "issuetype", "labels", "summary", "components", "assignee", "customfield_10836", "description", "reporter"]

# logger
logger = logging.getLogger('MyLogger')
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s (%(funcName)20s:%(lineno)4d) [%(levelname)s]: %(message)s')

# Print in terminal
if DISPLAY_LOG_IN_TERMNINAL:
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

# Write in file
today = datetime.now()
today = today.strftime('%Y_%m_%d')
filename = '%s.log' % today

# If file exist, remove it
if os.path.isfile(filename):
    os.remove(filename)

file_handler = logging.FileHandler(filename)
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)

class MainDialog(QDialog):
    def __init__(self, fn=None ,parent=None):
        # Display minimize, close button
        super(MainDialog, self).__init__(parent, flags=Qt.WindowMinimizeButtonHint |Qt.WindowCloseButtonHint)
        self.initUI()
        self.jira = None
        self.tableWidgetInit()
        self.jira_create_thread = None
        self.project_items = []
        self.show()

    def initUI(self):
        uic.loadUi(QT_UI, self)
        self.btn_login.clicked.connect(self.connect_jira)
        self.btn_open.clicked.connect(self.open_file)
        self.btn_create_jira_issue.clicked.connect(self.create_jira_issue)
        # For JIRA Login

        if os.path.isfile(project_file_name):
            with open(project_file_name, 'r') as f:
                self.project_items = f.read().split("\n")
                self.add_log(f'{project_file_name} load success.')
        else:
            self.add_log(f'{project_file_name} load fail please check {project_file_name}.')
        for item in self.project_items:
            self.combo_project.addItem(item.strip())

    def tableWidgetInit(self):
        self.tableWidget.setColumnCount(9)
        self.tableWidget.setHorizontalHeaderLabels(table_fiedls)
        self.tableWidget.setRowCount(1)
        self.tableWidget.setEditTriggers(QAbstractItemView.DoubleClicked)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

    def create_jira_issue(self):
        try:
            data = []
            for row in range(self.tableWidget.rowCount()):
                item = []
                for column in range(self.tableWidget.columnCount()):
                    if self.tableWidget.item(row, column).text() == None:
                        item.append("")
                    else:
                        item.append(self.tableWidget.item(row, column).text())
                upload_item = self.parse_csv(item)
                data.append(upload_item)
            if self.jira == None:
                self.add_log('Please connect Jira!')
            else:
                self.jira_create_thread = CreateThread(self.jira, data)
                self.jira_create_thread.logSignal.connect(self.add_log)
                self.jira_create_thread.finished.connect(self.finish_thread)
                self.jira_create_thread.start()
        except Exception as e:
            self.add_log('--> Exception is "%s" (Line: %s)' % (e, sys.exc_info()[-1].tb_lineno))

    def trans_jira_type(self, line):
        data = []
        project_key = self.combo_project.currentText()
        data.append(project_key)
        data.append(line[0])
        data.append(line[1])
        data.append(line[2])
        data.append(line[4])
        data.append(line[6])
        data.append(line[12].replace(".","-"))
        data.append(line[14])
        data.append(line[15])
        return data

    def parse_csv(self, upload_item):
        data = dict()
        data["project"] = upload_item[0]
        data["issuetype"] = {'name': upload_item[1]}
        data["labels"] = [upload_item[2], ]
        data["summary"] = upload_item[3]
        if upload_item[4] != "":
            data["components"] = [{'name': upload_item[4]}, ]
        # data["assignee"] = {'name': upload_item[5]}
        data["assignee"] = {'name': "ms.jang"}
        data["customfield_10836"] = upload_item[6]
        data["description"] = upload_item[7]
        data["reporter"] = {'name': upload_item[8]}
        return data

    def open_csv(self, name):
        try:
            with open(name, 'r', encoding="utf-8") as f:
                rdf = csv.reader(f)
                upload_datas = []
                for i, line in enumerate(rdf):
                    if i == 0 or line[9] == "Closed":
                        continue
                    d = self.trans_jira_type(line)
                    upload_datas.append(d)
                return upload_datas
        except Exception as e:
            self.add_log('--> Exception is "%s" (Line: %s)' % (e, sys.exc_info()[-1].tb_lineno))

    def add_table(self, idx, data):
        for j, value in enumerate(data):
            self.tableWidget.setItem(idx, j, QTableWidgetItem(value))

    def open_file(self):
        try:
            fname = QFileDialog.getOpenFileName(self, 'Open file', './', 'CSV File(*.csv);;All file(*)')
            if fname[0]:
                self.line_edit_csv.setText(fname[0])
                datas = self.open_csv(fname[0])
                self.tableWidget.setRowCount(len(datas))
                for idx, data in enumerate(datas):
                    self.add_table(idx, data)
                self.btn_create_jira_issue.setEnabled(True)

        except Exception as e:
            print('--> Exception is "%s" (Line: %s)' % (e, sys.exc_info()[-1].tb_lineno))
            self.add_log('--> Exception is "%s" (Line: %s)' % (e, sys.exc_info()[-1].tb_lineno))

    def connect_jira(self):
        try:
            id = self.line_edit_id.text()
            pw = self.line_edit_pw.text()
            self.jira = JIRA(server=URL, basic_auth=(id, pw))
            self.add_log('Login SUCCESS!')
            self.label_status.setText("Connect")
            self.label_status.setFont(QtGui.QFont("궁서",14))
            self.label_status.setStyleSheet("Color : green")
        except Exception as e:
            self.add_log('Login File, Check ID or PASSWORD!')
    ###########################################################################################
    # Signal pyqtslot
    @pyqtSlot(str)
    def add_log(self, message):
        now = datetime.now()
        now = now.strftime("%H:%M:%S")
        log_message = '[%s]: %s' % (now, message)
        self.tb_log.append(log_message)
        logger.info(message)

    @pyqtSlot()
    def finish_thread(self):
        self.btn_create_jira_issue.setEnabled(False)

###########################################################################################
# Create Thread class
class CreateThread(QThread):
    logSignal = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, jira_ins, data):
        super(self.__class__, self).__init__()
        self.jira_ins = jira_ins
        self.data = data
        self.isRunning = True

    def run(self):
        try:
            self.logSignal.emit("Start JIRA issue create")
            for idx, item in enumerate(self.data):
                self.sleep(1)
                issue = self.jira_ins.create_issue(fields=item)
                self.sleep(1)
                self.logSignal.emit(f"{issue.key} is Created")
            self.logSignal.emit("Finish JIRA issue create")
            self.finished.emit()
        except Exception as e:
            print('--> Exception is "%s" (Line: %s)' % (e, sys.exc_info()[-1].tb_lineno))
            self.logSignal.emit('--> Exception is "%s" (Line: %s)' % (e, sys.exc_info()[-1].tb_lineno))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MainDialog()
    myWindow.show()
    app.exec()
