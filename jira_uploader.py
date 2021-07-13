# pip install pyqt5

from PyQt5.QtWidgets import *
from PyQt5 import uic, QtGui
from PyQt5.QtCore import Qt, QThread, pyqtSlot, pyqtSignal
from PyQt5.QtGui import QIcon
import csv
from jira import JIRA
import logging
from datetime import datetime
import os
import sys
import time

# PREFIX
URL = 'http://jira.lxsemicon.com/'

DISPLAY_LOG_IN_TERMNINAL = True

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


project_list = { "SW08009_dxtest_DV2_regression_PDM":"SWDXDRP"}
table_fiedls = ["project" , "Issue Type", "Label", "Summary", "Component/s", "Assignee", "expected", "Description", "Reporter"]
issue_fiedls = ["project" , "issuetype", "labels", "summary", "components", "assignee", "customfield_10836", "description", "reporter"]

class MainDialog(QDialog):
    def __init__(self, fn=None ,parent=None):
        # Display minimize, close button
        super(MainDialog, self).__init__(parent, flags=Qt.WindowMinimizeButtonHint |Qt.WindowCloseButtonHint)
        self.initUI()
        self.jira = None

    def initUI(self):
        uic.loadUi("./JIRA_uploader.ui", self)
        self.btn_login.clicked.connect(self.connect_jira)
        self.btn_open.clicked.connect(self.open_file)
        self.btn_create_jira_issue.clicked.connect(self.create_jira_issue)
        # For JIRA Login
        self.combo_project.addItem('SW08009_dxtest_DV2_regression_PDM')

    ###########################################################################################
    # Table widget
        self.tableWidget.setRowCount(1)
        self.tableWidget.setColumnCount(9)
        self.tableWidget.setEditTriggers(QAbstractItemView.DoubleClicked)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        for j, col in enumerate(table_fiedls):
            self.tableWidget.setItem(0, j, QTableWidgetItem(col))

    def create_jira_issue(self):
        try:
            data = []
            for row in range(self.tableWidget.rowCount()-1):
                item = []
                for column in range(self.tableWidget.columnCount()):
                    if self.tableWidget.item(row+1, column).text() == None:
                        item.append("")
                    else:
                        item.append(self.tableWidget.item(row+1, column).text())
                upload_item = self.parse_csv(item)
                data.append(upload_item)
            if self.jira == None:
                self.add_log('Please connect Jira!')
            else:
                for idx, item in enumerate(data):
                    issues = self.jira.create_issue(fields=item)
                    self.add_log(f"{idx+1} item Upload SUCCESS")
        except Exception as e:
            print('--> Exception is "%s" (Line: %s)' % (e, sys.exc_info()[-1].tb_lineno))
            self.add_log('--> Exception is "%s" (Line: %s)' % (e, sys.exc_info()[-1].tb_lineno))

    def trans_jira_type(self, line):
        data = []
        project_key = self.combo_project.currentText()
        data.append(project_list[project_key])
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
        with open(name, 'r', encoding="utf-8") as f:
            rdf = csv.reader(f)
            upload_datas = []
            for i, line in enumerate(rdf):
                if i == 0 or line[9] == "Closed":
                    continue
                d = self.trans_jira_type(line)
                upload_datas.append(d)
            return upload_datas

    def add_table(self, idx, data):
        for j, value in enumerate(data):
            self.tableWidget.setItem(idx + 1, j, QTableWidgetItem(value))

    def open_file(self):
        try:
            fname = QFileDialog.getOpenFileName(self, 'Open file', './', 'CSV File(*.csv);;All file(*)')
            if fname[0]:
                self.line_edit_csv.setText(fname[0])
                datas = self.open_csv(fname[0])
                self.tableWidget.setRowCount(len(datas)+1)
                for idx, data in enumerate(datas):
                    self.add_table(idx, data)

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
    @pyqtSlot(int)
    def count(self, count):
        self.add_log('ProgressBar Value: %s' % count)
        self.progressBar.setValue(count)

    @pyqtSlot()
    def thread_is_stopped(self):
        self.set_enable_buttons(True)

    @pyqtSlot(str)
    def add_log(self, message):
        now = datetime.now()
        now = now.strftime("%H:%M:%S")
        log_message = '[%s]: %s' % (now, message)
        self.tb_log.append(log_message)
        logger.info(message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MainDialog()
    myWindow.show()
    app.exec()
