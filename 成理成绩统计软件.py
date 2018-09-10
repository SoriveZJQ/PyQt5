#!/usr/bin/env python 
# -*- coding:utf-8 -*-


# 应在爬虫这一个线程中使用并发编程


import sys
import xlrd
import pymongo
import threading
import time
from selenium import webdriver
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from pyquery import PyQuery as pq
from PyQt5.QtWidgets import QApplication, QPushButton, QLabel, QFileDialog, QInputDialog, QTextBrowser, QFrame
from PyQt5.QtWidgets import QMessageBox, QLineEdit, QDialog, QProgressBar
from PyQt5.QtGui import QIcon, QPixmap, QPalette, QBrush, QMovie
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QBasicTimer
import PyQt5.sip


time_start = time.time()
time_stop = None


LOGIN_PWDS = ['201707030103']
login_url = 'http://202.115.133.173:805/Default.aspx'
score_url = 'http://202.115.133.173:805/SearchInfo/Score/ScoreList.aspx'
options = webdriver.FirefoxOptions()
options.add_argument('--headless')
firefoxprofile = FirefoxProfile()
firefoxprofile.set_preference('permissions.default.stylesheet', 2)
firefoxprofile.set_preference('permissions.default.image', 2)
firefoxprofile.set_preference('dom.ipc.plugins.enabled.libflashplayer.so', 'false')
SLEEP_TIME = 1
MAX_THREADS = 5


data = None
table = None
names = None
n = None
accounts = None
passwords = None


class LoginDialog(QDialog):
    login_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setFixedSize(350, 250)
        self.setWindowIcon(QIcon('app.ico'))
        self.setWindowTitle('登录')
        self.setWindowFlags(Qt.FramelessWindowHint)

        self.picLabel = QLabel('', self)
        self.picLabel.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        self.picLabel.setGeometry(0, 0, 350, 140)
        self.movie = QMovie('杜甫.gif')
        if self.movie.isValid():
            self.picLabel.setMovie(self.movie)
            self.picLabel.setScaledContents(True)
            self.movie.start()
        else:
            self.picLabel.setPixmap(QPixmap('.\CDUT_pic_1.jpeg'))

        self.pwdEdit = QLineEdit(self)
        self.pwdEdit.setGeometry(90, 160, 200, 30)
        self.pwdEdit.setPlaceholderText('请输入密钥')
        self.pwdEdit.setEchoMode(QLineEdit.Password)

        self.submitBtn = QPushButton('登录', self)
        self.submitBtn.setGeometry(106, 210, 60, 30)
        self.submitBtn.setStyleSheet("QPushButton{color:black}"
                                      "QPushButton:hover{color:red}"
                                      "QPushButton{background-color:lightgreen}"
                                      "QPushButton{border:2px}"
                                      "QPushButton{border-radius:15px}"
                                      "QPushButton{padding:2px 4px}")

        self.closeBtn = QPushButton('关闭', self)
        self.closeBtn.setGeometry(215, 210, 60, 30)
        self.closeBtn.setStyleSheet("QPushButton{color:black}"
                                     "QPushButton:hover{color:red}"
                                     "QPushButton{background-color:lightgreen}"
                                     "QPushButton{border:2px}"
                                     "QPushButton{border-radius:15px}"
                                     "QPushButton{padding:2px 4px}")

        self.submitBtn.clicked.connect(self.submitLogin)
        self.closeBtn.clicked.connect(self.closeLogin)

        self.show()

    def submitLogin(self):
        self.text = self.pwdEdit.text()

        if self.text not in LOGIN_PWDS:
            msgBox = QMessageBox()
            msgBox.setWindowOpacity(0.8)
            msgBox.setWindowTitle('错误')
            msgBox.setWindowIcon(QIcon('app.ico'))
            msgBox.setIcon(QMessageBox.Critical)
            msgBox.setText('密钥输入错误！')
            msgBox.setInformativeText('请联系QQ:<span style="color: red">1792575431</span>')
            msgBox.addButton('确定', QMessageBox.AcceptRole)
            msgBox.exec_()
            self.pwdEdit.clear()
            self.login_signal.emit('fail')

        else:
            self.login_signal.emit('success')
            self.close()

    def closeLogin(self):
        self.close()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.m_flag = True
            self.m_Position = event.globalPos() - self.pos()
            event.accept()

    def mouseMoveEvent(self, QMouseEvent):
        if Qt.LeftButton and self.m_flag:
            self.move(QMouseEvent.globalPos() - self.m_Position)
            QMouseEvent.accept()

    def mouseReleaseEvent(self, QMouseEvent):
        self.m_flag = False


class GetScore(QDialog):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setFixedSize(350, 600)
        self.setWindowFlags(Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)
        self.setWindowTitle('成理成绩查询')
        self.setWindowIcon(QIcon('.\\app.ico'))
        window_pale = QPalette()
        window_pale.setBrush(self.backgroundRole(), QBrush(QPixmap('.\pic3.jpg')))
        self.setPalette(window_pale)
        self.setWindowOpacity(1)
        self.setStyleSheet("QPushButton{background: transparent} QTextBrowser{background: transparent; border: 0}")

        self.aboutBtn = QPushButton('关于', self)
        self.aboutBtn.setGeometry(145, 580, 40, 25)
        self.aboutBtn.setStyleSheet('color: red; text-decoration: underline')

        self.fileLabel = QLabel('信息文件：', self)
        self.fileLabel.move(20, 20)
        self.fileValueLabel = QLabel('', self)
        self.fileValueLabel.setStyleSheet("color: blue")
        self.fileValueLabel.setGeometry(100, 15, 160, 25)
        self.fileBtn = QPushButton('...', self)
        self.fileBtn.setStyleSheet("QPushButton:hover{color: orange}")
        self.fileBtn.setGeometry(300, 15, 40, 30)

        self.termLabel = QLabel('学期：', self)
        self.termLabel.move(20, 80)
        self.termValueLabel = QLabel('', self)
        self.termValueLabel.setStyleSheet("color: blue")
        self.termValueLabel.setGeometry(100, 75, 160, 25)
        self.termBtn = QPushButton('...', self)
        self.termBtn.setStyleSheet("QPushButton:hover{color: orange}")
        self.termBtn.setGeometry(300, 75, 40, 30)

        self.saveLabel = QLabel('存储位置：', self)
        self.saveLabel.move(20, 140)
        self.saveValueLabel = QLabel('', self)
        self.saveValueLabel.setStyleSheet("color: blue")
        self.saveValueLabel.setGeometry(100, 135, 160, 25)
        self.saveBtn = QPushButton('...', self)
        self.saveBtn.setStyleSheet("QPushButton:hover{color: orange}")
        self.saveBtn.setGeometry(300, 135, 40, 30)

        self.infoLabel = QLabel('执行情况↓↓↓', self)
        self.infoLabel.move(20, 200)
        self.infoText = QTextBrowser(self)
        self.infoText.setContextMenuPolicy(Qt.NoContextMenu)
        self.infoText.move(20, 240)

        self.executeBtn = QPushButton('执行', self)
        self.executeBtn.setStyleSheet("QPushButton:hover{color: green}"
                                      "")
        self.executeBtn.setGeometry(145, 480, 40, 25)

        self.progress = QProgressBar(self)
        self.progress.setMinimum(0)
        self.progress.setMaximum(0)
        self.progress.setGeometry(65, 520, 200, 20)
        self.progress.setStyleSheet(
            "QProgressBar{border-radius: 10px; border-radius: 10px; text-align: center; color: red}"
            "QProgressBar:chunk{border: 1px solid grey; border-radius: 5px; background-color: grey;}")

        self.fileBtn.clicked.connect(self.fileopen)
        self.termBtn.clicked.connect(self.chooseterm)
        self.saveBtn.clicked.connect(self.filestore)
        self.aboutBtn.clicked.connect(self.about)
        self.executeBtn.clicked.connect(self.thread_start)

    def showOrcloseDialog(self, info):
        if info == 'success':
            if not self.isVisible():
                self.show()
        elif info == 'fail':
            if self.isVisible():
                self.close()

    def changetxt(self, file_inf):
        if file_inf == 'stop':
            self.executeBtn.setEnabled(True)
            self.infoText.append('\n耗时：%.2fs' % (time_stop - time_start))
        else:
            self.infoText.append(file_inf)
            value = self.progress.value()
            value += 1
            self.progress.setValue(value)

    def showerror(self, info):
        if info == 0:
            self.infoText.append('打开文件失败！')
        else:
            self.infoText.append('打开文件成功！')

    def thread_start(self):
        if self.fileValueLabel.text() == '':
            self.infoText.append('<span style="color: red">请选择信息文件！</span>')
            return
        elif self.termValueLabel.text() == '':
            self.infoText.append('<span style="color :red">请选择学期！</span>')
            return
        elif self.saveValueLabel.text() == '':
            self.infoText.append('<span style="color: red">请选择存储位置！</span>')
            return

        global n
        self.progress.setMaximum(n)
        self.progress.setValue(0)

        self.executeBtn.setEnabled(False)
        self.term = self.termValueLabel.text()
        self.thread_1 = Thread_get(term=self.term, file=self.fname[0])
        self.thread_1.file_changed_signal.connect(self.changetxt)
        self.thread_1.start()

    def fileopen(self):
        fname = QFileDialog.getOpenFileName(self, '打开文件', './', filter='XLS files (*.xls)')
        if fname[0]:
            filename = []
            i = -1
            while True:
                if fname[0][i] == '/':
                    break
                filename.append(fname[0][i])
                i -= 1
            filename = ''.join(filename)
            self.fileValueLabel.setText(filename[::-1])

            try:
                thread_open = Thread_open(file=fname[0])
                thread_open.info.connect(self.showerror)
                thread_open.start()
            except Exception:
                self.infoText.append('<span style="color: red">打开文件异常！</span>')

    def chooseterm(self):
        inputdialog = QInputDialog(self)
        inputdialog.setIntRange(100000, 999999)
        inputdialog.setIntValue(201701)
        inputdialog.setWindowTitle('学期')
        inputdialog.setLabelText('输入学期(如201701)：')
        inputdialog.setOkButtonText('确定')
        inputdialog.setCancelButtonText('取消')
        inputdialog.setWindowOpacity(0.7)
        ok = inputdialog.exec_()
        if ok:
            self.termValueLabel.setText(str(inputdialog.intValue()))

    def filestore(self):
        self.fname = QFileDialog.getSaveFileName(self, '保存文件', directory='./', filter='XLS files (*.xls)')
        if self.fname[0]:
            filename = []
            i = -1
            while True:
                if self.fname[0][i] == '/':
                    break
                filename.append(self.fname[0][i])
                i -= 1
            filename = ''.join(filename)
            self.saveValueLabel.setText(filename[::-1])

    def about(self):
        msgBox = QMessageBox(QMessageBox.NoIcon, '关于', '<b>钦哥出品</b>')
        msgBox.setWindowOpacity(0.7)
        msgBox.addButton('确定', QMessageBox.AcceptRole)
        msgBox.setWindowIcon(QIcon('.\\app.ico'))
        msgBox.setInformativeText('如有疑问请联系QQ:<span style="color: red">1792575431</span>')
        msgBox.setIconPixmap(QPixmap('.\My_logo.jpg'))
        msgBox.exec_()


class Thread_open(QThread):
    info = pyqtSignal(int)

    def __init__(self, file=None, parent=None):
        self.file = file
        super(Thread_open, self).__init__(parent)

    def __del__(self):
        self.wait()

    def run(self):
        try:
            global data
            global table
            global names
            global n
            global accounts
            global passwords
            data = xlrd.open_workbook(self.file)
            table = data.sheets()[0]
            names = table.col_values(0)
            names.pop(0)
            names.pop(0)
            n = len(names)
            accounts = table.col_values(1)
            accounts.pop(0)
            accounts.pop(0)
            passwords = table.col_values(5)
            passwords.pop(0)
            passwords.pop(0)
            self.info.emit(1)
        except Exception:
            self.info.emit(0)


class Thread_get(QThread):
    file_changed_signal = pyqtSignal(str)

    def __init__(self, term=None, file=None, parent=None):
        super(Thread_get, self).__init__(parent)
        global n
        global accounts
        global passwords
        global names
        self.n = n
        self.accounts = accounts
        self.passwords = passwords
        self.names = names
        self.term = term
        self.file = file
        self.errorCnt = 0

    def __del__(self):
        # self.working = False
        self.wait()

    def run(self):
        for i in range(self.n):
            try:
                self.driver = webdriver.Firefox(firefox_options=options, firefox_profile=firefoxprofile)
                self.account = self.accounts[i]
                password = self.passwords[i]
                self.driver.get(login_url)
                input_account = self.driver.find_element_by_name('txtUser')
                input_password = self.driver.find_element_by_name('txtPWD')
                button = self.driver.find_element_by_class_name('btn_login')
                input_account.send_keys(self.account)
                input_password.send_keys(password)
                button.click()
                self.driver.get(score_url)
                html = self.driver.page_source
                doc = pq(html)
                score = {}
                score['姓名'] = self.names[i]
                courses = doc('.score_right_infor_list.listUl')
                courses = courses.children()
                for item in courses.items():
                    term = item.children('.floatDiv20').text().strip()
                    if term == self.term:
                        title = item.find('div:nth-child(3)').text().strip()
                        cj = item.find('div:nth-child(6)').text().strip()
                        score[title] = cj
                self.driver.close()

                while len(score) == 1:
                    self.driver = webdriver.Firefox(firefox_options=options)
                    self.account = self.accounts[i]
                    password = self.passwords[i]
                    self.driver.get(login_url)
                    input_account = self.driver.find_element_by_name('txtUser')
                    input_password = self.driver.find_element_by_name('txtPWD')
                    button = self.driver.find_element_by_class_name('btn_login')
                    input_account.send_keys(self.account)
                    input_password.send_keys(password)
                    button.click()
                    self.driver.get(score_url)
                    html = self.driver.page_source
                    doc = pq(html)
                    score = {}
                    score['姓名'] = self.names[i]
                    courses = doc('.score_right_infor_list.listUl')
                    courses = courses.children()
                    for item in courses.items():
                        term = item.children('.floatDiv20').text().strip()
                        if term == self.term:
                            title = item.find('div:nth-child(3)').text().strip()
                            cj = item.find('div:nth-child(6)').text().strip()
                            score[title] = cj
                    self.driver.close()

                self.file_changed_signal.emit('{} 存储成功！'.format(self.account))

                thread_save = Thread_save(score=score, file=self.file)
                thread_save.start()

            except Exception:
                self.errorCnt += 1
                self.file_changed_signal.emit('{} 存储<span style="color: red">失败</span>！'.format(self.account))
                self.driver.close()

        self.file_changed_signal.emit('\n\n执行完毕！')
        self.file_changed_signal.emit('错误：<span style="color: red">{}</span> 处'.format(self.errorCnt))
        global time_stop
        time_stop = time.time()
        self.file_changed_signal.emit('stop')


MONGO_URI = 'localhost'
MONGO_DB = 'My_item_1'
MONGO_COLLECTION = 'score_1'


class Thread_save(QThread):
    def __init__(self, score=None, file=None, parent=None):
        super(Thread_save, self).__init__(parent)
        self.score = score
        self.file = file
        self.client = pymongo.MongoClient(MONGO_URI)
        self.db = self.client[MONGO_DB]
        self.collection = MONGO_COLLECTION

    def __del__(self):
        self.wait()

    def run(self):
        # with open(self.file, 'w') as fp:
        #     fp.write(self.score)
        self.db[self.collection].insert(dict(self.score))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    lg = LoginDialog()
    gs = GetScore()
    lg.login_signal.connect(gs.showOrcloseDialog)
    sys.exit(app.exec_())
