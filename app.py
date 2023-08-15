from PyQt5.uic import loadUi
from PyQt5.QtWidgets import (
    QMainWindow,
    QApplication,
    QFileDialog,
    QMessageBox,
    QDialog,
    QProgressBar,
    QVBoxLayout,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QUrl
from py.txtimport import txtImport
from py.about import AboutWindow
from PyQt5.QtWebEngineWidgets import QWebEngineView
import ssl
import smtplib
import json
from email.message import EmailMessage


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        loadUi("data/window.ui", self)

        self.textEditContent = None
        self.verticalLayout_2.setContentsMargins(9, 9, 9, 9)

        self.pB_loaddata.clicked.connect(self.loaddata)
        self.pB_addVar.clicked.connect(self.addVar)
        self.pB_send.clicked.connect(self.sendmails)
        self.pB_testconn.clicked.connect(self.testconn)
        self.pB_preview.clicked.connect(self.preview)
        self.pB_prevpreview.clicked.connect(self.prevpreview)
        self.pB_nextpreview.clicked.connect(self.nextpreview)
        self.sB_preview.valueChanged.connect(self.setpreview)
        self.actionAbout.triggered.connect(self.about)
        self.actionDocs.triggered.connect(self.docs)
        self.loadinputs()

    def loadinputs(self):
        with open("inputs.json", "r") as file:
            inputs = json.load(file)

        self.lE_host.setText(inputs["host"])
        self.lE_port.setText(inputs["port"])
        self.lE_user.setText(inputs["user"])
        self.lE_sender.setText(inputs["sendermail"])

    def loaddata(self):
        file_path = QFileDialog.getOpenFileName(
            self, "Load Data...", filter="Text files (*.txt *.csv);;All Files (*.*)"
        )[0]
        self.path_coords = file_path.split("/")[-1]
        if file_path == "":
            return

        data = []
        with open(file_path, encoding="utf-8") as file:
            [data.append(row.splitlines()[0]) for row in file]

        self.txtImportWindow = txtImport(data)
        self.txtImportWindow.exec_()
        self.header = self.txtImportWindow.header
        self.data = self.txtImportWindow.data
        self.pB_addVar.setEnabled(True)
        self.cB_vars.addItems(self.header)
        self.cB_vars.setEnabled(True)
        self.lbl_receiver.setEnabled(True)
        header = self.header.copy()
        header.insert(0, "")
        self.cB_receiver.addItems(header)
        self.cB_receiver.setEnabled(True)
        self.lbl_loadeddata.setText(f"{len(self.data)} rows were loaded.")
        self.sB_preview.setMaximum(len(self.data))
        self.pB_preview.setEnabled(True)

    def addVar(self):
        var = "$?" + self.cB_vars.currentText() + "?$"
        self.textEdit.insertPlainText(var)
        self.textEdit.setFocus()

    def preview(self):
        if not self.sB_preview.isEnabled():
            self.pB_preview.setText("Hide Preview")
            self.textEdit.setDisabled(True)
            if self.sB_preview.value() != self.sB_preview.minimum():
                self.pB_prevpreview.setEnabled(True)
            if self.sB_preview.value() != self.sB_preview.maximum():
                self.pB_nextpreview.setEnabled(True)
            self.sB_preview.setEnabled(True)
            self.textEditContent = self.textEdit.toPlainText()
            self.setpreview(self.textEditContent)
        else:
            self.pB_preview.setText("Show Preview")
            self.textEdit.setEnabled(True)
            self.pB_prevpreview.setDisabled(True)
            self.pB_nextpreview.setDisabled(True)
            self.sB_preview.setDisabled(True)
            self.textEdit.setPlainText(self.textEditContent)
            self.textEditContent = None

    def setpreview(self, string=None):
        if type(string) != int:
            msg = self.insertvars(self.sB_preview.value() - 1, string)
        elif self.textEditContent:
            msg = self.insertvars(self.sB_preview.value() - 1, self.textEditContent)
        else:
            return
        self.textEdit.setPlainText(msg)

    def prevpreview(self):
        self.sB_preview.stepBy(-1)
        if self.sB_preview.value() == self.sB_preview.minimum():
            self.pB_prevpreview.setDisabled(True)
        else:
            self.pB_nextpreview.setEnabled(True)

    def nextpreview(self):
        self.sB_preview.stepBy(1)
        if self.sB_preview.value() == self.sB_preview.maximum():
            self.pB_nextpreview.setDisabled(True)
        else:
            self.pB_prevpreview.setEnabled(True)

    def sendmails(self):
        host = self.lE_host.text()
        port = self.lE_port.text()
        user = self.lE_user.text()
        pw = self.lE_pw.text()
        sender_email = self.lE_sender.text()
        subject = self.lE_subject.text()
        receiver = self.cB_receiver.currentText()
        if self.textEditContent:
            msg_string = self.textEditContent
        else:
            msg_string = self.textEdit.toPlainText()
        if not subject:
            ret = QMessageBox.information(
                self,
                "No subject",
                "You haven't filled a subject. Would you like to send the E-Mail without subject?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if ret == QMessageBox.No:
                return
        if not (host and port and user and pw and sender_email):
            QMessageBox.warning(self, "Not filled", "Please fill all fields and retry")
            return
        if not receiver:
            QMessageBox.warning(
                self, "No Receiver", "Please select a Receiver from the dropdown!"
            )
            return
        receiver_index = self.header.index(receiver)
        thread = sending(
            host,
            port,
            user,
            pw,
            msg_string,
            subject,
            sender_email,
            receiver_index,
            self.data,
            self.header,
        )
        thread.sended.connect(self.sended)
        thread.success.connect(self.success)
        thread.error.connect(self.error)
        thread.start()
        self.progressdialog = QDialog()
        self.progressdialog.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        self.progbar = QProgressBar(self.progressdialog)
        layout = QVBoxLayout()
        layout.addWidget(self.progbar)
        self.progressdialog.setLayout(layout)
        self.progbar.setMaximum(len(self.data))
        self.progressdialog.setWindowTitle("Sending Progress...")
        self.progressdialog.setWindowModality(Qt.ApplicationModal)
        self.progressdialog.exec_()
        self.saveInputs()

    def sended(self, num):
        self.progbar.setValue(num)

    def success(self):
        QMessageBox.information(
            self, "Sending successfull", "All E-Mails were sent successfully."
        )
        self.progressdialog.close()

    def error(self, e, i, addr):
        QMessageBox.warning(
            self,
            "Could not connect",
            f"The connection failed. Please check the input and retry. Error occured at element {i} ({addr}). Error: {e}",
        )
        self.progressdialog.close()

    def insertvars(self, dataindex, msg_string):
        msg_split = msg_string.split("$?")
        for i in range(1, len(msg_split)):
            msg_split[i] = msg_split[i].split("?$")
            msg_split[i][0] = self.data[dataindex][self.header.index(msg_split[i][0])]
            msg_split[i] = "".join(msg_split[i])
        msg_split = "".join(msg_split)
        return msg_split

    def testconn(self):
        host = self.lE_host.text()
        port = self.lE_port.text()
        user = self.lE_user.text()
        pw = self.lE_pw.text()
        if not (host and port and user and pw):
            QMessageBox.warning(self, "Not filled", "Please fill all fields and retry")
        context = ssl.create_default_context()
        try:
            with smtplib.SMTP(host, port) as server:
                server.starttls(context=context)
                server.login(user, pw)
            QMessageBox.information(
                self, "Connection successful", "The connection was successfull."
            )
        except Exception as e:
            QMessageBox.warning(
                self,
                "Could not connect",
                f"The connection failed. Please check the input and retry. Error: {e}",
            )
        self.saveInputs()

    def saveInputs(self):
        host = self.lE_host.text()
        port = self.lE_port.text()
        user = self.lE_user.text()
        sendermail = self.lE_sender.text()

        inputs = {"host": host, "port": port, "user": user, "sendermail": sendermail}

        with open("inputs.json", "w") as file:
            json.dump(inputs, file)

    def about(self):
        about = AboutWindow()
        about.exec_()

    def docs(self):
        docs = QDialog()
        docs.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        layout = QVBoxLayout()
        webengine = QWebEngineView()
        webengine.load(QUrl().fromLocalFile("/docs/index.html"))
        layout.addWidget(webengine)
        docs.setLayout(layout)
        docs.setWindowTitle("Documentation")
        docs.setWindowModality(Qt.ApplicationModal)
        docs.exec_()


class sending(QThread):
    sended = pyqtSignal(int)
    success = pyqtSignal()
    error = pyqtSignal(Exception, int, str)

    def __init__(
        self,
        host,
        port,
        user,
        pw,
        msg_string,
        subject,
        sender_email,
        receiver_index,
        data,
        header,
    ):
        super(sending, self).__init__()
        self.roundnum = 1
        self.host = host
        self.port = port
        self.user = user
        self.pw = pw
        self.msg_string = msg_string
        self.subject = subject
        self.sender_email = sender_email
        self.receiver_index = receiver_index
        self.data = data
        self.header = header

    def run(self):
        context = ssl.create_default_context()
        with smtplib.SMTP(self.host, self.port) as server:
            try:
                server.starttls(context=context)
                server.login(self.user, self.pw)
            except Exception as e:
                self.error.emit(e, 0, "Connection error")
                self.exit()
                return
            for i in range(len(self.data)):
                try:
                    msg = EmailMessage()
                    msg.set_content(self.insertvars(i, self.msg_string))
                    msg["Subject"] = self.insertvars(i, self.subject)
                    msg["From"] = self.sender_email
                    msg["To"] = self.data[i][self.receiver_index]
                    server.send_message(msg)
                    self.sended.emit(self.roundnum)
                    self.roundnum += 1
                except Exception as e:
                    self.error.emit(
                        e, self.roundnum - 1, self.data[i][self.receiver_index]
                    )
                    self.exit()
                    return
        self.success.emit()
        self.exit()

    def insertvars(self, dataindex, msg_string):
        msg_split = msg_string.split("$?")
        for i in range(1, len(msg_split)):
            msg_split[i] = msg_split[i].split("?$")
            msg_split[i][0] = self.data[dataindex][self.header.index(msg_split[i][0])]
            msg_split[i] = "".join(msg_split[i])
        msg_split = "".join(msg_split)
        return msg_split


if __name__ == "__main__":
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication([])

    window = MainWindow()
    window.show()
    app.exec()
