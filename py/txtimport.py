from PyQt5.QtWidgets import QDialog, QTableWidgetItem
from PyQt5 import uic
from PyQt5.QtCore import Qt

# own Window for Import of coordinates of control points
class txtImport(QDialog):
    def __init__(self, data):
        super(txtImport, self).__init__()
        uic.loadUi("data/txtimport.ui", self)
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)

        self.data = data
        self.header = []
        self.rowCount = 0

        # Search the best delimiter
        data_string = "\n".join(self.data)
        counts = {
            ",": data_string.count(","),
            ";": data_string.count(";"),
            " ": data_string.count(" "),
            "\t": data_string.count("\t"),
            ":": data_string.count(":")
            }
        
        delimiter = max(counts, key=counts.get)
        
        self.lineE_delimiter.setText(delimiter)
        self.spinB_skiprows.setMinimum(0)
        self.spinB_skiprows.setMaximum(len(self.data) - 1)
        self.spinB_skiprows.setValue(0)

        # Check if file has header or not
        data = []
        data.append(self.data[0].split(delimiter))
        data.append(self.data[1].split(delimiter))
        int_first = 0
        int_second = 0
        for i in range(len(data[0])):
            try:
                int(data[0][i])
                int_first += 1
            except:
                pass
            try:
                int(data[1][i])
                int_second += 1
            except:
                pass
        if int_first < int_second:
            header = True
        else:
            header = False
        self.cB_header.setChecked(header)

        self.fillTable()

        # Connect
        self.pB_import.clicked.connect(self.importdata)
        self.pB_cancel.clicked.connect(self.cancel)
        self.lineE_delimiter.textChanged.connect(self.fillTable)
        self.spinB_skiprows.valueChanged.connect(self.fillTable)
        self.cB_header.clicked.connect(self.fillTable)

    # shows filled table below to see how it gets imported
    def fillTable(self):
        delimiter = self.lineE_delimiter.text()
        skiprows = self.spinB_skiprows.value()
        if not delimiter:
            return
        try:
            skiprows = int(skiprows)
        except:
            return
    
        data = []
        [data.append(i.split(delimiter)) for i in self.data]

        if self.cB_header.checkState():
            self.header = data[0]
            self.rowCount = len(data) - skiprows - 1
            skiprows += 1
        else:
            self.header = [str(i) for i in range(len(data[0]))]
            self.rowCount = len(data) - skiprows
        self.tableWidget.setColumnCount(len(self.header))
        self.tableWidget.setHorizontalHeaderLabels(self.header)
        self.tableWidget.setRowCount(self.rowCount)
        irow = 0
        icol = 0
        
        for row in data[skiprows:]:
            for col in row:
                self.tableWidget.setItem(irow, icol, QTableWidgetItem(col))
                icol += 1
            icol = 0
            irow += 1
        self.tableWidget.resizeColumnsToContents()

    # actual import
    def importdata(self):
        delimiter = self.lineE_delimiter.text()
        data = []
        [data.append(i.split(delimiter)) for i in self.data]
        self.data = data[len(data) - self.rowCount:]
        self.accept()
    
    def cancel(self):
        self.state = False
        self.close()