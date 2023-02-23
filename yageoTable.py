from searchDB_WS import enterData
import sys,os
from time import sleep
from PyQt5.QtWidgets import QApplication,QVBoxLayout,QHBoxLayout,QWidget,QFormLayout,QTableWidget,QTableWidgetItem,QErrorMessage,QPushButton
# os.chdir(os.path.dirname(__file__))

class lineEditDemo(QWidget):
        def __init__(self,apnDict,parent=None):
                super().__init__(parent)
                self.storeData =''
                self.error_dialog = QErrorMessage()
                self.error_dialog.setMinimumSize(800,200)
                self.tableWidget = QTableWidget()
                self.tableWidget.setColumnCount(2)
                self.tableWidget.setHorizontalHeaderLabels(["Original APN","Modified APN"])
                self.tableWidget.setRowCount(len(apnDict))
                for i,k in enumerate(apnDict):
                        self.tableWidget.setItem(i,0,QTableWidgetItem(k))
                self.tableWidget.setColumnWidth(0,250)
                self.tableWidget.setColumnWidth(1,250)
                self.button = QPushButton(self)
                self.button.setShortcut('Ctrl+H')
                self.button.setText("Update Database")
                self.button.clicked.connect(lambda: self.test(self.tableWidget))
                flo = QFormLayout()
                flo.addRow(self.tableWidget)
                buttonLayout = QHBoxLayout()
                buttonLayout.addWidget(self.button)
                mainLayout = QVBoxLayout()
                mainLayout.addLayout(flo)
                mainLayout.addLayout(buttonLayout)
                self.setLayout(mainLayout)
                self.setWindowTitle("Project Inventory Database")
                self.setMinimumSize(546,600)

        def test(self,tableWidget):
            for x in range(self.tableWidget.rowCount()):
                if tableWidget.item(x, 1):
                    enterData(tableWidget.item(x, 0).text(),tableWidget.item(x, 1).text())
            print('Database Updated')
            sleep(3)
            self.close()

            
def table(apnDict):
    app = QApplication(sys.argv)
    win = lineEditDemo(apnDict)
    win.show()

    app.exec_()
    return


