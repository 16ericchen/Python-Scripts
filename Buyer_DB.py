from PyQt5.QtWidgets import QApplication,QLineEdit,QVBoxLayout,QHBoxLayout,QWidget,QFormLayout,QTableWidget,QTableWidgetItem,QErrorMessage
from GUI_Functions import enterData
import sys

class lineEditDemo(QWidget):
        def __init__(self,parent=None):
                super().__init__(parent)
                self.storeData =''
                self.error_dialog = QErrorMessage()
                self.error_dialog.setMinimumSize(800,200)
                tableWidget = QTableWidget(self)
                tableWidget.setColumnCount(2)
                tableWidget.setHorizontalHeaderLabels(["Project Name","Quantity"])
                tableWidget.setColumnWidth(0,250)
                tableWidget.setColumnWidth(1,250)
                self.APN = QLineEdit()
                self.APN.setClearButtonEnabled(True)
                self.APN.returnPressed.connect(lambda: self.on_click(tableWidget))
                flo = QFormLayout()
                flo.addRow(tableWidget)
                flo.addRow("APN:",self.APN)
                buttonLayout = QHBoxLayout()
                mainLayout = QVBoxLayout()
                mainLayout.addLayout(flo)
                mainLayout.addLayout(buttonLayout)
                self.setLayout(mainLayout)
                self.setWindowTitle("Project Inventory Database")
                self.setMinimumSize(500,600)

                      
        def on_click(self,tableWidget):
                if self.APN.text():   
                        tableWidget.setRowCount(0)        
                        row = tableWidget.rowCount()
                        tableWidget.setRowCount(row+1)
                        items = enterData(self.APN.text())
                        data = items
                        qty = ['None']
                        if items[0] != "No Projects Found":
                                data = items[0][1].split(',')
                                qty = items[0][2].split(',')
                        for i in range(len(data)):
                                cell = QTableWidgetItem(str(data[i]))
                                cell2 = QTableWidgetItem(str(qty[i]))
                                tableWidget.setItem(row,0,cell)
                                tableWidget.setItem(row,1,cell2)
                                row+=1
                                tableWidget.setRowCount(row+1)

if __name__ == "__main__":
        app = QApplication(sys.argv)
        win = lineEditDemo()
        win.show()
        sys.exit(app.exec_())