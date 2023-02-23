from PyQt5.QtWidgets import QApplication,QLineEdit,QVBoxLayout,QHBoxLayout,QWidget,QFormLayout,QGridLayout,QTextEdit,QPushButton,QLabel,QTableWidget,QTableWidgetItem,QTabWidget,QErrorMessage
from PyQt5.QtCore import Qt,QEvent
from GUI_Functions import exportExcel,enterData,deleteEntry
import sys

class lineEditDemo(QWidget):
        def __init__(self,parent=None):
                super().__init__(parent)
                self.storeData =''
                self.error_dialog = QErrorMessage()
                self.error_dialog.setMinimumSize(800,200)
                tableWidget = QTableWidget(self)
                tableWidget.setColumnCount(5)
                tableWidget.setHorizontalHeaderLabels(["Project Name","Assembly","Matrix","Note","Serial Number"])
                tableWidget.rowCount()
                self.projectName = QLineEdit()
                self.projectName.setClearButtonEnabled(True)
                self.projectAssembly = QLineEdit()
                self.projectAssembly.setClearButtonEnabled(True)
                self.projectMatrix = QLineEdit()
                self.projectMatrix.setClearButtonEnabled(True)
                self.projectNote = QLineEdit()
                self.projectNote.setClearButtonEnabled(True)
                self.projectSerial = QLineEdit()
                flo = QFormLayout()
                flo.addRow(tableWidget)
                flo.addRow("Name:",self.projectName)
                flo.addRow("Assembly:",self.projectAssembly)
                flo.addRow("Matrix:",self.projectMatrix)
                flo.addRow("Note:",self.projectNote)
                flo.addRow("Serial Number:", self.projectSerial)
                
                self.projectSerial.returnPressed.connect(lambda: self.on_click(tableWidget))
                buttonLayout = QHBoxLayout()
                exportButton = QPushButton("Export",self)
                exportButton.clicked.connect(self.uploadMessage)
                deleteButton = QPushButton("Delete",self)
                deleteButton.clicked.connect(lambda: self.deleteRow(tableWidget))
                clearButton = QPushButton("Clear")
                clearButton.clicked.connect(self.clearTextLine)
                buttonLayout.addWidget(exportButton)
                buttonLayout.addWidget(deleteButton)
                buttonLayout.addWidget(clearButton)
                mainLayout = QVBoxLayout()
                mainLayout.addLayout(flo)
                mainLayout.addLayout(buttonLayout)
                self.setLayout(mainLayout)
                self.setWindowTitle("Shipping Database")
                self.setMinimumSize(525,600)

        def event(self, event):
                if event.type() == QEvent.KeyPress and event.key() in (
                        Qt.Key_Enter,
                        Qt.Key_Return,
                ):
                        widget = QApplication.focusWidget()
                        if widget != self.projectSerial:
                                self.focusNextPrevChild(True)
                        widget = QApplication.focusWidget()
                return super().event(event)

        def clearTextLine(self):
                items= [self.projectName,self.projectAssembly,self.projectMatrix,self.projectNote,self.projectSerial]
                for x in items:
                        x.clear()
                        
        def on_click(self,tableWidget):

                items= [self.projectName.text(),self.projectAssembly.text(),self.projectMatrix.text(),self.projectNote.text(),self.projectSerial.text()]
                x = enterData(items)
                if x:           
                        self.error_dialog.showMessage("Serial Number Already Exists: "+"Project Date: "+x[-1]+', Project Name: '+x[0]+", Project Assembly: "+x[1]+", Serial Number: "+x[2]+", Project Matrix: "+x[3]+", Project Note: "+x[4])
                else: 
                        row = tableWidget.rowCount()
                        tableWidget.setRowCount(row+1)
                        col = 0
                        for elements in items:
                                cell = QTableWidgetItem(str(elements))
                                tableWidget.setItem(row,col,cell)
                                col+=1
                        self.storeData = self.projectSerial.text()
                        self.projectSerial.clear()

        def deleteRow(self,tableWidget):
                row = tableWidget.rowCount()
                items= [self.projectName.text(),self.projectAssembly.text(),self.projectMatrix.text(),self.projectNote.text(),self.storeData]
                if row>0:
                        deleteEntry(items)
                        tableWidget.removeRow(row-1)
                        if row-2>0:
                                current = tableWidget.item(row-2,4)
                                self.storeData = current.text()
        
        def uploadMessage(self):
                exportExcel()
                if exportExcel:
                        self.error_dialog.showMessage("Export Done")
if __name__ == "__main__":
        app = QApplication(sys.argv)
        win = lineEditDemo()
        win.show()
        sys.exit(app.exec_())