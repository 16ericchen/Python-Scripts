from PyQt5.QtWidgets import QApplication,QLineEdit,QVBoxLayout,QHBoxLayout,QWidget,QFormLayout,QGridLayout,QTextEdit,QPushButton,QLabel,QTableWidget,QTableWidgetItem,QTabWidget,QErrorMessage
from PyQt5.QtCore import Qt,QEvent

import sys

class lineEditDemo(QWidget):
        def __init__(self,parent=None):
                super().__init__(parent)
                self.error_dialog = QErrorMessage()
                tableWidget = QTableWidget(self)
                tableWidget.setColumnCount(5)
                tableWidget.setHorizontalHeaderLabels(["Project Name","Assembly","Matrix","Note","Serial Number"])
                tableWidget.rowCount()
                self.e8 = QLabel("Output")
                self.e8.setAlignment(Qt.AlignCenter)
                self.projectName = QLineEdit()
                self.projectAssembly = QLineEdit()
                self.projectMatrix = QLineEdit()
                self.projectNote = QLineEdit()
                self.projectSerial = QLineEdit()
                print(self.projectSerial)
                flo = QFormLayout()
                flo.addRow(tableWidget)
                flo.addRow(self.e8)
                flo.addRow("Name",self.projectName)
                flo.addRow("Assembly",self.projectAssembly)
                flo.addRow("Matrix",self.projectMatrix)
                flo.addRow("Note",self.projectNote)
                flo.addRow("Serial Number", self.projectSerial)
                
                self.projectSerial.returnPressed.connect(lambda: self.on_click(tableWidget))
                buttonLayout = QHBoxLayout()
                exportButton = QPushButton("Export",self)
                # exportButton.clicked.connect()
                deleteButton = QPushButton("Delete",self)
                deleteButton.clicked.connect(lambda: self.deleteRow(tableWidget))
                
                buttonLayout.addWidget(exportButton)
                buttonLayout.addWidget(deleteButton)
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

        def on_click(self,tableWidget):
                if self.projectName.text() == "What":

                        row = tableWidget.rowCount()
                        tableWidget.setRowCount(row+1)
                        col = 0
                        items= [self.projectName.text(),self.projectAssembly.text(),self.projectMatrix.text(),self.projectNote.text(),self.projectSerial.text()]
                        for elements in items:
                                cell = QTableWidgetItem(str(elements))
                                tableWidget.setItem(row,col,cell)
                                col+=1
                        self.projectSerial.clear()
                else:
                        self.error_dialog.showMessage("Not What")

        def deleteRow(self,tableWidget):
                row = tableWidget.rowCount()
                if row>0:
                        tableWidget.removeRow(row-1)
if __name__ == "__main__":
        app = QApplication(sys.argv)
        win = lineEditDemo()
        win.show()
        sys.exit(app.exec_())