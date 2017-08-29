import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QCoreApplication
from PyQt5 import QtCore
from PyQt5.QtGui import QIcon
import os
from PyQt5 import QtWidgets

class Example(QWidget):

	def __init__(self):
		super().__init__()
		self.setGeometry(50, 50, 400, 300)
		self.setWindowTitle("Log In to your facebook account")
		self.setStyleSheet("background-color: #3b5998;")
		self.show()

		self.initUI()

	def initUI(self):
		LineEditStyle = "background-color: white;padding: 10px 10px 10px 20px; font-size: 14px; font-family: consolas;" \
                        "border: 2px solid #3BBCE3; border-radius: 4px; width:100px;"
		logo = QLabel("Zillow Scraper")
		logo.setStyleSheet("font-size: 50px; font-weight:bold;color:white;")
		logo.setAlignment(QtCore.Qt.AlignCenter | QtCore.Qt.AlignVCenter)

		email = QLineEdit()
		email.setPlaceholderText("Enter Your Search Param")
		email.setStyleSheet(LineEditStyle)

		btn = QtWidgets.QPushButton('Select a CSV', self)
		btn.setStyleSheet(LineEditStyle)
		btn.resize(btn.sizeHint())
		btn.clicked.connect(self.browse_file)

		s_btn = QtWidgets.QPushButton('Start Crawler', self)
		s_btn.setStyleSheet(LineEditStyle)
		s_btn.resize(btn.sizeHint())
		# s_btn.clicked.connect(self.browse_file)
		self.textEdit = QLineEdit()
		
		layout = QVBoxLayout()

		self.setLayout(layout)
		layout.addWidget(logo)
		layout.addStretch()
		layout.addWidget(email)
		layout.addWidget(btn)
		layout.addWidget(self.textEdit)
		layout.addStretch(1)
		layout.addWidget(s_btn)
		layout.addStretch(1)

	def browse_file(self):
		fname = QFileDialog.getOpenFileName(self, 
		                  "Open file", os.path.expanduser('~'))
		if fname[0] and "csv" in fname[0]:
			f = open(fname[0], 'r')
			with f:
				self.textEdit.setText(fname[0])
				data = f.read()
				print(data)


if __name__ == '__main__':

	app = QApplication(sys.argv)
	ex = Example()
	sys.exit(app.exec_())
