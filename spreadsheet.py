import os,sys,getpass,datetime,csv,regex as re
from PyQt5.QtWidgets import QTableWidget, QApplication, QMainWindow, QMessageBox
from PyQt5 import (QtWidgets, QtCore)
from PyQt5 import QtGui
from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import (QApplication, QComboBox, QDialog,
		QDialogButtonBox, QFormLayout, QGridLayout, QGroupBox, QHBoxLayout,
		QLabel, QLineEdit, QMenu, QMenuBar, QPushButton, QSpinBox, QTextEdit,
		QVBoxLayout, QFileDialog,QTableWidgetItem,qApp, QAction,QCalendarWidget)



class MyTable(QTableWidget):
	def __init__(self, r, c):
		super().__init__(r, c)
		self.FileOpenProcess = False
		self.ChangeDateProcess = False
		self.initUI()
		self.ProjectName = ""
		self.ProjectManager = ""
		self.Start = ""
		self.cal = QCalendarWidget(self)


	def initUI(self):
		self.cellChanged.connect(self.c_Current)
		self.cellClicked.connect(self.c_Clicked)
		self.show()

	def c_Current(self):
		if self.FileOpenProcess == False:
			row = self.currentRow()
			col = self.currentColumn()
			self.FileOpenProcess = True
			try:
				if int(col) == 6 :
					print ("in calculation")
					endDate = self.calculateEndDate(row)
					print (endDate)
					self.setItem(row, 7, QTableWidgetItem(str(endDate)))
				elif int(col) in (2,4):
					print("in change dependency" )
					self.recalcDates(row)
			except:
				if (self.item(row,4) is None ):
					self.setItem(row,4,QTableWidgetItem("FS"))
					pass
				elif (self.item(row,2) is None ):
					self.setItem(row,2, QTableWidgetItem("0"))
					pass
				elif (self.item(row, 5) is None):
					self.setItem(row, 5, QTableWidgetItem(str(datetime.datetime.today())))
					if (self.item(row, 6) is None):
						self.setItem(row, 6, QTableWidgetItem("1d"))

				elif (self.item(row, 6) is None):
					self.setItem(row, 6, QTableWidgetItem("1d"))
					self.recalcDates(row)

				else:
					warn = QMessageBox(QMessageBox.Warning, "Error!",  "Could Not Calculate")
					warn.setStandardButtons(QMessageBox.Ok);
					warn.exec_()
					self.setItem(row, col, QTableWidgetItem(str("")))

		self.FileOpenProcess = False

	def c_Clicked(self):
		if self.FileOpenProcess == False:
			row = self.currentRow()
			col = self.currentColumn()
			self.FileOpenProcess == True
			if int(col) in(5,7):
				self.myCal()

	def myCal(self):
		self.cal.setGridVisible(True)
		self.cal.move(QtGui.QCursor.pos().x(), QtGui.QCursor.pos().y()-60)
		print (self.mapToGlobal(QtGui.QCursor.pos()))
		self.ChangeDateProcess = True
		self.cal.clicked.connect(self.getDate)
		self.cal.setGridVisible(True)	
		self.cal.show()
		
	def getDate(self):
		row = self.currentRow()
		col = self.currentColumn()
		if self.ChangeDateProcess:
			print (self.cal.selectedDate().toString())
			date = str(datetime.datetime.strptime( self.cal.selectedDate().toString(), '%a %b %d %Y') + datetime.timedelta(hours=8))
			self.setItem(row,col,QTableWidgetItem(date))
			print (date)
			self.ChangeDateProcess = False
			self.cal.hide()
		else:
			pass

	def recalcDates(self, row):
		type = self.item(int(row),4).text()
		print (type)
		if type.lower() == "fs":
			startDate = self.item(int(self.item(int(row-1), 2).text()), 7).text()
		elif type.lower() == 'ss':
			startDate = self.item(int(self.item(int(row-1), 2).text()), 5).text()
		self.setItem(int(row),5, QTableWidgetItem(str(startDate)))
		self.setItem(int(row),7, QTableWidgetItem(str(self.calculateEndDate (row))))


	def calculateEndDate(self, row):
		r = re.compile("([0-9]+)([a-zA-Z]+)")
		duration = str(self.item(row,6).text())
		print(duration)
		m = r.match(duration)
		duration = int(m.group(1))
		period = str(m.group(2))
		##2018-05-12 08:00:00
		endDate =  self.item(row,6).text()
		date = datetime.datetime.strptime( self.item(row,5).text() , '%Y-%m-%d %H:%M:%S')
		if period.lower() == 'h':
			enddate = date + datetime.timedelta(hours=int(duration))
			print  (endDate)
		elif period.lower() == 'm':
			enddate = date + datetime.timedelta(minutes=int(duration))
		elif period.lower() == 'd':
			enddate = date + datetime.timedelta(days=int(duration))
		elif period.lower() == 'w':
			enddate = date + datetime.timedelta(weeks=int(duration))
		else:
			QMessageBox.warning(self.view, 'Wrong Format', msg)

		return(enddate)




	def openProject(self):
		self.FileOpenProcess = True
		file = QFileDialog.getOpenFileName(self, 'Open NPF', os.getenv('HOME'), 'NPF(*.npf)')
		if file[0] != '':
			print (file[0])
			with open(file[0], newline='') as csv_file:
				self.setRowCount(0)
				self.setColumnCount(9)
				my_file = csv.reader(csv_file, dialect='excel')
				for row_data in my_file:
					row = self.rowCount()
					self.insertRow(row)
					if len(row_data) > 8:
						self.setColumnCount(len(row_data))
					for column, stuff in enumerate(row_data):
						item = QTableWidgetItem(stuff)
						self.setItem(row, column, item)
		self.FileOpenProcess = False

	def saveProject(self):
		path = QFileDialog.getSaveFileName(self, 'Save Project', os.getenv('HOME'), 'NPF(*.npf)')
		if path[0] != '':
			with open(path[0], 'w', newline = '') as csv_file:
				writer = csv.writer(csv_file, dialect='excel')
				for row in range(self.rowCount()):
					row_data = []
					for column in range(self.columnCount()):
						item = self.item(row, column)
						if item is not None:
							row_data.append(item.text())
						else:
							row_data.append('')
					writer.writerow(row_data)
					
	def add_row(self):
		rows = self.rowCount()
		self.setRowCount(self.rowCount() + 1)

	def remove_row(self):
		row = self.currentRow()
		cancel = False
		if row is not None:
			for i in range (0,8):
				data = self.item(row,i)
				if data is not None:
					reply = QMessageBox.question(
							self,
							'Confirm Delete, There is data in this row',
							'Are you sure you want to delete this row?',
							QMessageBox.Yes | QMessageBox.No,
							QMessageBox.No)
					if reply == QMessageBox.Yes:
						i = 9
						cancel = False
					else:
						cancel = True
	
			if cancel == False:
				self.removeRow(row)

	def indent_task(self):
		self.FileOpenProcess = True
		row = self.currentRow()
		print ("Current Row == " + str(row))
		if row != 0:
			currentItemIndent = len(str(self.item(row,2).text())) - len(str(self.item(row,2).text()).lstrip())
			parentrow = self.findParent(row,currentItemIndent)
			parent = str(self.item(parentrow,3).text())
			parentIndent = len(parent) - len(parent.lstrip())
			currentItem = str(self.item(row,3).text()).lstrip()
			print ("Indent Level == " + str(parentIndent+4))
			if currentItemIndent < parentIndent:
				futureIndent = parentIndent +4
			elif currentItemIndent == 0:
				futureIndent = 4
			elif currentItemIndent > parentIndent:
				futureIndent = parentIndent + 4

			for i in range (1,futureIndent):
				currentItem = " " + currentItem
			self.setItem(row, 3, QTableWidgetItem(currentItem))
			self.renumber()
		self.FileOpenProcess = False

	def findParent(self,row,currentItemIndent):
		parentrow = row
		for i in range (row,0, -1):
			parent_indent = len(str(self.item(i,3).text())) - len(str(self.item(i,3).text()).lstrip())
			if parent_indent > currentItemIndent:
				parentrow = i
				print ("Parent Indent = " + str (parent_indent) + "Current Indent ==" + str(currentItemIndent))
				break
		print (parentrow)
		return (parentrow -1)

	def renumber(self):
		rowsNo = int(self.rowCount())
		for i in range (0,rowsNo):
			if i == 0:
				self.setItem(i, 1, QTableWidgetItem("1.0"))
			else:
				#print ("i ====" + str(i))
				#print ("Previous Item Value ====" + self.item(i-1,3).text())
				#print("Previous WBS ID ====" + self.item(i-1,1).text())
				currentText = str(self.item(i,3).text())
				PrevWbsID = float(str(self.item(i-1,1).text()))
				prevIndent =  len(currentText) - len(currentText.lstrip())
				wbsDeapth = int(prevIndent/4)
				#print (wbsDeapth)
				self.setItem(i,1,QTableWidgetItem(str(PrevWbsID +1)))

	def outdent_task(self):
		row = self.currentRow()
		itemName = str(self.item(row,3).text())
		checkforSpaces = itemName[:4]
		print ('Left' + itemName[:4])
		if checkforSpaces.strip() == '':
			newText =itemName[4:]
			self.setItem(row,3,QTableWidgetItem(newText))  

class Sheet(QMainWindow):
	def __init__(self):
		super().__init__()
		self.form_widget = MyTable(1, 9)
		self.setCentralWidget(self.form_widget)
		col_headers = [' % Complete','WBS Id','Predicisors','TaskName','Task Mode','Start Date','Duration','End Date','Resources']
		self.form_widget.setHorizontalHeaderLabels(col_headers)
		##insertFromData(self)
		bar = self.menuBar()

		file = bar.addMenu('File')
		save_action = QAction('&Save', self)
		save_action.setShortcut('Ctrl+S')
		open_action = QAction('&Open', self)
		open_action.setShortcut('Ctrl+O')
		quit_action = QAction('&Quit', self)
		quit_action.setShortcut('Ctrl+q')
		file.addAction(save_action)
		file.addAction(open_action)
		file.addAction(quit_action)
		quit_action.triggered.connect(self.quit_app)
		save_action.triggered.connect(self.form_widget.saveProject)
		open_action.triggered.connect(self.form_widget.openProject)

		edit = bar.addMenu('Edit')
		add_new_row = QAction("&Add Row", self)
		add_new_row.setShortcut("Ins")
		remove_selected_row = QAction("&Remove Row", self)
		remove_selected_row.setShortcut("Del")
		indent_task = QAction("&Indent Task", self)
		indent_task.setShortcut('Ctrl+Right')
		outdent_task = QAction("&Outdent Task", self)
		outdent_task.setShortcut('Ctrl+Left')
		edit.addAction(add_new_row)
		edit.addAction(remove_selected_row)
		edit.addAction(indent_task)
		edit.addAction(outdent_task)
		add_new_row.triggered.connect(self.form_widget.add_row)
		remove_selected_row.triggered.connect(self.form_widget.remove_row)
		indent_task.triggered.connect(self.form_widget.indent_task)
		outdent_task.triggered.connect(self.form_widget.outdent_task)
		self.showMaximized()

	def quit_app(self):
		sys.exit(app.exec_())


app = QtWidgets.QApplication(sys.argv)
sheet = Sheet()


sys.exit(app.exec_())