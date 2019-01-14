import sys
from PyQt5.QtWidgets import QMainWindow, QApplication
from PyQt5 import uic, QtCore, QtGui
import pandas as pd
import datetime as dt
import os
from bs4 import BeautifulSoup
import re
import xlrd
import requests
from docx import *
import time
from difflib import SequenceMatcher as SM
import glob
import win32com.client as win32

Ui_MainWindow, QtBaseClass = uic.loadUiType('dialog.ui')

class PandasModel(QtCore.QAbstractTableModel): 
	def __init__(self, df = pd.DataFrame(), parent=None): 
		QtCore.QAbstractTableModel.__init__(self, parent=parent)
		self._df = df

	def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
		if role != QtCore.Qt.DisplayRole:
			return QtCore.QVariant()

		if orientation == QtCore.Qt.Horizontal:
			try:
				return self._df.columns.tolist()[section]
			except (IndexError, ):
				return QtCore.QVariant()
		elif orientation == QtCore.Qt.Vertical:
			try:
				# return self.df.index.tolist()
				return self._df.index.tolist()[section]
			except (IndexError, ):
				return QtCore.QVariant()

	def data(self, index, role=QtCore.Qt.DisplayRole):
		if role != QtCore.Qt.DisplayRole:
			return QtCore.QVariant()

		if not index.isValid():
			return QtCore.QVariant()

		return QtCore.QVariant(str(self._df.ix[index.row(), index.column()]))

	def setData(self, index, value, role):
		row = self._df.index[index.row()]
		col = self._df.columns[index.column()]
		if hasattr(value, 'toPyObject'):
			# PyQt4 gets a QVariant
			value = value.toPyObject()
		else:
			# PySide gets an unicode
			dtype = self._df[col].dtype
			if dtype != object:
				value = None if value == '' else dtype.type(value)
		self._df.set_value(row, col, value)
		return True

	def rowCount(self, parent=QtCore.QModelIndex()): 
		return len(self._df.index)

	def columnCount(self, parent=QtCore.QModelIndex()):
		return(len(self._df.columns))

class MyApp(QMainWindow):
	path_pattern='path/to/contact_list.docx'
	def __init__(self):
		super(MyApp, self).__init__()
		self.ui = Ui_MainWindow()
		self.setFixedSize(800,600)
		self.ui.setupUi(self)
		self.ui.cl1_vs_cl2.clicked.connect(self.main_func1)
		self.ui.cl2_vs_cl1.clicked.connect(self.main_func2)

	def main_func1(self):
		self.username = str(self.ui.username_input.text())
		self.password = str(self.ui.password_input.text())
		self.protocol = str(self.ui.protocol_input.text())
		self.get_harddrive_cl_text()
		self.get_online_contacts()
		self.compare_cl1vcl2()
		model = PandasModel(self.df_cl1vcl2_comparison)
		self.ui.comparison_table.setModel(model)

	def main_func2(self):
		self.username = str(self.ui.username_input.text())
		self.password = str(self.ui.password_input.text())
		self.protocol = str(self.ui.protocol_input.text())
		self.get_firstcl_text()
		self.get_secondcl_text()
		self.compare_cls()
		model = PandasModel(self.df_cl2vcl1_comparison)
		self.ui.comparison_table.setModel(model)

	def get_harddrive_cl_text(self):
		yr_part=re.search("(\d{2})\-\d{3,4}",self.protocol)[1]
		self.mpath = MyApp.path_pattern.format(y=yr_part,p=self.protocol)
		exclude = ['email','e-mail']
		list_of_cls = []
		for file in glob.glob(self.mpath+self.protocol+"*List*"):
			if not any(i in file.lower() for i in exclude):
				list_of_cls.append(file)
		latest_cl = max(list_of_cls, key=os.path.getctime)
		document = Document(latest_cl)
		names = []
		for para in document.paragraphs:
			if para.alignment != 1:
				if all(run.bold and not run.underline for run in para.runs):
					names.append(para.text)
		names = [i for i in names if len(i) > 2 and ":" not in i]
		self.cl_text = list(set(names))

	def get_online_contacts(self):
		cms = "www.target-page.com/login.aspx"
		ua = {'User-Agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.110 Safari/537.36',
		'X-Requested-With': 'XMLHttpRequest',
		'Accept': 'application/json, text/javascript, */*; q=0.01',
		'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
		'X-Requested-With': 'XMLHttpRequest'}
		data = {'txtUserName':self.username,'txtPassword':self.password}
		# target urls
		url1 = "www.target-page.com/page_1"
		url2 = "www.target-page.com/page_2"
		search_url = "www.target-page.com/page_3"
		cl_url = "www.target-page.com/page_3/_DownloadContacts?Id={p}&name={f}"
		# sign into cms
		s = requests.Session()
		s.headers['User-Agent'] = ua['User-Agent']
		r1 = s.get(cms)
		soup = BeautifulSoup(r1.text,'html.parser')
		for elem in soup.form.findAll('input'):
			try:
				if 'value' in elem.attrs.keys():
					data[elem['name']] = elem['value']
				else:
					data[elem['name']] = data[elem['name']]
			except:
				continue
		data["btnSignIn.x"] = 52
		data["btnSignIn.y"] = 13
		data['RememberMe'] = 'on'
		r2 = s.post(cms,data=data)
		r2 = s.get(url1)
		prot_soup = BeautifulSoup(r2.text,'html.parser')
		for option in prot_soup.find_all('option'):
			if option.text == self.protocol:
				self.protid = option['value']
		r3 = s.get(url2)
		data = {'protocol':self.protid}
		r3 = s.post(search_url,data=data,headers=ua)
		response_soup = BeautifulSoup(r3.text, 'html.parser')
		filename = response_soup.find(id="CommitteeName")['value']
		r4 = s.get(cl_url.format(p=self.protid,f=filename)) 
		if r4.status_code == 200:
			df_cms = pd.read_excel(xlrd.open_workbook(file_contents=r4.content), engine='xlrd')
			name_columns = df_cms['Name'].tolist()
			name_list = []
			for i in name_columns: # changes last name, first name to first name last name
				name_parts = i.split(", ",1)
				full_name = name_parts[1] + " " + name_parts[0]
				name_list.append(full_name)
		self.name_list = name_list

	def compare_cl1vcl2(self):
		exclude = ['names','to','exclude']
		all_scores = []
		for name in self.name_list:
			d = {}
			if not any(i in name.lower() for i in exclude):
				d['CL1 Name'] = name
				score = [] # similarity score of cms name to each line in text file
				for elem in self.cl_text:
					if ", " in elem:
						elem = elem.split(", ",1)[0]
					if " (" in elem:
						elem = elem.split(" (",1)[0]
					s = SM(None, name, elem).ratio()
					score.append(s*100)
				m = max(score) # get text from mdrive cl thats most similar to cms cl
				d['Nearest CL2 Name'] = self.cl_text[score.index(m)]
				d['score'] = m
				all_scores.append(d)
		self.comparison = all_scores
		self.df_cms2mdrive_comparison = pd.DataFrame(all_scores)

	def compare_cl2vcl1(self): # compares m_cl to cms_cl
		exclude = ['names','to','exclude']
		all_scores = []
		for elem in self.cl_text:
			if ", " in elem:
				elem = elem.split(", ",1)[0]
			if " (" in elem:
				elem = elem.split(" (",1)[0]
			d = {}
			d['CL2 Name'] = elem
			score = []
			for name in self.name_list:
				s = SM(None, elem, name).ratio()
				score.append(s*100)
			m = max(score)
			d['Nearest CL1'] = self.name_list[score.index(m)]
			d['score'] = m
			all_scores.append(d)
		self.comparison = all_scores
		self.df_mdrive2cms_comparison = pd.DataFrame(all_scores)

if __name__ == '__main__':
	app = QApplication(sys.argv)
	window = MyApp()
	window.show()
	sys.exit(app.exec_())
