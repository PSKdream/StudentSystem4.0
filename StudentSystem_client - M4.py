import os
import io
import binascii
import sys
import tkinter as tk
import pymysql as m 
#import traits
import codecs
import tkinter 
from PIL import Image
from smartcard.System import readers
from smartcard.util import HexListToBinString, toHexString, toBytes
# Thailand ID Smartcard
from tkinter import *
from tkinter import ttk, messagebox
import webbrowser as web
from tkinter.ttk import Notebook
from tkinter.ttk import Combobox
import csv
from datetime import datetime
from tkcalendar import DateEntry
import sqlite3
from operator import itemgetter, attrgetter
#pdf
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
#printer
import win32ui
import win32api
import win32print
import win32con
import csv
gui = Tk()
gui.iconbitmap(r'iconn.ico')
gui.geometry('1080x650')
gui.resizable(width=False,height=False)
gui.title("Student Recruitment")
###########Font######################################
gui.option_add('*Font','"Angsana New" 16')
s = ttk.Style()
s.configure('my.TButton', font=('Angsana New', 16))

logo = PhotoImage(file='logo.gif').subsample(11)
guilogo = Label(gui,image=logo)
guilogo.place(x=300,y=25)

Label(gui,text='ระบบรับสมัครนักเรียนโรงเรียนพรหมานุสรณ์จังหวัดเพชรบุรี',font=('Angsana New',20,'bold')).place(x=350,y=40)
c = None
conn = None


saveIDtest = 0
cid=""
cid=""
nameTH=""
cid=""
nameEN=""
nameTH=""
Date=""
Gender=""
Address=""
mail=''
tel=''
readIDtest = None
readID=[]
tID = None
readCount = None
listRoom = None
search_Room =[]
searchRoom= []
InListRoom = None
search_RoomData=None
count = 0
cr = 1
showroomtest=''
readCounttest = None
scarchschool = None
u = 0
trueID = 0




def main_pro():
	
	c.execute(""" CREATE TABLE IF NOT EXISTS data(
			ID INTEGER PRIMARY KEY AUTO_INCREMENT,dateAdd TEXT,idTest INTEGER,room INTEGER,No_Room INTEGER,cid TEXT ,nameTH TEXT ,date1 TEXT,gender TEXT,school TEXT,oorder INTEGER,program INTEGER)""") 
	c.execute(""" CREATE TABLE IF NOT EXISTS IDtest(
				ID INTEGER PRIMARY KEY AUTO_INCREMENT,idTest INTEGER )""")
	c.execute(""" CREATE TABLE IF NOT EXISTS FIDtest(
				ID INTEGER PRIMARY KEY AUTO_INCREMENT,fidTest INTEGER )""")
	c.execute(""" CREATE TABLE IF NOT EXISTS Counttest(
				ID INTEGER PRIMARY KEY AUTO_INCREMENT,Count TEXT )""")
	c.execute(""" CREATE TABLE IF NOT EXISTS dataRoom(
				ID INTEGER PRIMARY KEY AUTO_INCREMENT,room TEXT)""")
	tebmain = tk.Toplevel() 
	tebmain.iconbitmap(r'iconn.ico')
	tebmain.geometry('1080x650')
	tebmain.resizable(width=False,height=False)
	#print(c)
	#print(conn)
	with conn:
		c.execute(""" INSERT INTO FIDtest VALUES (%s,%s)""",(None,0))
		conn.commit()		
	with conn:
		c.execute(""" SELECT * FROM IDtest""")
		readID = c.fetchall()
		if(readID != ()):
			readIDtest = int(readID[0][1])
	with conn:
		c.execute(""" SELECT * FROM data""")
		tID = c.fetchall()
	u = 0
	with conn:
		c.execute(""" SELECT * FROM FIDtest""")
		u = c.fetchall()
	trueID = 0
	def resetnum():
		with conn:
			c.execute(""" SELECT * FROM data""")
			tID = c.fetchall()
		with conn:
			c.execute(""" SELECT * FROM FIDtest""")
			u = c.fetchall()
		global trueID
		trueID = u[0][1]
		ch = 0
		while True :
			check = True
			for j in tID :
				if(trueID+ch == int(j[2])) :
					check = False
					#print("nooo")
					break
			if(check==True) :
				break;
			if(check==False) :
				ch = ch + 1
		trueID += ch
	if(u[0][1]!=0) :
		resetnum()
	with conn:
		c.execute(""" SELECT * FROM Counttest""")
		readCount = c.fetchall()
		if(readCount != ()):
			readCounttest = int(readCount[0][1])
	'''
	mainmenu =Menu(tebmain)
#----------------------file--------------------------
	menufile=Menu(mainmenu,tearoff=0)
	menufile.add_command(label="Export CSV File",)
	menufile.add_command(label="ตั้งค่าเลขประจำตัวผู้สมัครสอบ",)
	menufile.add_command(label="ตั้งค่าห้องสอบ",)
	menufile.add_command(label="Close",)
	menuabout = Menu(mainmenu,tearoff=0)
	menuabout.add_command(label="About Us",)
	menuabout.add_command(label="Goto Facebook",)
	#------------------set menu
	mainmenu.add_cascade(label='File',menu=menufile)
	mainmenu.add_cascade(label='About',menu=menuabout)
	gui.config(menu=mainmenu)'''
###############################teb######################################################
	#------------------tab
	Teb = ttk.Notebook(tebmain)
	tebProfile =Frame(Teb)
	tebManage=Frame(Teb)
	tebViwe=Frame(Teb)
	tebSum=Frame(Teb)
	#tebdata=Frame(Teb)
	#tebEdit=Frame(Teb)
	#tebRoom=Frame(Teb)
	
	#Teb.add(tebdata,text="จัดการ Database")
	Teb.add(tebProfile,text="กรอกข้อมูล")
	Teb.add(tebManage,text='แก้ไขข้อมูล')
	Teb.add(tebViwe,text="View")
	Teb.add(tebSum,text='สรุป')
	#Teb.add(tebEdit,text='ตั้งค่าเลขประจำตัวผู้สมัครสอบ')
	#Teb.add(tebRoom,text='ตั้งค่าห้องสอบ')
	Teb.pack(fill=BOTH,expand=1) #fill X ,V BOTH	
###################################ปุ่มหน้าแรก################################################
	def EditBT():
		EntryID.config(state='enabled')
		EntryNameTH.config(state='enabled')
		SHDate.config(state='enabled')
		SHGender.config(state='enabled')
	def cleanBT():
		VarshowID.set('')
		VarshowNameTH.set('')
		VarshowDate.set('')
		VarshowGender.set('')
		Varshowtest.set('')
		Varshowroomtest.set('')
		VarshowOrder.set('')
		SHschool.delete(first=0,last=100)
		Varshowroom.set('')
		EntryID.config(state='disabled')
		EntryNameTH.config(state='disabled')
		SHDate.config(state='disabled')
		SHGender.config(state='disabled')
	def readcardBT():
		def thai2unicode(data):
			result = ''
			#print (bytes(data))
			#print(bytes(data).decode('tis-620'))
			result = bytes(data).decode('tis-620')
			'''
			if isinstance(data, list):
				for d in data:
					
					result += bytes(d)'''
			return result.strip();

		def getData(cmd, req = [0x00, 0xc0, 0x00, 0x00]):
			data, sw1, sw2 = connection.transmit(cmd)
			data, sw1, sw2 = connection.transmit(req + [cmd[-1]])
			return [data, sw1, sw2];

		# define the APDUs used in this script
		# https://github.com/chakphanu/ThaiNationalIDCard/blob/master/APDU.md

		###########################read idtest##################################

		with conn:
			c.execute(""" SELECT * FROM IDtest""")
			readID = c.fetchall()
			if(readID != ()):
				readIDtest = int(readID[0][1])
			#print(readIDtest)

		# Check card
		SELECT = [0x00, 0xA4, 0x04, 0x00, 0x08]
		THAI_CARD = [0xA0, 0x00, 0x00, 0x00, 0x54, 0x48, 0x00, 0x01]
		# CID
		CMD_CID = [0x80, 0xb0, 0x00, 0x04, 0x02, 0x00, 0x0d]
		# TH Fullname
		CMD_THFULLNAME = [0x80, 0xb0, 0x00, 0x11, 0x02, 0x00, 0x64]
		# EN Fullname
		CMD_ENFULLNAME = [0x80, 0xb0, 0x00, 0x75, 0x02, 0x00, 0x64]
		# Date of birth
		CMD_BIRTH = [0x80, 0xb0, 0x00, 0xD9, 0x02, 0x00, 0x08]
		# Gender
		CMD_GENDER = [0x80, 0xb0, 0x00, 0xE1, 0x02, 0x00, 0x01]
		# Card Issuer
		CMD_ISSUER = [0x80, 0xb0, 0x00, 0xF6, 0x02, 0x00, 0x64]
		# Issue Date
		CMD_ISSUE = [0x80, 0xb0, 0x01, 0x67, 0x02, 0x00, 0x08]
		# Expire Date
		CMD_EXPIRE = [0x80, 0xb0, 0x01, 0x6F, 0x02, 0x00, 0x08]
		# Address
		CMD_ADDRESS = [0x80, 0xb0, 0x15, 0x79, 0x02, 0x00, 0x64]
		# Get all the available readers
		readerList = readers()
		#print ('Available readers:')
		# Select reader
		readerSelectIndex = 0 #int(input("Select reader[0]: ") or "0")
		try:
			print(readerList[readerSelectIndex])
		except:
			messagebox.showwarning('แจ้งเตือน','กรุณาทำการต่อเครื่องอ่านบัตร',parent=tebmain)
		else:
			reader = readerList[readerSelectIndex]
			#print ("Using:", reader)
			connection = reader.createConnection()
			try:
			  connection.connect()	
			except:
			  messagebox.showwarning('แจ้งเตือน','กรุณาเสียบบัตรประชาชน หรือทำความสะอาด chip',parent=tebmain)
			else:
				connection.connect()
				atr = connection.getATR()
				#print (atr)
				#print ("ATR: " + toHexString(atr))
				if (atr[0] == 0x3B & atr[1] == 0x67):
					req = [0x00, 0xc0, 0x00, 0x01]
				else :
					req = [0x00, 0xc0, 0x00, 0x00]

				# Check card

				data, sw1, sw2 = connection.transmit(SELECT + THAI_CARD)
				#print ("Select Applet: %02X %02X" % (sw1, sw2))

				# CID
				data = getData(CMD_CID, req)
				cid = thai2unicode(data[0])
				#print ("CID: " + cid)

				# TH Fullname

				data = getData(CMD_THFULLNAME, req)
				nameTH = ''
				for i in thai2unicode(data[0]):
					if(i == "#"):
						nameTH += " "
					elif ( i != "#" ) :
						nameTH += i
				#print ("TH Fullname: " +  thai2unicode(data[0]))
				#print(thai2unicode2(data[0])))

				# EN Fullname
				data = getData(CMD_ENFULLNAME, req)
				nameEN = ''
				for i in thai2unicode(data[0]):
					if(i == "#"):
						nameEN += " "
					elif ( i != "#" ) :
						nameEN += i
				#print(nameEN)
				#print ("EN Fullname: " + thai2unicode(data[0]))

				# Date of birth
				data = getData(CMD_BIRTH, req)
				#print( "Date of birth: " + thai2unicode(data[0]))
				birth=thai2unicode(data[0])
				Date = birth[6]+birth[7]+"/"+birth[4]+birth[5]+'/'+birth[0]+birth[1]+birth[2]+birth[3]
				#print(Date)
				# Gender
				data = getData(CMD_GENDER, req)
				if(thai2unicode(data[0]) == "1"):
					Gender = "ชาย"
				elif (thai2unicode(data[0]) == "2"):
					Gender = "หญิง"
				else:
					Gender = "Error"
				#print ("Gender: " + thai2unicode(data[0]))

				# Card Issuer
				data = getData(CMD_ISSUER, req)
				#print ("Card Issuer: " + thai2unicode(data[0]))

				# Issue Date
				data = getData(CMD_ISSUE, req)
				#print ("Issue Date: " + thai2unicode(data[0]))
				IssueDate=thai2unicode(data[0])
				#print('Issue Date: %c%c/%c%c/%c%c%c%c'%(IssueDate[6],IssueDate[7],IssueDate[4],IssueDate[5],IssueDate[0],IssueDate[1],IssueDate[2],IssueDate[3],))

				# Expire Date
				data = getData(CMD_EXPIRE, req)
				#print ("Expire Date: " + thai2unicode(data[0]))
				ExpireDate=thai2unicode(data[0])
				#print('Expire Date: %c%c/%c%c/%c%c%c%c'%(ExpireDate[6],ExpireDate[7],ExpireDate[4],ExpireDate[5],ExpireDate[0],ExpireDate[1],ExpireDate[2],ExpireDate[3],))

				# Address
				data = getData(CMD_ADDRESS, req)
				Address=''
				for i in thai2unicode(data[0]):
					if(i == "#"):
						Address += " "
					elif ( i != "#" ) :
						Address += i

				with conn:
					c.execute(""" SELECT * FROM dataRoom""")			
					check = c.fetchall()
					#print(check)
				with conn:
					c.execute(""" SELECT * FROM Counttest""")
					readCount = c.fetchall()
					if(readCount != ()):
						readCounttest = int(readCount[0][1])
				with conn:
					c.execute(""" SELECT * FROM FIDtest""")
					saveIDtest = c.fetchall()
				with conn:
					c.execute(""" SELECT * FROM data""")
					datach = c.fetchall()
					checkdata = 1
					if(datach == ()) :
						checkdata=0
				try:
					print(int(readCounttest))	
				except:
				  messagebox.showwarning('แจ้งเตือน','กรุณากรอกจำนวนผู้เข้าสอบแต่ละห้อง',parent=tebmain)
				else:
					for i in check:
						counti = 0
						if(checkdata!=0) :
							for j in datach:
								if(str(i[1])==str(j[3])):
									counti = counti +1
						if(counti<readCounttest) :
							readRoomShowtest = i[1]
							break
					try:
						print(readRoomShowtest)	
					except:
					  messagebox.showwarning('แจ้งเตือน','มีจำนวนห้องไม่พอ กรุณาเพิ่มห้องสอบ หรือตรวจสอบข้อมูล',parent=tebmain)
					else:
						resetnum()
						Varshowroomtest.set(readRoomShowtest)
						VarshowID.set(cid)
						VarshowNameTH.set(nameTH)
						VarshowDate.set(Date)
						VarshowGender.set(Gender)
						Refresh()
					return[cid,nameEN,nameTH,Date,Gender,Address,trueID];
	def insert_data (c1,c2,c3,c4,c5,c6,c7,c8,c9,c10):
		dt = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
		with conn:
			c.execute(""" INSERT INTO data VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) """,(None,dt,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10))
			conn.commit()
		#print("insert")
	def add_data():
		global trueID
		if (Varshowroomtest.get()==""or Varshowtest.get()=="" or VarshowID.get()=="" 
			or VarshowNameTH.get()=="" or SHschool.get()=='' 
			 or VarshowDate.get()==''or VarshowGender.get()=='' or CheckVar.get() == 0):
			messagebox.showwarning('แจ้งเตือน','กรุณากรอกข้อมูลให้ครับถ้วน',parent=tebmain)

		else:
			#print(str(VarshowID.get()))
			with conn:
				c.execute("""SELECT * FROM data WHERE cid LIKE '%s'  """%VarshowID.get())
				search = c.fetchall()
				#print(search)
				if(search == ()):
					with conn:
						c.execute("""SELECT * FROM data WHERE idTest LIKE '%s'  """%Varshowtest.get())
						search = c.fetchall()
					if(search==()):
						with conn:
							c.execute(""" SELECT * FROM FIDtest""")
							saveIDtest = c.fetchall()
						with conn:
							c.execute(""" SELECT * FROM IDtest""")
							readID = c.fetchall()
							readIDtest1 = int(readID[0][1])
						with conn:
							c.execute(""" SELECT * FROM Counttest""")
							readCount = c.fetchall()
							if(readCount != ()):
								#print ("aaaaaaaaaaaaaaaa");
								readCounttest = int(readCount[0][1])
						order = ((int(trueID)-int(saveIDtest[0][1]))%int(readCounttest))+1
						insert_data(int(trueID),Varshowroomtest.get(),Varshowroom.get(),VarshowID.get(),VarshowNameTH.get()
							,VarshowDate.get(),VarshowGender.get(),SHschool.get(),order,CheckVar.get())
						with conn:
							readIDtestSUM = int(trueID)+1
							c.execute("""UPDATE IDtest set idTest = %s WHERE %s """,(str(readIDtestSUM),str(trueID)))
							conn.commit()
						messagebox.showinfo('แจ้งเตือน','บันทึกข้อมูลเรียบร้อย',parent=tebmain)
						'''
						with conn:
							c.execute(""" SELECT * FROM dataRoom""")			
							check = c.fetchall()
							#print(check)
						for i in check:
							print(i[1])
							print(int(Varshowroomtest.get()))
							if(int(i[1])==int(Varshowroomtest.get())):
								with conn:
									c.execute("""UPDATE dataRoom set checkR = %d WHERE room = %s """%(i[2]+1,str(i[1])))
									conn.commit()
									break
						#print(check)'''
						pdfPrint()
					else:
						messagebox.showwarning('แจ้งเตือน','เลขประจำตัวผู้สมัครสอบซ้ำ\nกรุณากดปุ่ม Refresh',parent=tebmain)
				else: 
					messagebox.showwarning('แจ้งเตือน','มีข้อมูลในระบบแล้ว',parent=tebmain)
	def pdfPrint():
		with conn:
				c.execute("""SELECT * FROM data WHERE cid LIKE '%s'  """ %VarshowID.get())
				search = c.fetchall()
		with conn:
				c.execute("""SELECT * FROM data WHERE idTest LIKE %s  """ %Varshowtest.get())
				search2 = c.fetchall()
		if search != () :
			with conn:
				c.execute(""" SELECT room FROM dataRoom""")
				searchRoom = c.fetchall()
				#search = c.fetchone()
				#search = c.fetchmany()
				#print(searchRoom)
				search_Room.clear()
				for g in searchRoom:
				   cutRoom = str(g).split("'")
				   search_Room.append(cutRoom[1])
				search_Room.sort()
		if (search != () and search2 != ()) :
			packet = io.BytesIO()
			#font
			pdfmetrics.registerFont(TTFont('THSarabun', 'THSarabun.ttf'))
			# create a new PDF with Reportlab

			can = canvas.Canvas(packet, pagesize=letter)
			can.setFont('THSarabun', 16)
			can.drawString(100, 645, SHschool.get())
			can.drawString(80, 665, VarshowNameTH.get())
			can.drawString(170, 626, Varshowtest.get())
			a= search_Room.index(str(SHroomtest.get()))
			a=int(a)+1
			can.drawString(100, 608, str(a))
			can.drawString(155, 608, Varshowroomtest.get())
			can.drawString(230, 608, OHroomtest.get())
			can.drawString(390, 645, SHschool.get())
			can.drawString(370, 665, VarshowNameTH.get())
			can.drawString(460, 626, Varshowtest.get())
			can.drawString(385, 608, str(a))
			can.drawString(443, 608, Varshowroomtest.get())
			can.drawString(523, 608, OHroomtest.get())
			thai_year = datetime.now().year + 543
			dt = datetime.now().strftime('%d / %m /')
			can.setFont('THSarabun', 14)
			can.drawString(488,  432, dt)
			can.drawString(528,  432, str(thai_year))
			can.drawString(193,  432, dt)
			can.drawString(235,  432, str(thai_year))
			if CheckVar.get() == 1:
				can.drawString(410, 590, 'คณิตศาสตร์ - วิทยาศาสตร์')
				can.drawString(115, 590, 'คณิตศาสตร์ - วิทยาศาสตร์')
			elif CheckVar.get() == 2:
				can.drawString(410, 589, "คณิตศาสตร์ - เทคโนโลยี")
				can.drawString(115, 590, 'คณิตศาสตร์ - เทคโนโลยี')
			elif CheckVar.get() == 3:
				can.drawString(410, 590, 'ภาษาต่างประเทศ')
				can.drawString(115, 590, 'ภาษาต่างประเทศ')
			elif CheckVar.get() == 4:
				can.drawString(410, 590, 'ทั่วไป')
				can.drawString(115, 590, 'ทั่วไป')


			can.save()

			#move to the beginning of the StringIO buffer
			packet.seek(0)
			new_pdf = PdfFileReader(packet)
			# read your existing PDF
			existing_pdf = PdfFileReader(open("mypdf.pdf", "rb"))
			output = PdfFileWriter()
			# add the "watermark" (which is the new pdf) on the existing page
			page = existing_pdf.getPage(0)
			page.mergePage(new_pdf.getPage(0))
			output.addPage(page)
			# finally, write "output" to a real file
			try :
				outputStream = open("destination.pdf", "wb")
			except:
				messagebox.showinfo('แจ้งเตือน','กรุณาปิด PDF',parent=tebmain)
			else :
				output.write(outputStream)
				outputStream.close()
				os.startfile("destination.pdf")
				messagebox.showinfo('แจ้งเตือน','พิมพ์ไฟล์แล้ว',parent=tebmain)
		else :
			messagebox.showwarning('แจ้งเตือน','กรุณาบันทึกข้อมูลลงระบบก่อนสั่งพิมพ์',parent=tebmain)
	def Refresh():
		with conn:
			 c.execute(""" SELECT * FROM IDtest""")
			 readID = c.fetchall()
			 if(readID != ()):
				 readIDtest = int(readID[0][1])
			 c.execute(""" SELECT * FROM dataRoom""")			
			 check = c.fetchall()
			 #print(check)
		with conn:
			c.execute(""" SELECT * FROM Counttest""")
			readCount = c.fetchall()
			if(readCount != ()):
				readCounttest = int(readCount[0][1])
		with conn:
				c.execute(""" SELECT room FROM dataRoom""")
				searchRoom = c.fetchall()
				#search = c.fetchone()
				#search = c.fetchmany()
				#print(searchRoom)
				search_Room.clear()
				for g in searchRoom:
				   cutRoom = str(g).split("'")
				   search_Room.append(cutRoom[1])
				search_Room.sort()
		checkdata=1
		with conn:
				c.execute(""" SELECT * FROM data""")
				datach = c.fetchall()
				if(datach==()) :
					checkdata=0
		try:
			 print(int(readCounttest))	
		except:
			 messagebox.showwarning('แจ้งเตือน','กรุณากรอกจำนวนผู้เข้าสอบแต่ละห้อง',parent=tebmain)
		else:
			 for i in check:
				 counti = 0
				 if(checkdata!=0) :
					 for j in datach:
						 if(i[1]==j[3]):
							 counti = counti + 1
				 if(counti<readCounttest) :
					 readRoomShowtest = i[1]
					 break
		try:
			 print(readRoomShowtest)	
		except:
			 messagebox.showwarning('แจ้งเตือน','มีจำนวนห้องไม่พอ กรุณาเพิ่มห้องสอบ หรือตรวจสอบข้อมูล',parent=tebmain)
		else:
			 resetnum()
			 #print(readRoomShowtest)
			 global trueID
			 with conn:
					c.execute(""" SELECT * FROM FIDtest""")
					saveIDtest = c.fetchall()
			 with conn:
					c.execute(""" SELECT * FROM dataRoom""")
					Room = c.fetchall()
			 b = (int(trueID)-int(saveIDtest[0][1]))/int(readCounttest)
			 Varshowroomtest.set(Room[int(b)][1])		
			 Varshowtest.set(trueID)
			 with conn:
					c.execute(""" SELECT * FROM FIDtest""")
					saveIDtest = c.fetchall()
			 VarshowOrder.set((int(trueID)-int(saveIDtest[0][1]))%int(readCounttest)+1)
			# print(int(trueID))
			# print(saveIDtest)
			# print(int(readCounttest))
			 with conn:
			 	c.execute(""" SELECT room FROM dataRoom""")
			 	search_Room.clear()
			 	for g in searchRoom:
				   cutRoom = str(g).split("'")
				   search_Room.append(cutRoom[1])
			 search_Room.sort()
			 a= search_Room.index(str(SHroomtest.get()))
			 a=int(a)+1
			 Varshowroom.set(str(a))


	NProgram = Label(tebProfile,text='แผนการเรียน',font=('Angsana New',20,'bold'))
	NProgram.place(x=20,y=330)
	CheckVar = IntVar()
	C1 = Radiobutton(tebProfile, text="แผน 1",variable=CheckVar, value=1,font=('Angsana New',20,'bold'))
	C2 = Radiobutton(tebProfile, text="แผน 2",variable=CheckVar, value=2,font=('Angsana New',20,'bold'))
	C3 = Radiobutton(tebProfile, text="แผน 3",variable=CheckVar, value=3,font=('Angsana New',20,'bold'))
	C4 = Radiobutton(tebProfile, text="แผน 4",variable=CheckVar, value=4,font=('Angsana New',20,'bold'))
	C1.place(x=290,y=325)
	C2.place(x=390,y=325)
	C3.place(x=490,y=325)
	C4.place(x=590,y=325)

	BTread = Frame(tebProfile,height=20,width=40) #set location
	BTread.place(x=880,y=80)
	Bread = ttk.Button(BTread,text='อ่านข้อมูลบัตร',style='my.TButton',command=readcardBT)
	Bread.pack(ipadx=25,ipady=25)


	BTclean = Frame(tebProfile,height=20,width=40) #set location
	BTclean.place(x=880,y=180)
	Bclean = ttk.Button(BTclean,text='ล้าง',style='my.TButton',command=cleanBT)
	Bclean.pack(ipadx=25,ipady=25)

	BTsave = Frame(tebProfile,height=20,width=40) #set location
	BTsave.place(x=880,y=280)
	Bsave = ttk.Button(BTsave,text='บันทึกข้อมูล',style='my.TButton',command=add_data)
	Bsave.pack(ipadx=25,ipady=25)

	BTPrint = Frame(tebProfile,height=20,width=40)
	BTPrint.place(x=880,y=380)
	Print = ttk.Button(BTPrint,text= 'แก้ไขข้อมูล',style='my.TButton',command=EditBT)
	Print.pack(ipadx=25,ipady=25)

	BTRe = Frame(tebProfile,height=10,width=40)
	BTRe.place(x=520,y=450)
	Re = ttk.Button(BTRe,text= 'Refresh',style='my.TButton',command=Refresh)
	Re.pack(ipadx=5,ipady=5)

	BTPrint = Frame(tebProfile,height=20,width=40)
	BTPrint.place(x=880,y=480)
	Print = ttk.Button(BTPrint,text= 'พิมพ์',style='my.TButton',command=pdfPrint)
	Print.pack(ipadx=25,ipady=25)

	######################################################

	guilogo5 = Label(tebProfile,image=logo)
	guilogo5.place(x=300,y=5)

	Label(tebProfile,text='ระบบรับสมัครนักเรียนโรงเรียนพรหมานุสรณ์จังหวัดเพชรบุรี',font=('Angsana New',20,'bold')).place(x=350,y=20)

	TextID = Label(tebProfile,text='เลขประจำตัวประชาชน',font=('Angsana New',20,'bold'))
	TextID.place(x=20,y=80)
	VarshowID = StringVar()
	EntryID = ttk.Entry(tebProfile,textvariable=VarshowID,font=('Angsana New',18,'bold'),width=60,state='disabled')
	EntryID.place(x=300,y=80)

	TextNameTH = Label(tebProfile,text='ชื่อ - นามสกุล',font=('Angsana New',20,'bold'),)
	TextNameTH.place(x=20,y=130)
	VarshowNameTH = StringVar()
	EntryNameTH = ttk.Entry(tebProfile,textvariable=  VarshowNameTH,font=('Angsana New',18,'bold'),width=60,state='disabled')
	EntryNameTH.place(x=300,y=130)

	NDate = Label(tebProfile,text='วว/ดด/ปปปป',font=('Angsana New',20,'bold'))
	NDate.place(x=20,y=180)
	VarshowDate = StringVar()
	SHDate = ttk.Entry(tebProfile,textvariable=  VarshowDate,font=('Angsana New',18,'bold'),width=60,state='disabled')
	SHDate.place(x=300,y=180)

	NGender = Label(tebProfile,text='เพศ',font=('Angsana New',20,'bold'))
	NGender.place(x=20,y=230)
	VarshowGender = StringVar()
	SHGender = ttk.Entry(tebProfile,textvariable=  VarshowGender,font=('Angsana New',18,'bold'),width=60,state='disabled')
	SHGender.place(x=300,y=230)

	Nmail = Label(tebProfile,text='โรงเรียน',font=('Angsana New',20,'bold'))
	Nmail.place(x=20,y=280)
	'''
	Varshowmail = StringVar()
	SHmail = ttk.Entry(tebProfile,textvariable=  Varshowmail,font=('Angsana New',18,),width=60)
	SHmail.place(x=300,y=280)'''
	school=['พรหมานุสรณ์จังหวัดเพชรบุรี','วัดจันทราวาส (ศุขประสารราษฎร์)','เบญจมเทพอุทิศจังหวัดเพชรบุรี','คงคาราม','โยธินบูรณะ เพชรบุรี','จุฬาภรณราชวิทยาลัย เพชรบุรี','อรุณวิทยา','ราษฎร์วิทยา','รักษ์วิทยา','หัวหินวิทยาลัย','ประจวบวิทยาลัย','อรุณประดิษฐ']
	school.sort()
	SHschool = ttk.Combobox(tebProfile,values=school,font=('Angsana New',18,'bold'),width=60)#,state='readonly')
	SHschool.place(x=300,y=280)

	NShowroomtest = Label(tebProfile,text='ห้อง',font=('Angsana New',25,'bold'),fg='white',bg='red')
	NShowroomtest.place(x=520,y=380)
	Varshowroomtest = StringVar()
	SHroomtest = ttk.Entry(tebProfile,textvariable=Varshowroomtest,font=('Angsana New',18,'bold'),width=20,state='disabled')
	SHroomtest.place(x=600,y=385)

	room = Label(tebProfile,text='ห้องที่',font=('Angsana New',25,'bold'),fg='white',bg='red')
	room.place(x=220,y=380)
	Varshowroom = StringVar()
	SHroom = ttk.Entry(tebProfile,textvariable=Varshowroom,font=('Angsana New',18,'bold'),width=20,state='disabled')
	SHroom.place(x=320,y=385)

	NShowordertest = Label(tebProfile,text='ลำดับที่',font=('Angsana New',25,'bold'),fg='white',bg='red')
	NShowordertest.place(x=220,y=450)
	VarshowOrder = StringVar()
	OHroomtest = ttk.Entry(tebProfile,textvariable=VarshowOrder,font=('Angsana New',18,'bold'),width=20,state='disabled')
	OHroomtest.place(x=320,y=460)

	Ntest = Label(tebProfile,text='เลขประจำตัวผู้สมัครสอบ',font=('Angsana New',25,'bold'),fg='white',bg='red')
	Ntest.place(x=220,y=520)
	Varshowtest = StringVar()
	SHtest = ttk.Entry(tebProfile,textvariable=Varshowtest,font=('Angsana New',18,'bold'),width=15,state='disabled')
	SHtest.place(x=590,y=528)

	def Be():
		#print("a")
		with conn:
			c.execute(""" SELECT * FROM FIDtest""")
			saveIDtest = c.fetchall()
		global trueID
		with conn:
			c.execute(""" SELECT * FROM data""")
			tID = c.fetchall()
		ch = 1
		while True :
			a = 0
			for j in tID :
				if((trueID-ch)==j[2]) :
					a = 1
					break
			if(int(trueID-ch)<int(saveIDtest[0][1])) :
				#print(trueID-ch)
				#print(saveIDtest[0][1])
				ch = 0
				break
			if(a==0) :
				break
			if(a==1) :
				ch = ch +1 
		trueID -= ch
		Varshowtest.set(trueID)		
		#SHtest.place(x=590,y=508)
		with conn:
					c.execute(""" SELECT * FROM Counttest""")
					readCount = c.fetchall()
					if(readCount != ()):
						readCounttest = int(readCount[0][1])
		with conn:
					c.execute(""" SELECT * FROM dataRoom""")
					Room = c.fetchall()
		a = (int(trueID)-int(saveIDtest[0][1]))%int(readCounttest)+1
		VarshowOrder.set(int(a))
		#OHroomtest = ttk.Entry(tebProfile,textvariable=showOrder,font=('Angsana New',18,'bold'),width=10,state='disabled')
		#OHroomtest.place(x=650,y=433)
		b = (int(trueID)-int(saveIDtest[0][1]))/int(readCounttest)
		Varshowroomtest.set(Room[int(b)][1])
		#SHroomtest = ttk.Entry(tebProfile,textvariable=showroomtest,font=('Angsana New',18,'bold'),width=10,state='disabled')
		#SHroomtest.place(x=360,y=433)
		Varshowroom.set(int(b)+1)
	def Nex():
		#print("b")
		count = 0;
		with conn:
					c.execute(""" SELECT * FROM Counttest""")
					alla = c.fetchall()
		with conn:
					c.execute(""" SELECT * FROM dataRoom""")
					Room = c.fetchall()
		for i in Room :
			if(int(i[0])>0)	:
				count = count + 1
		'''print("count = ")	
		print(count)
		print("alla = ")	
		print(alla[0][1])
		print("time = ")
		print(int(count)*int(alla[0][1]))'''
		with conn:
					c.execute(""" SELECT * FROM FIDtest""")
					saveIDtest = c.fetchall()
		global trueID
		#print("Tue =")
		#print(trueID)
		with conn:
			c.execute(""" SELECT * FROM data""")
			tID = c.fetchall()
		ch = 1
		while True :
			a = 0
			for j in tID :
				if((trueID+ch)==j[2]) :
					a = 1
					break
			if( int(trueID+ch) >= int(saveIDtest[0][1])+int(count)*int(alla[0][1])) :
				#print(int(trueID+ch))
				#print(int(count)*int(alla[0][1]))
				ch = 0
			if(a==0) :
				break
			if(a==1) :
				ch = ch +1 
		trueID += ch
		Varshowtest.set(trueID)		
		#SHtest.place(x=590,y=508)
		with conn:
					c.execute(""" SELECT * FROM Counttest""")
					readCount = c.fetchall()
					if(readCount != ()):
						readCounttest = int(readCount[0][1])
		a = (int(trueID)-int(saveIDtest[0][1]))%int(readCounttest)+1
		VarshowOrder.set(int(a))
		#OHroomtest = ttk.Entry(tebProfile,textvariable=VarshowOrder,font=('Angsana New',18,'bold'),width=10,state='disabled')
		#OHroomtest.place(x=650,y=433)
		b = (int(trueID)-int(saveIDtest[0][1]))/int(readCounttest)
		#print(Room[int(b)][1])
		Varshowroomtest.set(Room[int(b)][1])
		#SHroomtest = ttk.Entry(tebProfile,textvariable=Varshowroomtest,font=('Angsana New',18,'bold'),width=10,state='disabled')
		#SHroomtest.place(x=360,y=433)
		Varshowroom.set(int(b)+1)


	BTPrint = Frame(tebProfile,height=5,width=5)
	BTPrint.place(x=480,y=528)
	Print = ttk.Button(BTPrint,text= 'ก่อนหน้า',style='my.TButton',command=Be)
	Print.pack(ipadx=0,ipady=0)

	BTPrint = Frame(tebProfile,height=5,width=5)
	BTPrint.place(x=750,y=528)
	Print = ttk.Button(BTPrint,text= 'ถัดไป',style='my.TButton',command=Nex)
	Print.pack(ipadx=0,ipady=0)
#################################จัดการข้อมูล##########################################
	guilogo2 = Label(tebManage,image=logo)
	guilogo2.place(x=300,y=5)

	Label(tebManage,text='ระบบรับสมัครนักเรียนโรงเรียนพรหมานุสรณ์จังหวัดเพชรบุรี',font=('Angsana New',20,'bold')).place(x=350,y=20)
	###################################Button###########################
	def pdfPrint2():
		with conn:
				c.execute("""SELECT * FROM data WHERE cid LIKE '%s'  """ %VarshowID2.get())
				search = c.fetchall()
				#print(search)
		with conn:
				c.execute("""SELECT * FROM data WHERE idTest LIKE %s  """ %Varshowtest2.get())
				search2 = c.fetchall()
				#print(search2)
		if search != () :
			with conn:
				c.execute(""" SELECT room FROM dataRoom""")
				searchRoom = c.fetchall()
				#search = c.fetchone()
				#search = c.fetchmany()
				#print(searchRoom)
				search_Room.clear()
				for g in searchRoom:
				   cutRoom = str(g).split("'")
				   search_Room.append(cutRoom[1])
				search_Room.sort()
		if (search != () and search2 != ()) :
			packet = io.BytesIO()
			#font
			pdfmetrics.registerFont(TTFont('THSarabun', 'THSarabun.ttf'))
			# create a new PDF with Reportlab

			can = canvas.Canvas(packet, pagesize=letter)
			can.setFont('THSarabun', 16)
			can.drawString(100, 645, SHschool2.get())
			can.drawString(80, 665, VarshowNameTH2.get())
			can.drawString(170, 626, Varshowtest2.get())
			a= search_Room.index(str(SHroomtest2.get()))
			a=int(a)+1
			can.drawString(100, 608, str(a))
			can.drawString(155, 608, Varshowroomtest2.get())
			can.drawString(230, 608, OHroomtest2.get())
			can.drawString(390, 645, SHschool2.get())
			can.drawString(370, 665, VarshowNameTH2.get())
			can.drawString(460, 626, Varshowtest2.get())
			can.drawString(385, 608, str(a))
			can.drawString(443, 608, Varshowroomtest2.get())
			can.drawString(523, 608, OHroomtest2.get())
			thai_year = datetime.now().year + 543
			dt = datetime.now().strftime('%d / %m /')
			can.setFont('THSarabun', 14)
			can.drawString(488,  432, dt)
			can.drawString(528,  432, str(thai_year))

			can.drawString(193,  432, dt)
			can.drawString(235,  432, str(thai_year))
			
			if CheckVar2.get() == 1:
				can.drawString(410, 590, 'คณิตศาสตร์ - วิทยาศาสตร์')
				can.drawString(115, 590, 'คณิตศาสตร์ - วิทยาศาสตร์')
			elif CheckVar2.get() == 2:
				can.drawString(410, 589, "คณิตศาสตร์ - เทคโนโลยี")
				can.drawString(115, 590, 'คณิตศาสตร์ - เทคโนโลยี')
			elif CheckVar2.get() == 3:
				can.drawString(410, 590, 'ภาษาต่างประเทศ')
				can.drawString(115, 590, 'ภาษาต่างประเทศ')
			elif CheckVar2.get() == 4:
				can.drawString(410, 590, 'ทั่วไป')
				can.drawString(115, 590, 'ทั่วไป')

			can.save()

			#move to the beginning of the StringIO buffer
			packet.seek(0)
			new_pdf = PdfFileReader(packet)
			# read your existing PDF
			existing_pdf = PdfFileReader(open("mypdf.pdf", "rb"))
			output = PdfFileWriter()
			# add the "watermark" (which is the new pdf) on the existing page
			page = existing_pdf.getPage(0)
			page.mergePage(new_pdf.getPage(0))
			output.addPage(page)
			# finally, write "output" to a real file
			try :
				outputStream = open("destination.pdf", "wb")
			except:
				messagebox.showinfo('แจ้งเตือน','กรุณาปิด PDF',parent=tebmain)
			else :
				output.write(outputStream)
				outputStream.close()
				os.startfile("destination.pdf")
				messagebox.showinfo('แจ้งเตือน','พิมพ์ไฟล์แล้ว',parent=tebmain)
		else :
			messagebox.showwarning('แจ้งเตือน','กรุณาบันทึกข้อมูลลงระบบก่อนสั่งพิมพ์',parent=tebmain)
	def deletedata():
		with conn:
			c.execute(""" SELECT * FROM data WHERE cid LIKE '%s' """%VarshowID2.get())
			data = c.fetchall()
			#print(data)
		with conn:
			c.execute(""" SELECT * FROM dataRoom WHERE room LIKE %s """%(data[0][3]))
			dataro = c.fetchall()
			#print(dataro)
		with conn:
			c.execute(""" DELETE FROM data WHERE cid = '%s' """%VarshowID2.get())
			'''
		dataR = int(dataro[0][2])-1
		print(dataR)
		print(data[0][3])
		with conn:
			c.execute(""" UPDATE dataRoom SET checkR  = %d WHERE room = %s """%(dataR,str(data[0][3])))
			#search = c.fetchone()'''
			#search = c.fetchmany()
			#print(data)
		messagebox.showwarning('แจ้งเตือน','ลบข้อมูลในระบบบแล้ว',parent=tebmain)
	def cleanBT2():
		VarshowID2.set('')
		VarshowNameTH2.set('')
		VarshowDate2.set('')
		VarshowGender2.set('')
		Varshowtest2.set('')
		Varshowroomtest2.set('')
		VarshowOrder2.set('')
		SHschool2.delete(first=0,last=100)
		Varshowroom2.set('')
	def saveEdit():
		if (Varshowroomtest2.get()==""or Varshowtest2.get()=="" or VarshowID2.get()=="" 
			or VarshowNameTH2.get()=="" or SHschool2.get()=='' 
			 or VarshowDate2.get()==''or VarshowGender2.get()=='' or CheckVar2.get() == 0):
				messagebox.showwarning('แจ้งเตือน','กรุณากรอกข้อมูลให้ครับถ้วน',parent=tebmain)

		else:
			with conn:
				c.execute("""UPDATE data set nameTH=%s,date1=%s,gender=%s,school=%s WHERE cid = %s """
					,(VarshowNameTH2.get(),VarshowDate2.get(),VarshowGender2.get()
					,SHschool2.get(),VarshowID2.get()))
				c.execute("""UPDATE data set program = %s WHERE cid = %s """,(CheckVar2.get(),VarshowID2.get()))
				conn.commit()
			messagebox.showinfo('แจ้งเตือน','บันทึกข้อมูลเรียบร้อย',parent=tebmain)

	BTsave22 = Frame(tebManage,height=20,width=40) #set location
	BTsave22.place(x=880,y=300)
	Bsave22 = ttk.Button(BTsave22,text='บันทึกการแก้ข้อมูล',style='my.TButton',command=saveEdit)
	Bsave22.pack(ipadx=25,ipady=20)

	BTclean = Frame(tebManage,height=20,width=40) #set location
	BTclean.place(x=880,y=400)
	Bclean = ttk.Button(BTclean,text='ล้าง',style='my.TButton',command=cleanBT2)
	Bclean.pack(ipadx=25,ipady=25)
	
	BTFind = Frame(tebManage,height=20,width=40) #set location
	BTFind.place(x=880,y=400)
	BF = ttk.Button(BTFind,text='ลบข้อมูลทั้งหมด',style='my.TButton',command=deletedata,)
	BF.pack(ipadx=25,ipady=25)
   
	BTclean2 = Frame(tebManage,height=20,width=40) #set location
	BTclean2.place(x=880,y=200)
	Bclean2 = ttk.Button(BTclean2,text='พิมพ์',style='my.TButton',command=pdfPrint2)
	Bclean2.pack(ipadx=25,ipady=20)

	##################################Lable########################
	def on_enter_id(event):
		with conn:
			c.execute(""" SELECT * FROM data WHERE cid LIKE '%s' """%VarshowID2.get())
			data = c.fetchall()
			#search = c.fetchone()
			#search = c.fetchmany()
			#print(data)
		if data != ():
			VarshowNameTH2.set(data[0][6])
			VarshowDate2.set(data[0][7])
			VarshowGender2.set(data[0][8])
			SHschool2.set(data[0][9])
			Varshowtest2.set(data[0][2])
			Varshowroomtest2.set(data[0][3])
			VarshowOrder2.set(data[0][10])
			Varshowroom2.set(data[0][4])
			CheckVar2.set(data[0][11])

		else :
			messagebox.showwarning('แจ้งเตือน','ไม่มีข้อมูลในระบบ',parent=tebmain)


	Fid2 = Label(tebManage,text='เลขประจำตัวประชาชน',font=('Angsana New',20,'bold'))
	Fid2.place(x=20,y=80)
	Fid22 = Label(tebManage,text='การค้นหาข้อมูล\nให้พิมพ์เลขบัตรประชาชน\nและกด"Enter"',font=('Angsana New',20,'bold'))
	Fid22.place(x=850,y=80)
	VarshowID2 = StringVar()
	Nid2 = ttk.Entry(tebManage,textvariable=  VarshowID2,font=('Angsana New',18,'bold'),width=60)
	Nid2.place(x=300,y=80)
	Nid2.focus()
	Nid2.bind('<Return>', on_enter_id)

	NnameTH2 = Label(tebManage,text='ชื่อ - นามสกุล',font=('Angsana New',20,'bold'),)
	NnameTH2.place(x=20,y=130)
	VarshowNameTH2 = StringVar()
	SHnameTH2 = ttk.Entry(tebManage,textvariable=  VarshowNameTH2,font=('Angsana New',18,'bold'),width=60)
	SHnameTH2.place(x=300,y=130)

	NDate2 = Label(tebManage,text='วว/ดด/ปปปป',font=('Angsana New',20,'bold'))
	NDate2.place(x=20,y=180)
	VarshowDate2 = StringVar()
	SHDate2 = ttk.Entry(tebManage,textvariable=  VarshowDate2,font=('Angsana New',18,'bold'),width=60)
	SHDate2.place(x=300,y=180)

	NGender2 = Label(tebManage,text='เพศ',font=('Angsana New',20,'bold'))
	NGender2.place(x=20,y=230)
	VarshowGender2 = StringVar()
	SHGender2 = ttk.Entry(tebManage,textvariable=  VarshowGender2,font=('Angsana New',18,'bold'),width=60)
	SHGender2.place(x=300,y=230)

	Nmail2 = Label(tebManage,text='โรงเรียน',font=('Angsana New',20,'bold'))
	Nmail2.place(x=20,y=280)
	SHschool2 = ttk.Combobox(tebManage,values=school,font=('Angsana New',18,'bold'),width=60)#,state='readonly')
	SHschool2.place(x=300,y=280)

	NArea2 = Label(tebManage,text='แผนการเรียน',font=('Angsana New',20,'bold'))
	NArea2.place(x=20,y=330)
	CheckVar2 = IntVar()
	C12 = Radiobutton(tebManage, text="แผน 1",variable=CheckVar2, value=1,font=('Angsana New',20,'bold'))
	C22 = Radiobutton(tebManage, text="แผน 2",variable=CheckVar2, value=2,font=('Angsana New',20,'bold'))
	C32 = Radiobutton(tebManage, text="แผน 3",variable=CheckVar2, value=3,font=('Angsana New',20,'bold'))
	C42 = Radiobutton(tebManage, text="แผน 4",variable=CheckVar2, value=4,font=('Angsana New',20,'bold'))
	C12.place(x=290,y=325)
	C22.place(x=390,y=325)
	C32.place(x=490,y=325)
	C42.place(x=590,y=325)

	NShowroomtest2 = Label(tebManage,text='ห้อง',font=('Angsana New',25,'bold'),fg='white',bg='red')
	NShowroomtest2.place(x=520,y=380)
	Varshowroomtest2 = StringVar()
	SHroomtest2 = ttk.Entry(tebManage,textvariable=Varshowroomtest2,font=('Angsana New',18,'bold'),width=20,state='disabled')
	SHroomtest2.place(x=600,y=385)

	room2 = Label(tebManage,text='ห้องที่',font=('Angsana New',25,'bold'),fg='white',bg='red')
	room2.place(x=220,y=380)
	Varshowroom2 = StringVar()
	SHroom2 = ttk.Entry(tebManage,textvariable=Varshowroom2,font=('Angsana New',18,'bold'),width=20,state='disabled')
	SHroom2.place(x=320,y=385)

	NShowroomtest2 = Label(tebManage,text='ลำดับที่',font=('Angsana New',25,'bold'),fg='white',bg='red')
	NShowroomtest2.place(x=220,y=450)
	VarshowOrder2 = IntVar()
	OHroomtest2 = ttk.Entry(tebManage,textvariable=VarshowOrder2,font=('Angsana New',18,'bold'),width=10,state='disabled')
	OHroomtest2.place(x=320,y=460)

	Ntest2 = Label(tebManage,text='เลขประจำตัวผู้สมัครสอบ',font=('Angsana New',25,'bold'),fg='white',bg='red')
	Ntest2.place(x=220,y=520)
	Varshowtest2 = StringVar()
	SHtest2 = ttk.Entry(tebManage,textvariable=Varshowtest2,font=('Angsana New',18,'bold'),width=30,state='disabled')
	SHtest2.place(x=500,y=528)

	###################################################################################################
#################################VIEW##########################################
##################################### สรุป 
	def sum_day():
		if SHday.get() == '' :
			with conn:
				c.execute(""" SELECT * FROM data WHERE gender LIKE '%ชา%' """)
				men = c.fetchall()
				c.execute(""" SELECT * FROM data WHERE gender LIKE '%หญิง%' """)
				women = c.fetchall()
			showMen.set(len(men))
			showWomen.set(len(women))
			showSum.set(len(men)+len(women))

		elif SHday.get() != '' :
			day =  str(SHday.get()) + "%"
			#print (day)
			with conn:
				c.execute(""" SELECT * FROM data WHERE gender LIKE %s AND dateAdd LIKE  %s """ ,('%ชาย%',day) )
				men = c.fetchall()
				c.execute(""" SELECT * FROM data WHERE gender LIKE %s AND dateAdd LIKE %s""" ,('%หญิง%',day) )
				women = c.fetchall()
			showMen.set(len(men))
			showWomen.set(len(women))
			showSum.set(len(men)+len(women))

	def exCSV ():
		c.execute(""" SELECT * FROM data""")
		data_csv = c.fetchall()
		with open("Data.csv", "w", newline='',encoding="utf-8") as csv_file:  # Python 3 version    
		#with open("out.csv", "wb") as csv_file:              # Python 2 version
		    csv_writer = csv.writer(csv_file)
		    csv_writer.writerow([i[0] for i in c.description]) # write headers
		    csv_writer.writerows(data_csv)
		messagebox.showinfo('แจ้งเตือน','Export CSV file เสร็จแล้ว',parent=tebmain)


	guilogo8 = Label(tebSum,image=logo)
	guilogo8.place(x=300,y=5)
	Label(tebSum,text='ระบบรับสมัครนักเรียนโรงเรียนพรหมานุสรณ์จังหวัดเพชรบุรี',font=('Angsana New',20,'bold')).place(x=350,y=20)
	
	Label(tebSum,text='Ex. 15/03/2019 (ค.ศ.)',font=('Angsana New',16)).place(x=650,y=90)
	Label(tebSum,text='วว/ดด/ปปปป',font=('Angsana New',20,'bold')).place(x=340,y=90)
	SHday = ttk.Entry(tebSum,font=('Angsana New',18,'bold'),width=20)
	SHday.place(x=470,y=90)

	BTre_day = Frame(tebSum,width=20) #set location
	BTre_day.place(x=473,y=140) 
	Bre_day = ttk.Button(BTre_day,text='Enter & Reresh',style='my.TButton',command= sum_day)
	Bre_day.pack(ipadx=30,ipady=8)

	Label(tebSum,text='Export CSV file',font=('Angsana New',20,'bold')).place(x=483,y=460)
	BTexcsv= Frame(tebSum,width=20) #set location
	BTexcsv.place(x=475,y=510) 
	Bex_csv = ttk.Button(BTexcsv ,text='Export to CSV',style='my.TButton',command=exCSV )
	Bex_csv.pack(ipadx=30,ipady=8)

	Label(tebSum,text='ชาย ',font=('Angsana New',22,'bold')).place(x=340,y=250)
	Label(tebSum,text='หญิง ',font=('Angsana New',22,'bold')).place(x=340,y=300)
	Label(tebSum,text='รวม ',font=('Angsana New',22,'bold'),fg='white',bg='red').place(x=340,y=350)

	Label(tebSum,text='คน ',font=('Angsana New',22,'bold')).place(x=680,y=250)
	Label(tebSum,text='คน ',font=('Angsana New',22,'bold')).place(x=680,y=300)
	Label(tebSum,text='คน ',font=('Angsana New',22,'bold'),fg='white',bg='red').place(x=680,y=350)

	with conn:
		c.execute(""" SELECT * FROM data WHERE gender LIKE '%ชาย%' """)
		men = c.fetchall()
		c.execute(""" SELECT * FROM data WHERE gender LIKE '%หญิง%' """)
		women = c.fetchall()

	showMen = IntVar()
	Label(tebSum,textvariable=showMen,font=('Angsana New',22)).place(x=510,y=250)
	showMen.set(len(men))
	showWomen = IntVar()
	Label(tebSum,textvariable=showWomen,font=('Angsana New',22,)).place(x=510,y=300)
	showWomen.set(len(women))
	showSum = IntVar()
	Label(tebSum,textvariable=showSum,font=('Angsana New',22,)).place(x=510,y=350)
	showSum.set(len(men)+len(women))



	guilogo4 = Label(tebViwe,image=logo)
	guilogo4.place(x=300,y=5)
	Label(tebViwe,text='ระบบรับสมัครนักเรียนโรงเรียนพรหมานุสรณ์จังหวัดเพชรบุรี',font=('Angsana New',20,'bold')).place(x=350,y=20)
	with conn:
			c.execute(""" SELECT room FROM dataRoom  """)
			Find = c.fetchall()
			Find =sorted(Find)
	SHFind = ttk.Combobox(tebViwe,values=Find,font=('Angsana New',18,'bold'),width=20)#,state='readonly')
	SHFind.place(x=470,y=80)

	def on_enter_find():
		#print(SHFind.get())
		for i in roomview1.get_children():
			roomview1.delete(i)
		with conn:
			c.execute(""" SELECT * FROM data WHERE room LIKE %s """ %SHFind.get()) 
			data = c.fetchall()
		for i in data :
			dataView=[i[10],i[2],i[6],i[9],i[11]]
			roomview1.insert("",'end',text="",values= dataView)

	Label(tebViwe,text='ห้องสอบที่',font=('Angsana New',20,'bold'),fg='white',bg='red').place(x=360,y=80)

	BTview= Frame(tebViwe,height=8,width=40) #set location
	BTview.place(x=680,y=73)
	bts = ttk.Button(BTview,text='ค้นหา',style='my.TButton',command= on_enter_find)
	bts.pack(ipadx=10,ipady=5)

	s.configure("Treeview.Heading", font=('Angsana New',18,'bold'))
	s.configure("Treeview", font=('Angsana New',18), rowheight=30)
	headerList1 = ['ลำดับ','เลขประจำตัวผู้สมัครสอบ','ชื่อ-สกุล','โรงเรียน',"แผนการเรียน"]
	roomview1=ttk.Treeview(tebViwe,height=13,columns=headerList1,show='headings')
	roomview1.grid(row=0,column=0,sticky='w',padx=50,pady=150,ipadx=70)


	roomview1.heading(headerList1[0].title(),text=headerList1[0].title())
	roomview1.heading(headerList1[1].title(),text=headerList1[1].title())
	roomview1.heading(headerList1[2].title(),text=headerList1[2].title())
	roomview1.heading(headerList1[3].title(),text=headerList1[3].title())
	roomview1.heading(headerList1[4].title(),text=headerList1[4].title())


	roomview1.column('ลำดับ', width=30)
	roomview1.column('เลขประจำตัวผู้สมัครสอบ', width=180)
	roomview1.column('ชื่อ-สกุล', width=280)
	roomview1.column('โรงเรียน', width=280)
	roomview1.column('แผนการเรียน', width=80)
def login():
    global c,conn
    try:
    	print(str(ipEntry.get()))
    	#conn = m.connect(host='178.128.126.68', user='admin', passwd='promma113',db='test')
    	conn = m.connect(host=str(ipEntry.get()), user=str(userEntry.get()), passwd=str(passEntry.get()),db=str(dataEntry.get()))
    	c= conn.cursor()
    	'''
    	c.execute("SELECT * FROM city ")
    	data = c.fetchone()
    	print (data)'''
    except m.Error:
    	messagebox.showwarning('แจ้งเตือน','ติดต่อฐานข้อมูลผิดพลาด')
    else:
    	main_pro()

Label(gui,text='IP Server',font=('Angsana New',22,'bold')).place(x=330,y=160)
Label(gui,text='User',font=('Angsana New',22,'bold')).place(x=330,y=210)
Label(gui,text='Password',font=('Angsana New',22,'bold')).place(x=330,y=260)
Label(gui,text='Database',font=('Angsana New',22,'bold')).place(x=330,y=310)

ipEntry = ttk.Entry(gui,font=('Angsana New',18,),width=30)
ipEntry.place(x=450,y=160)
userEntry = ttk.Entry(gui,font=('Angsana New',18,),width=30)
userEntry.place(x=450,y=210)
passEntry = ttk.Entry(gui,font=('Angsana New',18,),show='*',width=30)
passEntry.place(x=450,y=260)
dataEntry = ttk.Entry(gui,font=('Angsana New',18,),width=30)
dataEntry.place(x=450,y=310)

BTlogin = Frame(gui,height=20,width=40) #set location
BTlogin.place(x=420,y=380)
Blogin = ttk.Button(BTlogin,text='เข้าสู่ระบบ',style='my.TButton',command=login)
Blogin.pack(ipadx=50,ipady=15)


'''
if c:
    c.close()
'''
gui.mainloop()