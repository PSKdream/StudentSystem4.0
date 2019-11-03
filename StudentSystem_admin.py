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
gui.title("Student Recruitment Admin")
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
readID=()
tID = None
readCount = None
listRoom = None
search_Room =[]
InListRoom = None
search_RoomData=None
count = 0
cr = 1
readCounttest = None
scarchschool = None
u = []
trueID = []
def resetnum():
	with conn:
		c.execute(""" SELECT * FROM data""")
		tID = c.fetchall()
	with conn:
		c.execute(""" SELECT * FROM FIDtest""")
		u = c.fetchall()
		#print(u)
	global trueID
	if(u != ()):
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
def insert_FIDtest(a1):
	with conn:
		c.execute("""DELETE FROM FIDtest """)
		conn.commit()
	with conn:
		c.execute(""" INSERT INTO FIDtest VALUES (%s,%s)""",(None,a1))
		conn.commit()
	resetnum()
def insert_IDtest(a1):
	with conn:
		c.execute("""DELETE FROM IDtest """)
		conn.commit()
	with conn:
		c.execute(""" INSERT INTO IDtest VALUES (%s,%s)""",(None,a1))
		conn.commit()
	resetnum()
def main_pro():
	c.execute(""" CREATE TABLE IF NOT EXISTS data(
			ID INTEGER PRIMARY KEY AUTO_INCREMENT,dateAdd TEXT,idTest INTEGER,room INTEGER,No_Room INTEGER,cid TEXT ,nameTH TEXT ,date1 TEXT,gender TEXT,school TEXT,oorder INTEGER)""") 
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
#----------------------file--------------------------
	'''
	mainmenu =Menu(tebmain)
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
	tebEdit=Frame(Teb)
	tebRoom=Frame(Teb)
	Teb.add(tebViwe,text="View")
	#Teb.add(tebdata,text="จัดการ Database")
	#Teb.add(tebProfile,text="กรอกข้อมูล")
	Teb.add(tebManage,text='แก้ไขข้อมูล')
	Teb.add(tebEdit,text='ตั้งค่าเลขประจำตัวผู้สมัครสอบ')
	Teb.add(tebRoom,text='ตั้งค่าห้องสอบ')
	Teb.add(tebSum,text='สรุป')
	Teb.pack(fill=BOTH,expand=1) #fill X ,V BOTH	
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
			 or VarshowDate2.get()==''or VarshowGender2.get()==''  ): #or CheckVar2.get() == 0):
				messagebox.showwarning('แจ้งเตือน','กรุณากรอกข้อมูลให้ครับถ้วน',parent=tebmain)

		else:
			with conn:
				c.execute("""UPDATE data set nameTH=%s,date1=%s,gender=%s,school=%s WHERE cid = %s """
					,(VarshowNameTH2.get(),VarshowDate2.get(),VarshowGender2.get()
					,SHschool2.get(),VarshowID2.get()))
				'''if(CheckVar2.get()==1) :
					c.execute("""UPDATE data set Area=%s WHERE cid = %s """,("ในเขต",VarshowID2.get()))
				if(CheckVar2.get()==2) :
					c.execute("""UPDATE data set Area=%s WHERE cid = %s """,("นอกเขต",VarshowID2.get()))
				if(CheckVar2.get()==3) :
					c.execute("""UPDATE data set Area=%s WHERE cid = %s """,("ต่างจังหวัด",VarshowID2.get()))'''
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
	BTFind.place(x=880,y=500)
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
	school = ['วัดดอนไก่เตี้ย','อนุบาลเพชรบุรี','ปริยัติรังสรรค์','อรุณประดิษฐ','เซนต์โยเซฟ','สารสาสน์วิเทศเพชรบุรี','สุวรรณรังสฤษฎิ์วิทยาลัย','ราษฎร์วิทยา','เทศบาล 1 วัดแก่นเหล็ก','เทศบาล 3 ชุมชนวัดจันทราวาส','เทศบาล 2 วัดพระทรง']
	school.sort()
	Nmail2 = Label(tebManage,text='โรงเรียน',font=('Angsana New',20,'bold'))
	Nmail2.place(x=20,y=280)
	SHschool2 = ttk.Combobox(tebManage,values=school,font=('Angsana New',18,'bold'),width=60)#,state='readonly')
	SHschool2.place(x=300,y=280)
	'''
	Varshowmail2 = StringVar()
	SHmail2 = ttk.Entry(tebManage,textvariable=  Varshowmail2,font=('Angsana New',18,),width=60)
	SHmail2.place(x=300,y=280)'''
	'''
	NArea2 = Label(tebManage,text='พื้นที่',font=('Angsana New',20,'bold'))
	NArea2.place(x=20,y=330)
	CheckVar2 = IntVar()
	C12 = Radiobutton(tebManage, text="ในเขต",variable=CheckVar2, value=1,font=('Angsana New',20,'bold'))
	C22 = Radiobutton(tebManage, text="นอกเขต",variable=CheckVar2, value=2,font=('Angsana New',20,'bold'))
	C32 = Radiobutton(tebManage, text="ต่างจังหวัด",variable=CheckVar2, value=3,font=('Angsana New',20,'bold'))
	C12.place(x=290,y=325)
	C22.place(x=500,y=325)
	C32.place(x=690,y=325)
	'''
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
#################################  VIEW  ##########################################

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
			dataView=[i[10],i[2],i[6],i[9]]
			roomview1.insert("",'end',text="",values= dataView)

	Label(tebViwe,text='ห้องสอบที่',font=('Angsana New',20,'bold'),fg='white',bg='red').place(x=360,y=80)

	BTview= Frame(tebViwe,height=8,width=40) #set location
	BTview.place(x=680,y=73)
	bts = ttk.Button(BTview,text='ค้นหา',style='my.TButton',command= on_enter_find)
	bts.pack(ipadx=10,ipady=5)

	s.configure("Treeview.Heading", font=('Angsana New',18,'bold'))
	s.configure("Treeview", font=('Angsana New',18), rowheight=30)
	headerList1 = ['ลำดับ','เลขประจำตัวผู้สมัครสอบ','ชื่อ-สกุล','โรงเรียน']
	roomview1=ttk.Treeview(tebViwe,height=13,columns=headerList1,show='headings')
	roomview1.grid(row=0,column=0,sticky='w',padx=50,pady=150,ipadx=70)


	roomview1.heading(headerList1[0].title(),text=headerList1[0].title())
	roomview1.heading(headerList1[1].title(),text=headerList1[1].title())
	roomview1.heading(headerList1[2].title(),text=headerList1[2].title())
	roomview1.heading(headerList1[3].title(),text=headerList1[3].title())


	roomview1.column('ลำดับ', width=30)
	roomview1.column('เลขประจำตัวผู้สมัครสอบ', width=200)
	roomview1.column('ชื่อ-สกุล', width=280)
	roomview1.column('โรงเรียน', width=320)
###############################  set room
	def insert_data_room (c1):
		with conn:
			c.execute(""" INSERT INTO dataRoom VALUES (%s,%s) """,(None,c1))
			conn.commit()
		#print("insert")
		search_data_room()
	def search_data_room():
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
			return [search_Room];
	def deleteRoom():
		with conn:
			aa = showlist.get()
			c.execute(""" DELETE FROM dataRoom WHERE room = %s """ %showlist.get())
			conn.commit()
			search_data_room()
			#print(search_Room)
			count = len(search_Room)
			#print(count)
			for i in roomview.get_children():
				roomview.delete(i)
			for line in search_Room:
			   roomview.insert("",'end',text="ห้องสอบที่  ",values= line)
		
		with conn:
			c.execute(""" SELECT room FROM dataRoom  """)
			Find = c.fetchall()
			#Find =sorted(Find)
			#SHFind = ttk.Combobox(tebViwe,values=Find,font=('Angsana New',18,'bold'),width=20,state='readonly')
			SHFind.config(values=Find)
		
	def addRoom():
		with conn:
			c.execute("""SELECT * FROM dataRoom WHERE room LIKE %s  """%showlist.get())
			search = c.fetchall()
			#print(search)
		if(search == ()):
			insert_data_room(showlist.get())  #save sql
			#print(search_Room)
			count = len(search_Room)
			#print(count)
			#listRoom.append(showlist.get())
			list_get = showlist.get()
			search_data_room()
			for i in roomview.get_children():
				roomview.delete(i)
			for line in search_Room:
				roomview.insert("",'end',text="ห้องสอบที่  ",values= line)
							#print(listRoom)
							#showlist = ''
							#SHlist.place(x=450,y=220)
			with conn:
				c.execute(""" SELECT room FROM dataRoom  """)
				Find = c.fetchall()
				#Find =sorted(Find)
				#SHFind = ttk.Combobox(tebViwe,values=Find,font=('Angsana New',18,'bold'),width=20,state='readonly')
				#SHFind.place(x=470,y=80)
				SHFind.config(values=Find)
		else: 
			messagebox.showwarning('แจ้งเตือน','มีข้อมูลในระบบแล้ว',parent=tebmain)
			
			
	def insert_Count_test(c1):
		with conn:
			c.execute("""DELETE FROM Counttest """)
			conn.commit()
		with conn:
			c.execute(""" INSERT INTO Counttest VALUES (%s,%s)""",(None,c1))
			conn.commit()
	def save_Count():
				#print(showcount.get())
		if (showcount.get()=="" ):
			messagebox.showwarning('แจ้งเตือน','กรุณากรอกข้อมูลให้ครับถ้วน',parent=tebmain)
		else:
			insert_Count_test(showcount.get())
			with conn:
				c.execute(""" SELECT * FROM Counttest""")
				readCount = c.fetchall()
				if(readCount != ()):
					readCounttest = int(readCount[0][1])
			text = str(readCounttest) + " คน"
			Label(tebRoom,text=text,font=('Angsana New',20,'bold')).place(x=845,y=240)
			#messagebox.showinfo('แจ้งเตือน','บันทึกข้อมูลเรียบร้อย')
	def on_click_room(event):
		addRoom()
	def on_enter_no(event):
		save_Count()

	guilogo3 = Label(tebRoom,image=logo)
	guilogo3.place(x=300,y=5)
	Label(tebRoom,text='ระบบรับสมัครนักเรียนโรงเรียนพรหมานุสรณ์จังหวัดเพชรบุรี',font=('Angsana New',20,'bold')).place(x=350,y=20)


	showlist = StringVar()
	#showAddress.set(InListRoom)
	SHlist = ttk.Entry(tebRoom,textvariable=showlist,font=('Angsana New',18,'bold'),width=20)
	SHlist.place(x=450,y=220)
	SHlist.focus()
	SHlist.bind('<Return>', on_click_room)

	'''
	Nlist = Frame(tebRoom)
	Nlist.place(x=100,y=80)
	mylist = Listbox(Nlist,font=('Angsana New',18))
	search_data_room()
	#print(search_Room)
	#count = len(search_Room)
	#print(count)

	for line in search_Room:
	   mylist.insert(END, "This is room number " + str(line))
	mylist.pack(ipadx=70,ipady=80)'''

	headerList = ['ห้องสอบ']
	roomview=ttk.Treeview(tebRoom,height=15,columns=headerList,show='headings')
	roomview.grid(row=2,column=0,sticky='w',padx=40,pady=80,ipadx=80)
	search_data_room()
	for line in search_Room:	
		roomview.insert("",'end',text="ห้องสอบที่  ",values= line)

	#----add head
	roomview.heading(headerList[0].title(),text=headerList[0].title())


	BTsaveRoom = Frame(tebRoom,height=20,width=40) #set location
	BTsaveRoom.place(x=465,y=300)
	BsaveRoom = ttk.Button(BTsaveRoom,text='บันทึกข้อมูล',style='my.TButton',command=addRoom)
	BsaveRoom.pack(ipadx=25,ipady=20)

	BTdelete = Frame(tebRoom,height=20,width=40) #set location
	BTdelete.place(x=465,y=400)
	Bdelete = ttk.Button(BTdelete,text='ลบข้อมูล',style='my.TButton',command= deleteRoom)
	Bdelete.pack(ipadx=25,ipady=20)

	SHroom = Label(tebRoom,text='จัดการห้องสอบ',font=('Angsana New',25,'bold'),fg='white',bg='red')
	SHroom.place(x=455,y=140)

	Couttest = Label(tebRoom,text='จำนวนผู้เข้าสอบ\nแต่ละห้อง',font=('Angsana New',20,'bold'),fg='white',bg='red')
	Couttest.place(x=800,y=140)
	readCounttest1=''
	with conn:
			c.execute(""" SELECT * FROM Counttest""")
			readCount = c.fetchall()
			if(readCount != ()):
				readCounttest1 = int(readCount[0][1])
	
	text = str(readCounttest1) + " คน"
	showcount = StringVar()
	showcount.set(readCounttest1)
	SHcount = ttk.Entry(tebRoom,textvariable=showcount,font=('Angsana New',18,'bold'),width=20)
	SHcount.place(x=778,y=300)
	Label(tebRoom,text=text,font=('Angsana New',20,'bold')).place(x=845,y=240)
	BTsave_count = Frame(tebRoom,height=20,width=40) #set location
	BTsave_count.place(x=795,y=360) #795
	Bsave_count = ttk.Button(BTsave_count,text='บันทึกข้อมูล',style='my.TButton',command= save_Count)
	Bsave_count.pack(ipadx=25,ipady=20)
	SHcount.focus()
	SHcount.bind('<Return>', on_enter_no) 
#################################  et test
	def insert_IDtest(a1):
		with conn:
			c.execute("""DELETE FROM IDtest """)
			conn.commit()
		with conn:
			c.execute(""" INSERT INTO IDtest VALUES (%s,%s)""",(None,a1))
			conn.commit()
		with conn:
			c.execute("""DELETE FROM FIDtest """)
			conn.commit()
		with conn:
			c.execute(""" INSERT INTO FIDtest VALUES (%s,%s)""",(None,a1))
			conn.commit()
	Label(tebEdit,text='ระบบรับสมัครนักเรียนโรงเรียนพรหมานุสรณ์จังหวัดเพชรบุรี',font=('Angsana New',20,'bold')).place(x=350,y=20)
	guilogo2 = Label(tebEdit,image=logo)
	guilogo2.place(x=300,y=5)
	def add_start_idtest():
		#print(showsetting.get())
		if (showsetting.get()=="" ):
			messagebox.showwarning('แจ้งเตือน','กรุณากรอกข้อมูลให้ครับถ้วน',parent=tebmain)
		else:
			insert_IDtest(showsetting.get())
			messagebox.showinfo('แจ้งเตือน','บันทึกข้อมูลเรียบร้อย',parent=tebmain)
			with conn:
				c.execute(""" SELECT * FROM FIDtest""")
				saveIDtest = c.fetchall()
			if(int(saveIDtest[0][1])==0):	
				insert_FIDtest(showsetting.get())
			with conn:
				c.execute(""" SELECT * FROM FIDtest""")
				saveIDtest = c.fetchall()
			#print(saveIDtest)

	def on_enter_id(event):
		add_start_idtest()
	Nsetting = Label(tebEdit,text='การตั้งค่าเลขประจำตัวผู้สมัครสอบเริ่มต้น',font=('Angsana New',25,'bold'),fg='white',bg='red')
	#Nsetting.place(x=400,y=150)
	readIDtest1 = None
	#print(readIDtest)
	with conn:
		c.execute(""" SELECT * FROM IDtest""")
		readID = c.fetchall()
		#print(readID)
		if(readID != ()):
			readIDtest1 = int(readID[0][1])
	#print(readIDtest)
	Nsetting.pack(pady=100)
	showsetting = StringVar()
	showsetting.set(readIDtest1)
	SHsetting = ttk.Entry(tebEdit,textvariable=showsetting,font=('Angsana New',18,'bold'),width=30)
	SHsetting.place(x=415,y=210)
	SHsetting.focus()
	SHsetting.bind('<Return>',on_enter_id)
	BTsave2 = Frame(tebEdit,height=20,width=40) #set location
	BTsave2.place(x=465,y=300)
	Bsave2 = ttk.Button(BTsave2,text='บันทึกข้อมูล',style='my.TButton',command=add_start_idtest)
	Bsave2.pack(ipadx=25,ipady=20)
########################### สรุป
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




def login():
    global c,conn
    try:
    	print(str(ipEntry.get()))
    	#conn = m.connect(host='178.128.126.68', user='admin', passwd='promma113',db='test')
    	conn = m.connect(host=str(ipEntry.get()), user=str(userEntry.get()), passwd=str(passEntry.get()),db=str(dataEntry.get()))
    	c= conn.cursor()
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