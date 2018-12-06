import xlrd 
import xlwt 
from xlwt import Workbook 
from xlutils.copy import copy 
import re
import MySQLdb

db = MySQLdb.connect(host="localhost",  # your host 
                     user="root",       # username
                     passwd="",     # password
                     db="project")   # name of the database

def totalcredits(i):
	rf = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/facwish.xls')
	sheetf = rf.sheet_by_index(0) 
	creditsalloted=sheetf.cell_value(i,6)+sheetf.cell_value(i,7)+sheetf.cell_value(i,8)+sheetf.cell_value(i,9)+sheetf.cell_value(i,10)
	wf = copy(rf) 
	sheet1 = wf.get_sheet(0)
	sheet1.write(i,11,creditsalloted)
	wf.save('C:/Users/Satyam/Desktop/Dbms/facwish.xls')
	return creditsalloted
	
def slotcheck(i,slot):
	word=slot
	xw=list();
	xw=word.split('+')
	lxw=len(xw)
	rf = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/facwish.xls') 
	sheetf = rf.sheet_by_index(0) 
	nrf = sheetf.nrows
	
	if((sheetf.cell_value(i,12))==0):
		return 1
	
	cell1=sheetf.cell_value(i,12)
	xcell1=list();
	xcell1=cell1.split('+')
	lxcell1=len(xcell1)
	
	if((sheetf.cell_value(i,13))!=0):
		cell2=sheetf.cell_value(i,13)
		xcell2=list();
		xcell2=cell2.split('+')
		lxcell2=len(xcell2)
	
	if((sheetf.cell_value(i,14))!=0):
		cell3=sheetf.cell_value(i,14)
		xcell3=list();
		xcell3=cell3.split('+')
		lxcell3=len(xcell3)
	
	checker1=1
	checker2=1
	checker3=1
	
	if((sheetf.cell_value(i,12))!=0):
		for k in range(0,lxw):
			for l in range(0,lxcell1):
				if(xw[k]==xcell1[l]):
					checker1=0;
					break;
	
	if((sheetf.cell_value(i,13))!=0):
		for k in range(0,lxw):
			for l in range(0,lxcell2):
				if(xw[k]==xcell2[l]):
					checker2=0;
					break;
	
	if((sheetf.cell_value(i,14))!=0):
		for k in range(0,lxw):
			for l in range(0,lxcell3):
				if(xw[k]==xcell3[l]):
					checker3=0;
					break;

	if((checker1*checker2*checker3)==1):
		if((sheetf.cell_value(i,13))==0):
			return 2
		elif((sheetf.cell_value(i,14))==0):
			return 3
	else:
		return 0
	
def slotalloted(i,slot):
	word=slot
	rf = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/facwish.xls') 
	sheetf = rf.sheet_by_index(0) 
	nrf = sheetf.nrows
	
	if((sheetf.cell_value(i,12))==0):
		rf = xlrd.open_workbook("C:/Users/Satyam/Desktop/Dbms/facwish.xls") 
		wf = copy(rf) 
		sheet1 = wf.get_sheet(0)
		sheet1.write(i,12,word)
		wf.save('C:/Users/Satyam/Desktop/Dbms/facwish.xls')
	elif((sheetf.cell_value(i,13))==0):
		rf = xlrd.open_workbook("C:/Users/Satyam/Desktop/Dbms/facwish.xls") 
		wf = copy(rf) 
		sheet1 = wf.get_sheet(0)
		sheet1.write(i,13,word)
		wf.save('C:/Users/Satyam/Desktop/Dbms/facwish.xls')	
	elif((sheetf.cell_value(i,14))==0):
		rf = xlrd.open_workbook("C:/Users/Satyam/Desktop/Dbms/facwish.xls") 
		wf = copy(rf) 
		sheet1 = wf.get_sheet(0)
		sheet1.write(i,14,word)
		wf.save('C:/Users/Satyam/Desktop/Dbms/facwish.xls')

def slotcompatible(t,sub,col):
	rf = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/facwish.xls') 
	sheetf = rf.sheet_by_index(0) 
	if(sheetf.cell_value(t,col)==0):
		rc = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/batcou.xls')
		sheetc = rc.sheet_by_index(0) 
		nrc = sheetc.nrows
		flag=0
		for i in range(0,nrc):
			if(sub==sheetc.cell_value(i,0)):
				x=sheetc.cell_value(i,1)
				slot=sheetc.cell_value(i,3)
				cr=sheetc.cell_value(i,2)
				if(x>0):
					y=slotcheck(t,slot)
					if(y>0):
						slotalloted(t,slot)
						flag=1
						x=x-1
						rc = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/batcou.xls')
						wc = copy(rc)
						sheet1 = wc.get_sheet(0)
						sheet1.write(i,1,x)
						wc.save('C:/Users/Satyam/Desktop/Dbms/batcou.xls')
												
						rf = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/facwish.xls') 
						wf = copy(rf)
						sheetff = wf.get_sheet(0)
						sheetff.write(t,col,cr)
						wf.save('C:/Users/Satyam/Desktop/Dbms/facwish.xls')
						return 1
			else:
				continue
		if(flag==1):
			return 1
		else:
			return 0
	else:
		return 0
		
def allocation():
	rf = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/facwish.xls') 
	rc = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/batcou.xls')

	sheetf = rf.sheet_by_index(0) 
	sheetc = rc.sheet_by_index(0)

	nrf = sheetf.nrows

	for alex in range(0,3):
		for i in range(0,nrf):	
				rf = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/facwish.xls') 
				rc = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/batcou.xls')
				sheetf = rf.sheet_by_index(0) 
				sheetc = rc.sheet_by_index(0) 
				subfw1=sheetf.cell_value(i,1)
				subfw2=sheetf.cell_value(i,2)
				subfw3=sheetf.cell_value(i,3)
				subfw4=sheetf.cell_value(i,4)
				subfw5=sheetf.cell_value(i,5)
						
				if(totalcredits(i)<8):	
					if(slotcompatible(i,subfw1,6)==1):
						abc='1'
					elif(slotcompatible(i,subfw2,7)==1):
						abc='2'
					elif(slotcompatible(i,subfw3,8)==1):
						abc='3'
					elif(slotcompatible(i,subfw4,9)==1):
						abc='4'
					elif(slotcompatible(i,subfw5,10)==1):
						abc='5'
					else:
						print('END')
					temp=totalcredits(i)

	for alex in range(0,2):
		for i in range(0,nrf):	
				rf = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/facwish.xls') 
				rc = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/batcou.xls')
				sheetf = rf.sheet_by_index(0) 
				sheetc = rc.sheet_by_index(0) 
				subfw1=sheetf.cell_value(i,1)
				subfw2=sheetf.cell_value(i,2)
				subfw3=sheetf.cell_value(i,3)
				subfw4=sheetf.cell_value(i,4)
				subfw5=sheetf.cell_value(i,5)
						
				if(totalcredits(i)==8):	
					if(slotcompatible(i,subfw1,6)==1):
						abc='1'
					elif(slotcompatible(i,subfw2,7)==1):
						abc='2'
					elif(slotcompatible(i,subfw3,8)==1):
						abc='3'
					elif(slotcompatible(i,subfw4,9)==1):
						abc='4'
					elif(slotcompatible(i,subfw5,10)==1):
						abc='5'
					else:
						print('END')
					temp=totalcredits(i)
					
def extractcourse():
	rcourse = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/course.xls')
	rsheetcourse = rcourse.sheet_by_index(0) 
	nrcourse = rsheetcourse.nrows
	
	rc = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/batcou.xls') 
	sheetc = rc.sheet_by_index(0)
	wc = copy(rc)
	
	for i in range(1,nrcourse):
		ccode=rsheetcourse.cell_value(i,0)
		nbatch=rsheetcourse.cell_value(i,5)
		kredits=rsheetcourse.cell_value(i,2)
		slots=rsheetcourse.cell_value(i,4)
		
		sheetcw = wc.get_sheet(0)
		sheetcw.write(i-1,0,ccode)
		sheetcw.write(i-1,1,nbatch)
		sheetcw.write(i-1,2,kredits)
		sheetcw.write(i-1,3,slots)
		sheetcw.write(i-1,6,nbatch)
		wc.save('C:/Users/Satyam/Desktop/Dbms/batcou.xls')
	
	print('Course Data Extracted')
	
def extractfaculty():
	rfaculty = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/wishlist.xls')
	rsheetfaculty = rfaculty.sheet_by_index(0) 
	nrfaculty = rsheetfaculty.nrows
	
	rf = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/facwish.xls') 
	sheetc = rf.sheet_by_index(0)
	wf = copy(rf)
	
	for i in range(1,nrfaculty):
		fid=rsheetfaculty.cell_value(i,0)
		sub1=rsheetfaculty.cell_value(i,2)
		sub2=rsheetfaculty.cell_value(i,3)
		sub3=rsheetfaculty.cell_value(i,4)
		sub4=rsheetfaculty.cell_value(i,5)
		sub5=rsheetfaculty.cell_value(i,6)
		
		sheetfw = wf.get_sheet(0)
		sheetfw.write(i-1,0,fid)
		sheetfw.write(i-1,1,sub1)
		sheetfw.write(i-1,2,sub2)
		sheetfw.write(i-1,3,sub3)
		sheetfw.write(i-1,4,sub4)
		sheetfw.write(i-1,5,sub5)
		sheetfw.write(i-1,17,0)
		for j in range(6,15):
			sheetfw.write(i-1,j,0)
			
		wf.save('C:/Users/Satyam/Desktop/Dbms/facwish.xls')
	
	print('Faculty Data Extracted')
	
def fbc():
	rf = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/facwish.xls') 
	sheetrf = rf.sheet_by_index(0) 
	nrf = sheetrf.nrows
	
	rfaculty = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/faculty.xls') 
	sheetrfaculty = rfaculty.sheet_by_index(0)
	
	rff = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/fbc.xls') 
	sheetwf = rff.sheet_by_index(0)
	wff = copy(rff)
	sheetwf = wff.get_sheet(0)
	sheetwf.write(0,0,'Faculty ID')
	sheetwf.write(0,1,'Name')
	sheetwf.write(0,2,'Designation')
	sheetwf.write(0,3,'Sub 1')
	sheetwf.write(0,4,'Credits')
	sheetwf.write(0,5,'Sub 2')
	sheetwf.write(0,6,'Credits')
	sheetwf.write(0,7,'Sub 3')
	sheetwf.write(0,8,'Credits')
	sheetwf.write(0,9,'Total Credits')
	
	for i in range(0,nrf):
		fid=sheetrfaculty.cell_value(i+1,0)
		fname=sheetrfaculty.cell_value(i+1,1)
		designation=sheetrfaculty.cell_value(i+1,2)
		
		sheetwf = wff.get_sheet(0)
		sheetwf.write(i+1,0,fid)
		sheetwf.write(i+1,1,fname)
		sheetwf.write(i+1,2,designation)
		
		sub1='NULL'
		sub2='NULL'
		sub3='NULL'
		sub4='NULL'
		sub5='NULL'
		if(sheetrf.cell_value(i,6)!=0):
			sub1=sheetrf.cell_value(i,1)
			cr1=sheetrf.cell_value(i,6)
		if(sheetrf.cell_value(i,7)!=0):
			sub2=sheetrf.cell_value(i,2)
			cr2=sheetrf.cell_value(i,7)
		if(sheetrf.cell_value(i,8)!=0):
			sub3=sheetrf.cell_value(i,3)
			cr3=sheetrf.cell_value(i,8)
		if(sheetrf.cell_value(i,9)!=0):
			sub4=sheetrf.cell_value(i,4)
			cr4=sheetrf.cell_value(i,9)
		if(sheetrf.cell_value(i,10)!=0):
			sub5=sheetrf.cell_value(i,5)
			cr5=sheetrf.cell_value(i,10)
		
		for n in range(3,9):
			sheetwf.write(i+1,n,0)
		
		for m in range(0,3):
			if(sub1!='NULL'):
				sheetwf.write(i+1,3+(2*m),sub1)	
				sheetwf.write(i+1,4+(2*m),cr1)
				sub1='NULL'	
			elif(sub2!='NULL'):
				sheetwf.write(i+1,3+(2*m),sub2)	
				sheetwf.write(i+1,4+(2*m),cr2)
				sub2='NULL'
			elif(sub3!='NULL'):
				sheetwf.write(i+1,3+(2*m),sub3)	
				sheetwf.write(i+1,4+(2*m),cr3)
				sub3='NULL'
			elif(sub4!='NULL'):
				sheetwf.write(i+1,3+(2*m),sub4)	
				sheetwf.write(i+1,4+(2*m),cr4)
				sub4='NULL'
			elif(sub5!='NULL'):
				sheetwf.write(i+1,3+(2*m),sub5)	
				sheetwf.write(i+1,4+(2*m),cr5)
				sub5='NULL'
		
		tc=sheetrf.cell_value(i,11)
		sheetwf.write(i+1,9,tc)
		
		wff.save('C:/Users/Satyam/Desktop/Dbms/fbc.xls')
	
	print('FBC EXCEL/DATA GENERATIED')

#extractcourse()
#extractfaculty()
#allocation()
#fbc()

def courserallocation():
	rfbc = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/fbc.xls') 
	sheetrfbc = rfbc.sheet_by_index(0) 
	nrfbc = sheetrfbc.nrows
	
	rc = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/course.xls') 
	sheetc = rc.sheet_by_index(0) 
	num = sheetc.nrows
	
	rf = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/facwish.xls') 
	sheetrf = rf.sheet_by_index(0)
	
	rca = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/courseallocation.xls') 
	sheetwca = rca.sheet_by_index(0)
	wca = copy(rca)
	sheetwca = wca.get_sheet(0)
	
	sheetwca.write(0,0,'Course Code')
	sheetwca.write(0,1,'Course Name')
	sheetwca.write(0,2,'Credits')
	sheetwca.write(0,3,'Program')
	sheetwca.write(0,4,'Slot')
	sheetwca.write(0,5,'Venue')
	sheetwca.write(0,6,'Faculty Id')
	sheetwca.write(0,7,'Faculty Name')
	slot=300
	lin=1
	for i in range(1,nrfbc):
		sheetwca = wca.get_sheet(0)
		
		if((sheetrfbc.cell_value(i,3))!=0):
			sheetwca.write(lin,0,sheetrfbc.cell_value(i,3))
			for k in range(1,num):
				if((sheetrfbc.cell_value(i,3))==(sheetc.cell_value(k,0))):
					cname=sheetc.cell_value(k,1)
					program=sheetc.cell_value(k,3)
					break;
			sheetwca.write(lin,1,cname)
			sheetwca.write(lin,2,sheetrfbc.cell_value(i,4))
			sheetwca.write(lin,3,program)
			sheetwca.write(lin,4,sheetrf.cell_value(i-1,12))
			sheetwca.write(lin,5,slot)
			sheetwca.write(lin,6,sheetrfbc.cell_value(i,0))
			sheetwca.write(lin,7,sheetrfbc.cell_value(i,1))
			slot=slot+1
			lin=lin+1
	
		if((sheetrfbc.cell_value(i,5))!=0):
			sheetwca.write(lin,0,sheetrfbc.cell_value(i,5))
			for k in range(1,num):
				if((sheetrfbc.cell_value(i,5))==(sheetc.cell_value(k,0))):
					cname=sheetc.cell_value(k,1)
					program=sheetc.cell_value(k,3)
					break;
			sheetwca.write(lin,1,cname)
			sheetwca.write(lin,2,sheetrfbc.cell_value(i,6))
			sheetwca.write(lin,3,program)
			sheetwca.write(lin,4,sheetrf.cell_value(i-1,13))
			sheetwca.write(lin,5,slot)
			sheetwca.write(lin,6,sheetrfbc.cell_value(i,0))
			sheetwca.write(lin,7,sheetrfbc.cell_value(i,1))
			slot=slot+1
			lin=lin+1
		
		if((sheetrfbc.cell_value(i,7))!=0):
			sheetwca.write(lin,0,sheetrfbc.cell_value(i,7))
			for k in range(1,num):
				if((sheetrfbc.cell_value(i,7))==(sheetc.cell_value(k,0))):
					cname=sheetc.cell_value(k,1)
					program=sheetc.cell_value(k,3)
					break;
			sheetwca.write(lin,1,cname)
			sheetwca.write(lin,2,sheetrfbc.cell_value(i,8))
			sheetwca.write(lin,3,program)
			sheetwca.write(lin,4,sheetrf.cell_value(i-1,14))
			sheetwca.write(lin,5,slot)
			sheetwca.write(lin,6,sheetrfbc.cell_value(i,0))
			sheetwca.write(lin,7,sheetrfbc.cell_value(i,1))
			slot=slot+1
			lin=lin+1
		wca.save('C:/Users/Satyam/Desktop/Dbms/courseallocation.xls')
	print('Course Allocation Donee!')

def sqlfbc():
	loc = "C:/Users/Satyam/Desktop/Dbms/fbc.xls" 
	wb = xlrd.open_workbook(loc) 
	sheet = wb.sheet_by_index(0) 
	  
	# For row 0 and column 0 
	sheet.cell_value(0, 0) 
	  
	nro=sheet.nrows-1
	nco=sheet.ncols

	fid=[];
	fname=[];
	fdesig=[];
	fsub1=[];
	ccredits1=[];
	fsub2=[];
	ccredits2=[];
	fsub3=[];
	ccredits3=[];


	for i in range(nro):
		fid.append(sheet.cell_value(i+1, 0))
		fname.append(sheet.cell_value(i+1, 1))
		fdesig.append(sheet.cell_value(i+1, 2))
		fsub1.append(sheet.cell_value(i+1, 3))
		ccredits1.append(sheet.cell_value(i+1, 4))
		fsub2.append(sheet.cell_value(i+1, 5))
		ccredits2.append(sheet.cell_value(i+1, 6))
		fsub3.append(sheet.cell_value(i+1, 7))
		ccredits3.append(sheet.cell_value(i+1, 8))

	#for i in range(nro):
		#print fid[i],fname[i],fdesig[i],fsub1[i],ccredits1[i],fsub2[i],ccredits2[i],fsub3[i],ccredits3[i]

	mycursor = db.cursor()
	#mycursor.execute("CREATE TABLE fbc(fid varchar(7),fname varchar(60),designation varchar(20),sub1 varchar(7),credits1 varchar(2),sub2 varchar(7),credits2 varchar(2),sub3 varchar(7),credits3 varchar(2)")
	sql = "INSERT INTO fbc(fid,fname,designation,sub1,credits1,sub2,credits2,sub3,credits3) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)"

	for i in range(nro):
		val = (fid[i],fname[i],fdesig[i],fsub1[i],ccredits1[i],fsub2[i],ccredits2[i],fsub3[i],ccredits3[i])
		mycursor.execute(sql, val)

	db.commit()

	print(mycursor.rowcount, "FBC RECORDED")
	print('FACULTY BASED ALLOCATION RECORDED TO SQL')
	
def sqlcourserallocation():
	loc = "C:/Users/Satyam/Desktop/Dbms/courseallocation.xls"
	wb = xlrd.open_workbook(loc) 
	sheet = wb.sheet_by_index(0) 
	  
	# For row 0 and column 0 
	sheet.cell_value(0, 0) 
	
	nro=sheet.nrows-1
	nco=sheet.ncols

	ccode=[];
	cname=[];
	ccredits=[];
	program=[];
	slot=[];
	venue=[];
	fid=[];
	fname=[];

	for i in range(nro):
		ccode.append(sheet.cell_value(i+1, 0))
		cname.append(sheet.cell_value(i+1, 1))
		ccredits.append(sheet.cell_value(i+1, 2))
		program.append(sheet.cell_value(i+1, 3))
		slot.append(sheet.cell_value(i+1, 4))
		venue.append(sheet.cell_value(i+1, 5))
		fid.append(sheet.cell_value(i+1, 6))
		fname.append(sheet.cell_value(i+1, 7))

	#for i in range(nro):
		#print ccode[i],cname[i],ccredits[i],program[i],slot[i],venue[i],fid[i],fname[i]

	mycursor = db.cursor()
	#mycursor.execute("CREATE TABLE courseallocation(course_code varchar(7),course_name varchar(60),credits varchar(2),program varchar(10),slot varchar(15),venue varchar(10),fid varchar(5),fname varchar(60)")
	sql = "INSERT INTO courseallocation(course_code,course_name,credits,program,slot,venue,fid,fname) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)"

	for i in range(nro):
		val = (ccode[i],cname[i],ccredits[i],program[i],slot[i],venue[i],fid[i],fname[i])
		mycursor.execute(sql, val)

	db.commit()
	print(mycursor.rowcount, "Record Inserted.")
	print('COURSE ALLOCATION IN SQL EXECUTED')
	
def sqlcourse():
	loc = "C:/Users/Satyam/Desktop/Dbms/course.xls"
	wb = xlrd.open_workbook(loc) 
	sheet = wb.sheet_by_index(0) 
	  
	# For row 0 and column 0 
	sheet.cell_value(0, 0) 
	
	nro=sheet.nrows-1
	nco=sheet.ncols

	ccode=[];
	cname=[];
	ccredits=[];
	program=[];
	slot=[];
	batches=[];

	for i in range(nro):
		ccode.append(sheet.cell_value(i+1, 0))
		cname.append(sheet.cell_value(i+1, 1))
		ccredits.append(sheet.cell_value(i+1, 2))
		program.append(sheet.cell_value(i+1, 3))
		slot.append(sheet.cell_value(i+1, 4))
		batches.append(sheet.cell_value(i+1, 5))

	#for i in range(nro):
		#print ccode[i],cname[i],ccredits[i],program[i],slot[i],venue[i],fid[i],fname[i]

	mycursor = db.cursor()
	#mycursor.execute("CREATE TABLE testcourse(course_code varchar(7),course_name varchar(50),credits varchar(2),program varchar(15),slot varchar(20),batches varchar(2)")
	sql = "INSERT INTO course(course_code,course_name,credits,program,slot,batches) VALUES (%s,%s,%s,%s,%s,%s)"

	for i in range(nro):
		val = (ccode[i],cname[i],ccredits[i],program[i],slot[i],batches[i])
		mycursor.execute(sql, val)

	db.commit()
	print(mycursor.rowcount, "RECORDED")
	print('\n~~~COURSE DETIALS~~~\n')

def sqlfacwish():
	loc = "C:/Users/Satyam/Desktop/Dbms/wishlist.xls" 
	wb = xlrd.open_workbook(loc) 
	sheet = wb.sheet_by_index(0) 
	  
	# For row 0 and column 0 
	sheet.cell_value(0, 0) 
	  
	nro=sheet.nrows-1
	nco=sheet.ncols

	fid=[];
	fname=[];
	fsub1=[];
	fsub2=[];
	fsub3=[];
	fsub4=[];
	fsub5=[];

	for i in range(nro):
		fid.append(sheet.cell_value(i+1, 0))
		fname.append(sheet.cell_value(i+1, 1))
		fsub1.append(sheet.cell_value(i+1, 2))
		fsub2.append(sheet.cell_value(i+1, 3))
		fsub3.append(sheet.cell_value(i+1, 4))
		fsub4.append(sheet.cell_value(i+1, 5))
		fsub5.append(sheet.cell_value(i+1, 6))

	#for i in range(nro):
		#print fid[i],fname[i],fdesig[i],fsub1[i],ccredits1[i],fsub2[i],ccredits2[i],fsub3[i],ccredits3[i]

	mycursor = db.cursor()
	#mycursor.execute("CREATE TABLE facwish(fid varchar(5),fname varchar(20),subject_1 varchar(7),subject_2 varchar(7),subject_3 varchar(7),subject_3 varchar(7),subject_5 varchar(7)")
	sql = "INSERT INTO facwish(fid,fname,subject_1,subject_2,subject_3,subject_4,subject_5) VALUES (%s,%s,%s,%s,%s,%s,%s)"

	for i in range(nro):
		val = (fid[i],fname[i],fsub1[i],fsub2[i],fsub3[i],fsub4[i],fsub5[i])
		mycursor.execute(sql, val)

	db.commit()

	print(mycursor.rowcount, "RECORDED")
	print('~~~FACULTY WISHLIST~~~')
	
choice=1
while choice>0:
	print('1. EXTRACT COURSE')
	print('2. EXTRACT FACULTY')
	print('3. PUSH SQL COURSE DETAILS')
	print('4. PUSH SQL FACULTY WISHLIST')
	print('5. RUN ALLOCATION ALGORITHM')
	print('6. EXECUTE FBC (FACULTY BASED CREDITS)')
	print('7. EXECUTE COURSE ALLOCATION')
	print('8. PUSH SQL FBC')
	print('9. PUSH SQL COURSE ALLOCATION')
	
	choice=int(input())
	if(choice==1):
		extractcourse()
	elif(choice==2):
		extractfaculty()
	elif(choice==3):
		sqlcourse()
	elif(choice==4):
		sqlfacwish()
	elif(choice==5):
		allocation()
	elif(choice==6):
		fbc()		
	elif(choice==7):
		courserallocation()
	elif(choice==8):
		sqlfbc()
	elif(choice==9):
		sqlcourserallocation()
	else:
		print('Enter Valid Input!')

#def alotvenue():
	##rfbc = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/fbc.xls') 
	##sheetrfbc = rfbc.sheet_by_index(0) 
	##nrfbc = sheetrfbc.nrows
	
	##rc = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/course.xls') 
	##sheetc = rc.sheet_by_index(0) 
	##num = sheetc.nrows
	
	##rf = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/facwish.xls') 
	##sheetrf = rf.sheet_by_index(0)
	
	##rca = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/courseallocation.xls') 
	##sheetwca = rca.sheet_by_index(0)
	##wca = copy(rca)
	##sheetwca = wca.get_sheet(0)
	
	################################################
	#rca = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/courseallocation.xls') 
	#sheetrca = rca.sheet_by_index(0) 
	#nrca = sheetrca.nrows
	
	#rv = xlrd.open_workbook('C:/Users/Satyam/Desktop/Dbms/venue.xls') 
	#sheetrv = rv.sheet_by_index(0)
	#wvv = copy(rv)
	#sheetwv = wvv.get_sheet(0)
	##sheetwv.write(1,1,0)
	##wvv.save('C:/Users/Satyam/Desktop/Dbms/venue.xls')

	#for i in range(1,nrca):
		#slot=sheetrca.cell_value(i,4)
		#word=slot
		#wv=list();
		#wv=word.split('+')
		#lwv=len(wv)
		#k=1
		#m=1
		#n=1
		#flag=1
		#while k>0:
			#if(sheetrv.cell_value(m,n)!=0):
				#cellw=sheetrv.cell_value(m,n)
				#wordw=list();
				#wordw=cellw.split('+')
				#lww=len(wordw)				
				#for big in range(lwv):
					#for small in range(lww):
						#if(wv[big]==wordw[small]):
							#flag=99
			#else:
				#sheetwv.write(m,n,slot)
				#sheetwv.write(m,n+1,0)
				#wvv.save('C:/Users/Satyam/Desktop/Dbms/venue.xls')
				#k=0
				
			#if(flag==99):
				#m=m+1;
			#else:
				#n=n+1
			#k=k+1
		#print 'END'
				
#alotvenue()
