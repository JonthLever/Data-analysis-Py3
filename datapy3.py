import random
import xlrd
import xlsxwriter
import os

filePath="C:\\Users\\Username\\Desktop\\File\\Product Template2.xlsx"
openFile=xlrd.open_workbook(filePath);
sheet=openFile.sheet_by_name("Products in stock")
workbook = xlsxwriter.Workbook('Product Template.xlsx')
worksheet = workbook.add_worksheet()
print("Rows", sheet.nrows)
print("Cols", sheet.ncols)
costot=0
cond=0
while(cond == 0):
		"""the constants in this "if" can vary, I am trying 
		to have the user enter their final total amount and have the
	 	code create a range to arrive at an approximate amount."""
	if costot>2861205 and costot<2861211:
		cond=1;
	else:
		aux1=0
		'''if costot>2900000:
			va1=va1-10'''
		costoxp=0
		costot=0
		rx1=0
		rx2=0
		re=0
		for i in range(2221):
			i+=1
			'''print(sheet.cell_value(i,1),"   ", sheet.cell_value(i,5))'''
			#This line of code becomes irrelevant later on.
			r = random.randrange(-1000,1000);
			e = 0
			if r<0:
				aux1=r*-1
			b = sheet.cell_value(i,5)
			cant=sheet.cell_value(i,1)
			'''if aux1>0 and int(cant)>0 and int(cant)<aux1:'''
			if int(cant)>0:
				rx1=(int(cant)*-1);
				rx2=int(cant)
				rx=int(random.randrange(rx1,rx2))
				r=rx
				re = int(random.randrange(1,2))
				if re > 1:
					r=r*1
				else:
					r=r*-1
				asd1=int(cant)+r
			else:
				rzero=int(random.randrange(0,22))
				asd1=int(cant)+rzero
			'''if asd1<0:
				asd1=asd1*-1'''
			
			asd2=round(float(b), 2)
			if asd2<=0.01:
				asd1=0
				asd2=0
			qwe1=str(asd1)
			qwe2=str(asd2)
			worksheet.write(i,0,qwe1)
			worksheet.write(i,1, qwe2)
			costoxp=asd1*asd2
			costot=costot+costoxp
		print("\U0001F973","\U0001F929","\U0001F973","\U0001F973","\U0001F929","\U0001F973")
		print("\U0001F929",round(costot,2),"\U0001F929")
		print("\U0001F973","\U0001F929","\U0001F973","\U0001F973","\U0001F929","\U0001F973")
		os.system("cls") 
		

print("Congratulations! We found an approximate amount")
print("The total cost is:", round(costot,2))


workbook.close()
		



