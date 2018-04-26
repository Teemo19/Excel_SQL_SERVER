import pandas as pd 
import time
import pyodbc

Server_Name="LAPTOP-D3HE5LCI"
Database="TEST_DEV"

def connectDB():
	cnxn = pyodbc.connect("Driver={ODBC Driver 17 for SQL Server};"
                    "Server="+Server_Name+";"                                   
                    "Database="+Database+";"
                    "Trusted_Connection=yes;")
	return cnxn


def insertDB(idx,nombre,cant,precio):
	cursor = connectDB().cursor()
	cursor.execute("INSERT INTO TEST (ID,Nombre,Cantidad,Precio) Values (?,?,?,?)",int(idx),str(nombre),int(cant),int(precio))
	cursor.commit()
	cursor.close()
	connectDB().close()
	time.sleep(0.1)

def read_Excel(file,hoja):
	file=pd.read_excel(file,"Hoja"+hoja)
	return file

def ExcelFile():
	try:
		file_name=str(raw_input("Introduce el nombre del archivo excel: "))
		file_sheet=str(input("Introduce el numero de hoja: "))
		Excel_file=read_Excel(file_name,file_sheet)
	except:
		print("ARCHIVO O HOJA NO ENCONTRADO !!")
		ExcelFile()
	return Excel_file

def InsertData(file):
	for i in range(0,len(file)):
		insertDB(file['ID'][i],file['Nombre'][i],file['Cantidad'][i],file['Precio'][i])

def main():
		File=ExcelFile()
		InsertData(File)

main()
