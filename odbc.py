import pyodbc 
cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                      "Server=LAPTOP-D3HE5LCI;"
                      "Database=UNSYS_DEV;"
                      "Trusted_Connection=yes;")


cursor = cnxn.cursor()
cursor.execute("UPDATE Company SET Addr1='asd' Where Addr1='dsa' ")
cursor.commit()



