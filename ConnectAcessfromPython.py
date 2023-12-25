import csv
import pyodbc
import pandas as pd

# Connect to the Access database
# Using a DSN
#cnxn = pyodbc.connect('DSN=odbc_datasource_name;UID=db_user_id;PWD=db_password')
connStr = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    r"DBQ=C:\DataProject\MYDB1.accdb;"
    )

#connStr = (r"DSN=CUSTOMERS;")

#cnxn = pyodbc.connect(connStr)

conn = pyodbc.connect(connStr)

# Create a cursor object
cursor = conn.cursor()

# Execute a query to fetch all records from the customer table
query='SELECT * FROM customers'
cursor.execute(query)

# Fetch all records
records = cursor.fetchall()
print(records)
print('TYPE OF RECORDS -> ',type(records))

print('ROWS in RECORDS -> ',len(records))
for row in records:
    print(row) 

#Fetch data in a dataframe
print('dataframe')
df=pd.read_sql(query,conn)
print(df)

# Define the path and filename for the CSV file
csv_file = 'C:\DataProject\output.csv'

# Write the records to the CSV file
with open(csv_file, 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow([column[0] for column in cursor.description])  # Write the column headers
    writer.writerows(records)  # Write the records

# Close the cursor and connection
cursor.close()
conn.close()
