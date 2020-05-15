import mysql.connector
from docxtpl import DocxTemplate
import glob, os
os.chdir("/home/dojo/templates")

mydb = mysql.connector.connect(
  host="172.18.0.3",
  user="defectdojo",
  passwd="defectdojo",
  database="defectdojo",
  port="3306"
)

#mycursor = mydb.cursor()


#try 
#catch error 
#finally close connection
#cursor = connection.cursor(prepared=True)
#sql_insert_query = """ INSERT INTO Employee (id, Name, Joining_date, salary) VALUES (%s,%s,%s,%s)"""
#insert_tuple_1 = (1, "Json", "2019-03-23", 9000)
#cursor.execute(sql_insert_query, insert_tuple_1)


#mycursor.execute("SELECT * FROM dojo_engagement")

#myresult = mycursor.fetchall()


#for x in myresult:
#  print(x)



clientName_input = input("Client name: ")

#mycursor.execute("SELECT * FROM dojo_product WHERE name = '"+clientName_input+"' ")

#mycursor.execute('SELECT * FROM dojo_product WHERE name = ?', clientName_input)
#sql_select_client_name =  """SELECT * FROM dojo_product WHERE name = %s"""
#select_client_name_data = (clientName_input)
#mycursor.execute(sql_select_client_name, select_client_name_data)

#result_row = mycursor.fetchone()
#number_of_rows = result_row[0]


sql_select_query = ("SELECT * FROM dojo_product WHERE name = '"+clientName_input+"'")
cursor = mydb.cursor()
cursor.execute(sql_select_query)
records = cursor.fetchall()
#print("Total number of rows: ",cursor.rowcount)

#if number_of_rows > 0:
#if cursor.rowcount > 0:
#   print(clientName_input + " found! ID: ")
#else:
#   print("Client not found. Closing script.")

client_id = 0

for row in records:
   client_id = row[0]
#   print("ID = ", row[0], )
#   print("Name = ", row[1], )
#   print("Team Manager ID = ", row[20], )

#print("\n")
#print(client_id)
#print("\n")

sql_select_engagements_query = ("SELECT id,name,product_id,target_start,target_end  FROM dojo_engagement WHERE product_id = %s")
#cursor.mydb.cursor()
cursor.execute(sql_select_engagements_query, (client_id, ))
records_engagements = cursor.fetchall()


print("\n")

print("Found {} engagements. Which engangement do you want to use?".format(cursor.rowcount))

print("0  -  All Engagements")
number_engagement = 1;
for row in records_engagements:
   #print(number_engagement, " - " ,row[1], " | " ,row[3], " |" ,row[4],)
   print(number_engagement, " - " ,row[1],)
   number_engagement += 1

#print for
#print("0 - All Engagements")
#print("1 - January Audit")
#print("2 - February Audit")
#print("3 - March Audit")

engagement_option = input("Select Option (0-" + str(number_engagement - 1) + "): ")
#Here, if 0 (All) is choosen, we will admit you wants full report, maybe future, we can see the yearly option

print("\n")
print("\n")

sql_select_tests_query = ("SELECT id,engagement_id,title FROM dojo_test WHERE engagement_id = %s")
#cursor.mydb.cursor()
cursor.execute(sql_select_tests_query, (engagement_option, ))
records_tests = cursor.fetchall()

#if title is empty use test type query? to present the scan type instead of the title of test
#what happens if we dont have scans, or only one scan
print("Found {} tests. Which ones do you want to use?".format(cursor.rowcount))
#print for
#print("0 - All Scans")
#print("1 - Burp Suite")
#print("2 - Nessus")
#print("3 - Manual")

print("0  -  All Tests")
number_test = 1;
for row in records_tests:
   #print(number_engagement, " - " ,row[1], " | " ,row[3], " |" ,row[4],)
   print(number_test, " - " ,row[2],)
   number_test += 1

scan_option = input("Select Option (0-" + str(number_test - 1) + "): ")

print("\n")
print("\n")

active_findings_option = input("Show Active Findings? No(0), Yes(1): ")

print("\n")
print("\n")

verified_findings_option = input("Show Verified Findings? No(0), Yes(1): ")

print("\n")
print("\n")

false_positive_option = input("Show False Positive Findings? No(0), Yes(1): ")

print("\n")
print("\n")

executive_summary_option = input("Do you need Executive Summary? No(0), Yes(1): ")

print("\n")
print("\n")

finding_notes_option = input("Do you need Finding Notes? No(0), Yes(1): ")

print("\n")
print("\n")

finding_images_option = input("Do you need Finding Images? No(0), Yes(1): ")


print("\n")
print("\n")


templates_count = 0
for file in glob.glob("*.docx"):
    #print(file)
    templates_count += 1

print("Found {} templates, choose one:".format(templates_count))
#print("1 - XXX.docx")
#print("2 - YYY.docx")
#print("3 - ZZZ.docx")

templates_count_help = 1
for file in glob.glob("*.docx"):
    print(templates_count_help, " - " ,file)
    templates_count_help += 1
print("\n")

template_option = input("Select Option (1-" + str(templates_count) + "): ")


print("Generating report, please wait...")


#sql_select_findings_query = ("SELECT title,cwe,severity,description FROM dojo_finding WHERE active = %s AND verified = %s AND false_p = %s AND test_id = %s")
sql_select_findings_query = ("SELECT title,cwe,severity,description FROM dojo_finding WHERE active = 1 AND verified = 1 AND false_p = 0 AND test_id = 4")
data_findings = (active_findings_option, verified_findings_option, false_positive_option, scan_option)
#cursor.execute(sql_select_findings_query, data_findings)
cursor.execute(sql_select_findings_query)
#cursor.execute(sql_select_findings_query, (active_findings_option, verified_findings_option, false_positive_option, scan_option,) )
records_findings = cursor.fetchall()

for row in records_findings:
   print(row[0], " - ",row[1]," - ",row[2]," - ",row[3])


doc = DocxTemplate("/home/dojo/templates/template.docx")
context = { 'engagement_option' : engagement_option }
doc.render(context)
doc.save("generated_doc.docx")

print("Progress Bar")
