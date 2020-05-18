import mysql.connector
from docxtpl import DocxTemplate
import glob, os
os.chdir("/home/dojo/templates")

mydb = mysql.connector.connect(
  host="172.18.0.2",
  user="defectdojo",
  passwd="defectdojo",
  database="defectdojo",
  port="3306"
)

#try
#catch error
#finally close connection
#cursor = connection.cursor(prepared=True)

client_name_input = input("Client Name: ")

cursor = mydb.cursor()

sql_select_client_name_query = ("SELECT id FROM dojo_product WHERE name = %s")
cursor.execute(sql_select_client_name_query, (client_name_input, ))
records_client_name = cursor.fetchall()

if cursor.rowcount > 0:
   print(client_name_input + " found!")
else:
   print("Client doesn't exist. Exiting...")
   sys.exit()


client_id = 0
for row in records_client_name:
   client_id = row[0]


sql_select_engagements_query = ("SELECT id,name,product_id,target_start,target_end  FROM dojo_engagement WHERE product_id = %s")
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


engagement_option = int(input("Select Option (0-" + str(number_engagement - 1) + "): "))
#Here, if 0 (All) is choosen, we will admit you wants full report, maybe future, we can see the yearly option

print("\n")
print("\n")

#if engagement_option == 0:
#  sql_select_tests_query = ("SELECT id,engagement_id,title FROM dojo_test")
#  cursor.execute(sql_select_tests_query, (engagement_option, ))
#  records_tests = cursor.fetchall()
#else:
sql_select_tests_query = ("select dojo_test.id, engagement_id, concat( if(title is not null or title = '', concat(dojo_test.title,'(',dojo_test_type.name,')' ), dojo_test_type.name) ) as nameEST  from dojo_test INNER JOIN dojo_test_type ON dojo_test_type.id = test_type_id where engagement_id = %s")
cursor.execute(sql_select_tests_query, (engagement_option, ))
records_tests = cursor.fetchall()

#  if cursor.rowcount > 0:
#    #if title is empty use test type query? to present the scan type instead of the title of test
#    #what happens if we dont have scans, or only one scan
print("Found {} tests. Which ones do you want to use?".format(cursor.rowcount))

n = cursor.rowcount + 1
tests_array = [None] * n
tests_array[0] = 0
print("0  -  All Tests")
number_test = 1;
for row in records_tests:
   print(number_test, " - " ,row[2],)
   tests_array[number_test] = row[0]
   number_test += 1

#print("end")
#for x in range(len(tests_array)):
#    print(tests_array[x])

scan_option = int(input("Select Option (0-" + str(number_test - 1) + "): "))
#print(type(scan_option))
helper_test = 0 #control
helper_number_test = 0 #row
#print("you choosed:" ,scan_option)
#print("no useful:",helper_number_test)

#print("Array 1 ID is: ",tests_array[1])

for test in tests_array:
   if scan_option == helper_number_test:
     helper_test = test
     helper_number_test += 1
     #print("here: ",helper_test)
     #print("here: ",helper_number_test)
   else:
     helper_number_test += 1
     #print("no:",helper_number_test)
     #print("no:",helper_test)


#print("helper test:",helper_number_test)
#print("helper test is: ",helper_test)
print("\n")
print("\n")

active_findings_option = int(input("Show Active Findings? No(0), Yes(1): "))

print("\n")
print("\n")

verified_findings_option = int(input("Show Verified Findings? No(0), Yes(1): "))

print("\n")
print("\n")

false_positive_option = int(input("Show False Positive Findings? No(0), Yes(1): "))

print("\n")
print("\n")

executive_summary_option = int(input("Do you need Executive Summary? No(0), Yes(1): "))

print("\n")
print("\n")

finding_notes_option = int(input("Do you need Finding Notes? No(0), Yes(1): "))

print("\n")
print("\n")

finding_images_option = int(input("Do you need Finding Images? No(0), Yes(1): "))


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

template_option = int(input("Select Option (1-" + str(templates_count) + "): "))


print("Generating report, please wait...")


#sql_select_findings_query = ("SELECT title,cwe,severity,description FROM dojo_finding WHERE active = %s AND verified = %s AND false_p = %s AND test_id = %s")
sql_select_findings_query = ("SELECT title,cwe,severity FROM dojo_finding WHERE active = %s AND verified = %s AND false_p = %s AND test_id = %s")
data_findings = (active_findings_option, verified_findings_option, false_positive_option, helper_test)
cursor.execute(sql_select_findings_query, data_findings)

records_findings = cursor.fetchall()

findings_array = [[0 for x in range(3)] for y in range(cursor.rowcount)]
numberhelper = 0
for row in records_findings:
   #print(row[0], " - ",row[1]," - ",row[2])
   findings_array[numberhelper][0] = row[0]
   findings_array[numberhelper][1] = row[1]
   findings_array[numberhelper][2] = row[2]
   numberhelper += 1

for x in range(len(findings_array)):
   print(findings_array[x])


doc = DocxTemplate("/home/dojo/templates/template.docx")
context = { 'engagement_option' : engagement_option }

#'findings': [
#        {'name': finding, 'severity': severity},
#
#        {
#            'status': status,
#            'dateDiscovered': dateDiscovered,
#            'age': age,
#            'reporter': reporter,
#            'description': description
#        },
#    ],

doc.render(context)
doc.save("generated_doc.docx")

print("Progress Bar")
