import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('VM.xlsx')
worksheet = workbook.add_worksheet()

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

worksheet.write(0, 0, 'Name')
worksheet.write(0, 1, 'CPU')
worksheet.write(0, 2, 'SSD')
worksheet.write(0, 3, 'RAM')


# Iterate over the data and write it out row by row.
#for item, cost in (expenses):
#    worksheet.write(row, col,     item)
#    worksheet.write(row, col + 1, cost)
#    row += 1

# Write a total using a formula.
#worksheet.write(row, 0, 'Total')
#worksheet.write(row, 1, '=SUM(B1:B4)')




my_dict = {}
currentCompanyName = ""

with open('myoutput.txt') as f:
   for line in f:
       
       splitLineName = line.split('-', 1)[0].strip().lower()
       
       splitLineNumber = line.split(' ', 1)
       
       if(splitLineNumber[0] == "CPU:" or splitLineNumber[0] == "SSD:" or splitLineNumber[0] == "RAM:"):
           #print("")
           value = float(splitLineNumber[1].strip())
           
           if(splitLineNumber[0] == "CPU:"):
               worksheet.write(row, 1, value)
           if(splitLineNumber[0] == "SSD:"):
               worksheet.write(row, 2, value)
           if(splitLineNumber[0] == "RAM:"):
               worksheet.write(row, 3, value)
               
               
           
           if my_dict[currentCompanyName] == "empty":
               valuesDic = {}
               valuesDic[splitLineNumber[0]] = value
               my_dict[currentCompanyName] = valuesDic
           elif splitLineNumber[0] not in my_dict[currentCompanyName].keys():
           
               my_dict[currentCompanyName][splitLineNumber[0]] = value
               
         
           else:
               temp = my_dict[currentCompanyName][splitLineNumber[0]]
               my_dict[currentCompanyName][splitLineNumber[0]] = temp + value
       else:
          
           currentCompanyName = splitLineName
           
           row = row + 1
           
           #write worksheet
           worksheet.write(row, 0, line)
           
           if splitLineName not in my_dict.keys():
               my_dict[splitLineName] = "empty"
               
       
       if 'str' in line:
          break

workbook.close()          
          
          
workbook = xlsxwriter.Workbook('Totals.xlsx')
worksheet = workbook.add_worksheet()

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

worksheet.write(0, 0, 'Name')
worksheet.write(0, 1, 'CPU')
worksheet.write(0, 2, 'SSD')
worksheet.write(0, 3, 'RAM')

for key in my_dict:
    row = row + 1
    #print(key)
    worksheet.write(row, 0, key)
    
    for innerkey in my_dict[key]:
        if(innerkey == "CPU:"):
            worksheet.write(row, 1, my_dict[key][innerkey])
        if(innerkey == "SSD:"):
            worksheet.write(row, 2, my_dict[key][innerkey])
        if(innerkey == "RAM:"):
            worksheet.write(row, 3, my_dict[key][innerkey])
               
        #print(innerkey + " ", end = '')
        #print(my_dict[key][innerkey])



workbook.close()
