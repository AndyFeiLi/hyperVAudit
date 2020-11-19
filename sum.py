import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('VM.xlsx')
worksheet = workbook.add_worksheet("VM")

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
           
           ########hard coded hacky stuff here - maintain this or rename VMs#########
           
           if(currentCompanyName == "riverstone"):
               currentCompanyName = "rsc"
           elif(currentCompanyName == "rscapp01"):
               currentCompanyName = "rsc"
           elif(currentCompanyName == "rscdc01"):
               currentCompanyName = "rsc"
           elif(currentCompanyName == "tbcwestperth"):
               currentCompanyName = "tbcwp"
           elif(currentCompanyName == "htx2"):
               currentCompanyName = "htx"
           elif(currentCompanyName == "nbc"):
               currentCompanyName = "northfleet"
           elif(currentCompanyName == "mgmt"):
               currentCompanyName = "cloudconnect"
           elif(currentCompanyName == "vm"):
               currentCompanyName = "cloudconnect"
           elif(currentCompanyName == "mgmt"):
               currentCompanyName = "cloudconnect"
           elif(currentCompanyName == "3cx"):
               currentCompanyName = "cloudconnect"
           elif(currentCompanyName == "cloudcore"):
               currentCompanyName = "cloudconnect"
           elif(currentCompanyName == "radius"):
               currentCompanyName = "cloudconnect"
           elif(currentCompanyName == "vm"):
               currentCompanyName = "cloudconnect"
           elif(currentCompanyName == "p1"):
               currentCompanyName = "cloudconnect"
           
           ############################################
           
           row = row + 1
           
           #write worksheet
           worksheet.write(row, 0, line)
           
           if currentCompanyName not in my_dict.keys():
               my_dict[currentCompanyName] = "empty"
               
       
       if 'str' in line:
          break

#workbook.close()          
          
          
#workbook = xlsxwriter.Workbook('Totals.xlsx')
worksheet1 = workbook.add_worksheet("Totals")

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

worksheet1.write(0, 0, 'Name')
worksheet1.write(0, 1, 'CPU')
worksheet1.write(0, 2, 'SSD')
worksheet1.write(0, 3, 'RAM')

for key in my_dict:
    row = row + 1
    #print(key)
    
    ########hard coded hacky stuff here - maintain this or rename VMs#########
           
    name = "default"
    
    if(key == "a2w"):
        name = "Aspire2 Wealth"
    elif(key == "acc"):
        name = "Access Solutions"
    elif(key == "cathedral"):
        name = "Cathedral Office Products"
    elif(key == "htx"):
        name = "Haultrax"
    elif(key == "jwms"):
        name = "Jetwave Marine Services"               
    elif(key == "kais"):
        name = "Kais Contractors"            
    elif(key == "pmm"):
        name = "PMM Wealth Advisors"             
    elif(key == "stk"):
        name = "Stickman Tribe"        
    elif(key == "tbcwp"):
        name = "Patter Merchants"             
    elif(key == "kng"):
        name = "Kongsberg Maritime"   
    elif(key == "sleep"):
        name = "Sleep Studies"        
    elif(key == "sleep"):
        name = "Sleep Studies"        
    elif(key == "rsc"):
        name = "Riverstone Custom Homes"  
    elif(key == "firstnational"):
        name = "First National Real Estate"
    elif(key == "northfleet"):
        name = "Northfleet Bus Contractors"  
    elif(key == "plr"):
        name = "Polaris Realty"    
    elif(key == "pme"):
        name = "Plantman Equipment"    
    elif(key == "efw"):
        name = "European Foods Wholesalers"        
    elif(key == "elev8"):
        name = "Insphire Australia"  
    elif(key == "get.trakka"):
        name = "GetTrakka"
    elif(key == "hiresociety"):
        name = "Hire Society"
    elif(key == "cloudconnect"):
        name = "Cloud Connect"
    elif(key == "tax"):
        name = "Taxcorp"
    elif(key == "pba"):
        name = "Pilbara Airlines"
        
   ############################################ 
    
    if(name == "default"):
        worksheet1.write(row, 0, key)
    else:
        worksheet1.write(row, 0, name)
    
    for innerkey in my_dict[key]:
        if(innerkey == "CPU:"):
            worksheet1.write(row, 1, my_dict[key][innerkey])
        if(innerkey == "SSD:"):
            worksheet1.write(row, 2, my_dict[key][innerkey])
        if(innerkey == "RAM:"):
            worksheet1.write(row, 3, my_dict[key][innerkey])
               
        #print(innerkey + " ", end = '')
        #print(my_dict[key][innerkey])



workbook.close()
