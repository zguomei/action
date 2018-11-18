import re
import xlwt
print("meiling")
with open('case_log.txt',"r",encoding='UTF-8') as f:
     test = f.read()
workbook_action = xlwt.Workbook()
sheet = workbook_action.add_sheet('data')  # 创建一个sheet表格
# #print (abc)
abc=re.findall(r"ACTION\s+(.*)\s+\[finished\]\s+result=SUCCESS\s+duration=(.*)sec", test)
#print(abc)
#print(abc[0])
#a=abc[0]
#print(a)
#print(a)
#print(a[0])
i=1
sheet.write(0,0,"Action_name")
sheet.write(0,1,"Cost_time")
sheet.write(0,2,"Occur_number")

for line in abc:
	a=line
	#print(a)
	print(a[0])
	sheet.write(i,0,a[0])
	sheet.write(i, 1, a[1])
	i = i + 1


workbook_action.save('action12.xls')
