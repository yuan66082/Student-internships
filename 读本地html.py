# 这里面的东西不用改，先看【说明.md】

from bs4 import BeautifulSoup

file = open('河南省职业院校实习备案.html', 'rb') 
html = file.read() 
bs = BeautifulSoup(html,"html.parser") # 缩进格式
namelist = bs.find_all(attrs={"lay-title":"学生信息"})
companylist = bs.find_all(attrs={"lay-title":"实习单位信息"})
joblist = bs.find_all(attrs={"lay-title":"实习岗位信息"})
指导学生 = []
z = zip(namelist,companylist,joblist)
for i,j,k in z:
    指导学生.append({'姓名':i.text.strip(),'公司':j.text.strip(),'岗位':k.text.strip()})
print(指导学生)
# print(bs.find_all(attrs={"lay-title":"学生信息"})) # 获取所有的a标签