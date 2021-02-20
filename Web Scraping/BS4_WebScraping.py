import requests
from bs4 import BeautifulSoup
import xlwt


url = r'https://tracxn.com/d/trending-themes/Startups-in-V2X'
reqs = requests.get(url)
soup = BeautifulSoup(reqs.text,'html.parser')
#print(soup)
links = []
link_to_site = soup.find_all("a", attrs={"style": 'color:#0000EE; text-decoration: none;'})
about_site = soup.find_all('p',attrs={"style":'line-height: 1.6; margin-top: 10px; margin-bottom: 30px; font-size:12pt; font-family: Arial; color:#293348;'})
count = 0
for a in link_to_site:

    dict1 = {}
    dict1['name'] = a.getText()
    dict1['site'] = a.get('href')
    #print(a.get('href'))
    #print(a.getText())
    about = about_site[0]
    dict1['about'] = about.getText()
    links.append(dict1)

print(links)

#for about in about_site:
#    print(about.getText())


para1 = soup.find_all("p", attrs = {"style":'line-height: 1.6; margin-top: 0pt; margin-bottom: 0pt; font-weight: 700'})
para2 = soup.find_all("p", attrs = {"style":'line-height: 1.6; margin-top: 0pt; margin-bottom: 0pt;'})

#print(len(para2))

list1= []
flag = True
dict1 = {}


for k in range(0,len(para2)):
    i = para1[k]
    j = para2[k]
    if(i.getText()=="Founded Year"):
        if(flag==True):
            dict1[i.getText()] = j.getText()
            flag=False
        else:
          list1.append(dict1)
          dict1={}
          dict1[i.getText()] = j.getText()
    else:
        dict1[i.getText()] = j.getText()
    if(k==len(para1)-1):
        list1.append(dict1)
print(list1)


workbook = xlwt.Workbook()

sheet = workbook.add_sheet("Company's Details")

#Now the length of all section will be same
sheet.write(0,0,'Name')
sheet.write(0,1,'Link')
sheet.write(0,2,'About')
sheet.write(0,3,'Other Info')
for i in range(len(list1)):
    sheet.write(i+1,0,links[i]["name"])
    sheet.write(i+1,1,links[i]["site"])
    sheet.write(i+1,2,links[i]["about"])
    sheet.write(i+1,3,str(list1[i]))

workbook.save("companies.xls") 
