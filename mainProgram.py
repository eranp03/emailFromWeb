import xlsxwriter, xlrd,re,requests
from xlutils.copy import copy
from bs4 import BeautifulSoup

def getMailFromWeb(url):
    s = "no mail"
    allLinks = [];mails=[]
    try:
        response = requests.get(url)
        soup=BeautifulSoup(response.text,'html.parser')
        links = [a.attrs.get('href') for a in soup.select('a[href]') ]
        for i in links:
            if(("contact" in i or "Contact")or("Career" in i or "career" in i))or('about' in i or "About" in i)or('Services' in i or 'services' in i):
                allLinks.append(i)


        def findMails(soup):
            for name in soup.find_all('a'):
                if(name is not None):
                    emailText=name.text
                    match=bool(re.match('[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$',emailText))
                    if('@' in emailText and match==True):
                        emailText=emailText.replace(" ",'').replace('\r','')
                        emailText=emailText.replace('\n','').replace('\t','')
                        #if(len(mails)==0)or(emailText not in mails):
                         #   print(emailText)
                        mails.append(emailText)

        allLinks=set(allLinks)
        for link in allLinks:
            if(link.startswith("http") or link.startswith("www")):
                r=requests.get(link)
                data=r.text
                soup=BeautifulSoup(data,'html.parser')
                findMails(soup)

            else:
                newurl=url+link
                r=requests.get(newurl)
                data=r.text
                soup=BeautifulSoup(data,'html.parser')
                findMails(soup)

        mails=set(mails)
        return mails

        if(len(mails)==0):
            mails.append("no Email")
            return mails
    except:
        try:
            s = "no mail"
            allLinks = [];
            mails = []
            r = requests.get(url)
            data = r.text
            soup = BeautifulSoup(data, "html.parser")

            for k in soup.find_all(href=re.compile("mailto")):
                mails.append(k.string)
            return mails
        except:
            mails.append("no Email")
            return mails



wb = xlrd.open_workbook("List of the Top Shopify Design and Development TEST.xlsx")
# get the first worksheet
sheet = wb.sheet_by_index(0)
print("eran")
for i in range(1,10):
    cell = sheet.cell(i, 4).value
    print("searching email address at: ", cell)
    print("the mail for site: ", cell, "mail is: ", getMailFromWeb(cell))





