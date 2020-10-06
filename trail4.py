# -*- coding: utf-8 -*-
"""
Created on Tue Sep 22 16:03:01 2020

@author: meghanasb
"""

import datetime

import pandas as pd
import numpy as np
from bs4 import BeautifulSoup
from selenium import webdriver
from urllib.parse import urlparse
import os,random,sys,time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import StaleElementReferenceException
from azure.storage.blob import BlobClient
from io import BytesIO
from flask import Flask, jsonify
from flask_cors import CORS 
from flask import request
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
#from tabulate import tabulate
app = Flask(__name__)   
CORS(app)
lists1=[]
class Linkedin_info(object):
    #__slots__=['Name','Location','Profile_title','Connection','Job_Title','Information']
    def __init__ (self,Name,Current_Location,Profile_title,Connection,LinkedIn_url,Email):
        self.Name=Name
        self.Current_Location=Current_Location
        self.Profile_title=Profile_title
        self.Connection=Connection
        self.LinkedIn_url=LinkedIn_url
        self.Email=Email 
    
class Experiences():
    #def __init__(self,Company_name,Total_Duration,Designation,Dates_Empoyed,Employment_Duration,Location,Description):
    def __init__(self,Name,Current_Location,Profile_title,Connection,LinkedIn_url,Email,Company_name,Total_Duration,Designation,Dates_Empoyed,Employment_Duration,Location,Description):
       
        Linkedin_info.__init__(self, Name, Current_Location, Profile_title, Connection, LinkedIn_url, Email)

        
        self.Company_name=Company_name
        self.Total_Duration=Total_Duration
        self.Designation=Designation
        self.Dates_Empoyed=Dates_Empoyed
        self.Employment_Duration=Employment_Duration
        self.Location=Location
        self.Description=Description

@app.route('/ScrapeData/v1.0/api/inputJson', methods=['POST']) 
def extraction():
    try:
        options = webdriver.ChromeOptions() 
        options.headless=True
        browser = webdriver.Chrome(ChromeDriverManager().install(),options=options)
       # browser=webdriver.Chrome(ChromeDriverManager().install())
        browser.get('https://www.linkedin.com/uas/login')
        #file=open('Config.txt')
        #lines=file.readlines()
        username='wmautomationnous@gmail.com'
        password='Nous@123'
        elementId=browser.find_element_by_id('username')
        elementId.send_keys(username)
        elementId=browser.find_element_by_id('password')
        elementId.send_keys(password)
        elementId.submit()
    except StaleElementReferenceException:
     pass
    req_data = request.get_json()
    columnName=req_data['columnName'] 
    blob_name=req_data['blob_name']
    Email_id=req_data['Email_id']
    #blob=BlobClient.from_connection_string(conn_str="DefaultEndpointsProtocol=https;AccountName=meghana;AccountKey=EGJXMigvJl4Vm1AqO4xzEPKdp2BJoyKSB3Yez+21qaLaZ/AQdOTgpvWX5Cdx8Ul+zzCxKa0cU328LyBY0sspew==;EndpointSuffix=core.windows.net",container_name='linkedinurl',blob_name=str(blob_name)+'.xlsx')  
    blob=BlobClient.from_connection_string(conn_str="DefaultEndpointsProtocol=https;AccountName=wmautomationstorage;AccountKey=bM1imp3ZuspUNtqzzR0vtDxm9Qxt+AVsBjVjeJrZDLsO+bYDuqQTpRdggm+MppVm6xtkvAfX36ZjP0NInYNAmA==;EndpointSuffix=core.windows.net",container_name='wmautomationblobcontainer',blob_name=str(blob_name)+'.xlsx')
    linkedinurl=blob.download_blob().readall()
    df = pd.read_excel(linkedinurl)
    for i in df.index:
        link=df[columnName][i]
        browser.get(link)
        try:
         elementId=browser.find_element_by_link_text('Contact info')
         browser.implicitly_wait(2)
         elementId.click()
        except StaleElementReferenceException:
         pass 
        #elementId=browser.find_element_by_link_text('Contact info')
        #webdriver.ActionChains(browser).move_to_element(elementId ).click(elementId ).perform()
        SCROLL_PAUSE_TIME=2
        #get scroll height
        last_height=browser.execute_script('return document.body.scrollHeight')
        for i in range(3):
                    browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(SCROLL_PAUSE_TIME)
                    new_height=browser.execute_script('return document.body.scrollHeight')
                    if new_height==last_height:
                        break
       # elementId=browser.find_element_by_link_text('Contact info')
       # webdriver.ActionChains(browser).move_to_element(elementId ).click(elementId ).perform()
        
        
        src=browser.page_source
        soup=BeautifulSoup(src,'lxml') 
        name_div=soup.find('div',{'class':'flex-1 mr5'})
        name_loc=name_div.find_all('ul')
        name=name_loc[0].find('li').get_text().strip()
        loc=name_loc[1].find('li').get_text().strip()
        profile_title=name_div.find('h2').get_text().strip()
        connection=name_loc[1].find_all('li')
        connection=connection[1].get_text().strip()
        ContactInfo=soup.find('div',{'class':'artdeco-modal__content ember-view'})
       # Linkedin_url=ContactInfo.find('div',{'class':'pv-contact-info__ci-container t-14'})
        #Linkedin_url=Linkedin_url.find('a').get_text().strip()
        Linkedin_url=link
        Email=ContactInfo.find('section',{'class':'pv-contact-info__contact-type ci-email'})
        if Email is not None:
                        Email=Email.find('div',{'class':'pv-contact-info__ci-container t-14'})
                        Email=Email.find('a').get_text().strip()
        else:
                        Email=""
        
        #data=Linkedin_info(name,loc,profile_title,connection,'','')
        data=Linkedin_info(name,loc,profile_title,connection,Linkedin_url,Email)
        #lists1.append(data.__dict__)                
        exp_section=soup.find('section',{'class':'pv-profile-section experience-section ember-view'})
        if exp_section is not None:
               getTags=exp_section.find('ul')
               li_tags=exp_section.find('ul').find('div')
               a_tags=li_tags.find('a')
               exp_sections=exp_section.find('ul').find_all('ul',{'class':'pv-entity__position-group mt2'})
               #nodescrip=exp_sections.find_all[1]('div',{'class':'pv-entity__role-details-container pv-entity__role-details-container--timeline pv-entity__role-details-container--bottom-margin'})
               if not exp_sections:
                   Designation=a_tags.find_all('h3')[0].get_text().strip()
                   company_name=a_tags.find_all('p',{'class':'pv-entity__secondary-title t-14 t-black t-normal'})[0].get_text().strip()
                   Dates_Employed=a_tags.find_all('h4',{'class':'pv-entity__date-range t-14 t-black--light t-normal'})[0].find_all('span')[1].get_text().strip()
                   Employment_Duration=a_tags.find_all('span',{'class':'pv-entity__bullet-item-v2'})[0].get_text().strip()
                  
                   try:
                       Location=a_tags.find_all('h4',{'class':'pv-entity__location t-14 t-black--light t-normal block'})[0].find_all('span')[1].get_text().strip()
                       
                   except IndexError:
                     Location=""
                   
                   info=exp_section.find_all('p',{'class':'pv-entity__description t-14 t-black t-normal inline-show-more-text inline-show-more-text--is-collapsed ember-view'})
                   information =info[0].get_text().strip('see more')
                  
                   total_Duration=Employment_Duration
                   exp_data=Experiences(name,loc,profile_title,connection,Linkedin_url,Email,company_name,total_Duration,Designation,Dates_Employed,Employment_Duration,Location,information)
                   #exp_data=Experiences(name,loc,profile_title,connection,"","",company_name,total_Duration,Designation,Dates_Employed,Employment_Duration,Location,information)
                   lists1.append(exp_data.__dict__)  
               else:
                   #Experence={}
                   #j=0
                   lists2=[]
                   company_name=a_tags.find_all('h3')[0].find_all('span')[1].get_text().strip()
                   total_Duration=a_tags.find_all('h4')[0].find_all('span')[1].get_text().strip()
                  
                   for el in exp_sections[0].find_all('li'):
                    
                     
                     Designation=el.find_all('h3')[0].find_all('span')[1].get_text().strip()
                     Date_emp=el.find_all('h4',{'class':'pv-entity__date-range t-14 t-black--light t-normal'})[0].find_all('span')[1].get_text().strip()         
                     Employment_Dur=el.find_all('h4',{'class':'t-14 t-black--light t-normal'})[0].find_all('span')[1].get_text().strip()
                     try:
                      Locat=el.find_all('h4',{'class':'pv-entity__location t-14 t-black--light t-normal block'})[0].find_all('span')[1].get_text().strip()
                     except IndexError:
                      Locat=""
                     try:
                         informat=el.find_all('p',{'class':'pv-entity__description t-14 t-black t-normal inline-show-more-text inline-show-more-text--is-collapsed ember-view'})[0].get_text().strip("see more")
                     except IndexError:
                         informat=""
                     # data2=Linkedin_info.getExperiences1(Designation,Date_emp,Employment_Dur,Locat,informat)
                     # Experence[j]=data2
                     # j+=1
                     if company_name in lists2 and total_Duration in lists2 and data in lists2:
                      
                       exp_data=Experiences("","","","","","","","",Designation,Date_emp,Employment_Dur,Locat,informat) 
                       lists1.append(exp_data.__dict__)  
                     else:
                        lists2.append(company_name)
                        lists2.append(total_Duration)
                        lists2.append(data)
                        exp_data=Experiences(name,loc,profile_title,connection,Linkedin_url,Email,company_name,total_Duration,Designation,Date_emp,Employment_Dur,Locat,informat)
                        #exp_data=Experiences(name,loc,profile_title,connection,"","",company_name,total_Duration,Designation,Date_emp,Employment_Dur,Locat,informat)
                        lists1.append(exp_data.__dict__)  
              
    else:
        pass;
    dff=pd.DataFrame().from_dict(lists1,orient='columns')  
    output=BytesIO()
    writer=pd.ExcelWriter(output,engine='xlsxwriter')
    dff.to_excel(writer,index=False)
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    # Add some cell formats.
   # format1 = workbook.add_format({'num_format': '#,##0.00'})
   # format2 = workbook.add_format({'num_format': '0%'})
    bg_format1 = workbook.add_format({'bg_color': '#EEEEEE'}) # blue cell background color
    bg_format2 = workbook.add_format({'bg_color': '#DDDDDD'}) # white cell background color
    # for i in range(30): # integer odd-even alternation 
    #  worksheet.set_row(i, cell_format=(bg_format1 if i%2==0 else bg_format2))
    wrap_format=workbook.add_format({'text_wrap':True})
   # worksheet.set_default_row(20)
    # Note: It isn't possible to format any cells that already have a format such
    # as the index or headers or any cells that contain dates or datetimes.
    # Set the column width and format.

    worksheet.set_column('A:A', 20 ,wrap_format) 
    worksheet.set_column('B:B', 20 ,wrap_format)
    worksheet.set_column('C:C', 20 ,wrap_format)
    worksheet.set_column('D:D', 18 ,wrap_format)
    worksheet.set_column('E:E', 50, wrap_format)
    worksheet.set_column('F:F', 18, wrap_format)
    worksheet.set_column('G:G', 20, wrap_format)
    worksheet.set_column('H:H', 15, wrap_format)
    worksheet.set_column('I:I', 20, wrap_format)
    worksheet.set_column('J:J', 20, wrap_format)
    worksheet.set_column('K:K', 20, wrap_format)
    worksheet.set_column('L:L', 18, wrap_format)   
    worksheet.set_column('M:M', 50 )
    #worksheet.autoFitColumn(12)
   # worksheet.set_default_row(30)

    writer.save()
    processed_data=output.getvalue()
   
    uniq_filename = str(datetime.datetime.now().date()) + '_' + str(datetime.datetime.now().time()).replace(':', '.')
    #blob=BlobClient.from_connection_string(conn_str="DefaultEndpointsProtocol=https;AccountName=meghana;AccountKey=EGJXMigvJl4Vm1AqO4xzEPKdp2BJoyKSB3Yez+21qaLaZ/AQdOTgpvWX5Cdx8Ul+zzCxKa0cU328LyBY0sspew==;EndpointSuffix=core.windows.net",container_name='linkedinurl',blob_name=str(blob_name)+str(uniq_filename)+"Scrapped"+'.xlsx')  
    blob=BlobClient.from_connection_string(conn_str="DefaultEndpointsProtocol=https;AccountName=wmautomationstorage;AccountKey=bM1imp3ZuspUNtqzzR0vtDxm9Qxt+AVsBjVjeJrZDLsO+bYDuqQTpRdggm+MppVm6xtkvAfX36ZjP0NInYNAmA==;EndpointSuffix=core.windows.net",container_name='wmautomationblobcontainer',blob_name=str(blob_name)+str(uniq_filename)+"Scrapped"+'.xlsx')
    blob.upload_blob(processed_data,overwrite=True)
    
    #SENDING EMAIL
    sender_email = 'wmautomationnous@gmail.com'
    password = 'Nous@123'
    receiver_email =  Email_id
    message = MIMEMultipart("alternative")
    message["Subject"] = "Scrapped LinkedIn Data"
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Cc"]='ambarjain@nousinfo.com'
    #AccountName = 'meghana'
   # container_name='linkedinurl'
    AccountName = 'wmautomationstorage'
    container_name='wmautomationblobcontainer'
    
    blob_name = str(blob_name)+str(uniq_filename)+"Scrapped"+'.xlsx'

    url = f"https://{AccountName}.blob.core.windows.net/{container_name}/{blob_name}"
    #Create the plain-text and HTML version of your message
    text = """\
        Please click Below Link to Download Excel
    """
    html = """\
    <html>
      <body>
          <p>Hi<br>
               Please click Below Link to Download Excel<br>
           <a href="%s" >Downloadable link</a> 
         </p>
         <p>Regards,<br>
           WmAutomationTeam
         </p>
      </body>
    </html>
    """ %(url) 
    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")
    
    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    message.attach(part2)
    
    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(
            sender_email, receiver_email, message.as_string()
        )

    
    
    return jsonify("Status:Success","Message:Resources Uploaded Successfully! and mail sent"), 200
    

if __name__ == '__main__':
    app.run()
