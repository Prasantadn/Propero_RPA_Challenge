from selenium import webdriver
import data
import time as t
import os
import pandas as pd
import smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import glob

root_dir = os.getcwd()
download_path = os.path.join(root_dir, 'downloads')
prefs = {"download.default_directory":download_path}
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(executable_path="driver/chromedriver.exe", options=options)
url = data.url

# step1 : open the link
def open_browser():
    driver.get(url)
    driver.maximize_window()
    t.sleep(5)

# step2 : COVID19 > Number of cases
def navigate_to_number_of_cases():
    covid19 = driver.find_element_by_xpath(data.covid19link_xpath)
    covid19.click()
    
    t.sleep(5)
    
    number_of_cases = driver.find_element_by_xpath(data.noofcase_xpath)
    number_of_cases.click()
    t.sleep(5)

# step3 : Downloading Data
def download_data():
    downloadxl = driver.find_element_by_xpath(data.download_xl_xpath)
    downloadxl.click()
    t.sleep(7)

# step4 : Extracting Madhya Pradesh Covid data and saving in email_attachments
def extract_madhyapradesh_data():
    files = glob.glob("./downloads/*")
    latest = max(files, key=os.path.getctime)
    df = pd.read_csv(latest)
    condition = df["Region"] == "Madhya Pradesh"
    df[condition].to_csv("Prasanta_Covid_19_cases.csv")
    t.sleep(7)

# step5 : Saving statewise covid data
def saving_statewise_data():
    d = []
    for tr in driver.find_elements_by_xpath(data.statewise_covid_table_xpath):
        tds = tr.find_elements_by_tag_name('td')
        if tds: 
            d.append([td.text for td in tds])
    def convert_to_csv(l):
        main = l[0]
        i = 0
        s = ""
        temp = []
        while i<len(main):
            for _ in range(6):
                s = s + main[i] + ","
                i = i+1
            s = s[:-1]
            temp.append(s)
            s = ""
            
        first_el = temp[0]
        temp.remove(first_el)
        temp.append(first_el)
        column = "#,State/UT,Confirmed Cases,Active Cases,Cured/Discharged,Death"
        temp.insert(0, column)
        
        content = ""
        for x in temp:
            content = content + x + "\n"
        
        f = open("State_Wise_Covid19_India.csv", "w")
        f.write(content)
        f.close()
    convert_to_csv(d)
    t.sleep(5)
    
# step6 : reading task.robot file and fetching data
def read_robot():
    f = open("task.robot", "r")
    l = f.readlines()
    f.close()

    username=l[0].split("=")[-1].strip()
    password=l[1].split("=")[-1].strip()
    receivers=l[2].split("=")[-1].split(",")

    data = []
    data.append(username)
    data.append(password)
    data.append(receivers)
    return data

# step7 : sending email with attachments
def send_email():
    filelist = ["State_Wise_Covid19_India.csv", "Prasanta_Covid_19_cases.csv"]
    subject = "Propero Challenge Email"
    body = "Hi, This is Prasanta. Sending this email from the robot created for Propero RPA challenge."
    sender_email = read_robot()[0]
    receiver_email = read_robot()[2]
    password = read_robot()[1]

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = ", ".join(receiver_email)
    message["Subject"] = subject

    message.attach(MIMEText(body, "plain"))
    for file in filelist:
        with open(file, "rb") as attachment:
            attachment_part = MIMEBase("application", "octet-stream")
            attachment_part.set_payload(attachment.read())
            encoders.encode_base64(attachment_part)
            attachment_part.add_header(
                "Content-Disposition",
                f"attachment; filename = {file}",
            )
        message.attach(attachment_part)
        text = message.as_string()

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)


if __name__ == "__main__":
    open_browser()
    navigate_to_number_of_cases()
    download_data()
    extract_madhyapradesh_data()
    saving_statewise_data()
    driver.close()
    send_email()
    