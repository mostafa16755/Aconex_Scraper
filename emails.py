import os
import csv
import requests
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#https://ksa1.aconex.com/hub/index.html

class Emails:
    @staticmethod
    def login(user_name, password):
        driver = webdriver.Chrome()
        driver.get('https://ksa1.aconex.com/hub/index.html')

        user_name_field = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="userName|input"]')))
        user_name_field.send_keys(user_name)
        next_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="nextButton_oj1|text"]')))
        next_btn.click()

        password_field = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ui-id-1|input"]')))
        password_field.send_keys(password)
        login_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="gbu-page-page-user-credentials"]/div/div/div[2]/div/div[2]/div[2]/oj-form-layout/div/div[3]/div/oj-button/button/div')))
        login_btn.click()

        time.sleep(15)
        cookies = driver.get_cookies()
        cookies_dict = {cookie['name']: cookie['value'] for cookie in cookies}
        csrf_token = cookies_dict['XSRF-TOKEN']
        cookies_string = "; ".join([f"{k}={v}" for k, v in cookies_dict.items()])

        print("Login Successfully!")
        driver.quit()
        return cookies_dict, cookies_string, csrf_token
    
    @staticmethod
    def fetch_emails(cookies_string, csrf_token):
        all_data = []
        page_number = 1
        while True:
            reqUrl = f"https://ksa1.aconex.com/internal/projects/1342180104/users/1342330444/mails/ALLBOX?pageNo={page_number}&pageSize=1000&selectedAttributes=hasAttachments&selectedAttributes=documentNo&selectedAttributes=subject&selectedAttributes=sentDate&selectedAttributes=fromUserName&selectedAttributes=fromOrganizationName&selectedAttributes=toOrganizationName&selectedAttributes=toUserName&selectedAttributes=toStatus&selectedAttributes=type&selectedAttributes=attributes&selectedAttributes=mailHyperlink&selectedAttributes=confidential&selectedAttributes=closedOutByOrg&selectedAttributes=closedOutByUser&selectedAttributes=closedOutDate&selectedAttributes=InvoiceDate_date&selectedAttributes=InvoiceNo_singleLineText&selectedAttributes=Discipline_singleSelect&selectedAttributes=Districts_multiSelect&selectedAttributes=toDueDate&selectedAttributes=PaymentType_singleSelect&selectedAttributes=PaymentDueDate_date&showMyMailOnly=false&sortDirection=DESC&sortField=sentdate"

            headersList = {
                "Accept": "application/json, text/plain, */*",
                "Accept-Language": "en-US,en;q=0.9,ar-MA;q=0.8,ar;q=0.7,en-GB;q=0.6,ar-EG;q=0.5",
                "Connection": "keep-alive",
                "Cookie": cookies_string,
                "Referer": "https://ksa1.aconex.com/rsrc/20240329.1234/en_US_DOC/mailSearch/CorrespondenceSearch.html?mailbox=ALLBOX&moduleKey=CORRESPONDENCE&projectId=1342180104&organizationId=1342177323",
                "Sec-Fetch-Dest": "empty",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Site": "same-origin",
                "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
                "X-XSRF-TOKEN": csrf_token,
                "sec-ch-ua": '"Google Chrome";v="123", "Not:A-Brand";v="8", "Chromium";v="123"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": "Linux" 
            }

            payload = ""

            response = requests.request("GET", reqUrl, data=payload,  headers=headersList)

            data = response.json()
            all_data.extend(data['mails'])
            print(f"Extract Page Number: {page_number} From Emails")

            if len(all_data) >= data["total"]:
                break
            page_number += 1

        return all_data

    @staticmethod
    def write_emails_to_csv(emails, folder_path, filename):
        fieldnames = ['Mail No', 'Subject', 'Date', 'From', 'From Organization', 'To Organization', 'Recipient', 'Status', 'Type', 'ID',
                    'Has Attachments', 'Closed Out By Org', 'Closed Out By User', 'Closed Out Date', 'Sent Date Time', 'To Due Date', 
                    'Parent Status', 'Overall Status', 'Mail Hyperlink', 'Discipline', 'Districts', 'Due', 'Invoice Date', 'Invoice No',
                    'Payment Due Date', 'Payment Type', 'Attribute 1', 'Confidential']
        
        try:
            file_path = os.path.join(folder_path, filename)

            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                for email in emails:
                    writer.writerow({
                        'Mail No': email.get('documentNo', ''),
                        'Subject': email.get('subject', ''),
                        'Date': email.get('sentDate', ''),
                        'From': email.get('fromUserName', ''),
                        'From Organization': email.get('fromOrganizationName', ''),
                        'To Organization': email.get('recipients', [{'toOrganizationName': ''}])[0].get('toOrganizationName', ''),
                        'Recipient': email.get('recipients', [{'toUserName': ''}])[0].get('toUserName', ''),
                        'Status': email.get('recipients', [{'toStatus': ''}])[0].get('toStatus', ''),
                        'Type': email.get('type', ''),
                        'ID': email.get('id', ''),
                        'Has Attachments': email.get('hasAttachments', ''),
                        'Closed Out By Org': email.get('closedOutByOrg', ''),
                        'Closed Out By User': email.get('closedOutByUser', ''),
                        'Closed Out Date': email.get('closedOutDate', ''),
                        'Sent Date Time': email.get('sentDateTime', ''),
                        'To Due Date': email.get('toDueDate', ''),
                        'Parent Status': email.get('parentStatus', ''),
                        'Overall Status': email.get('overallStatus', ''),
                        'Mail Hyperlink': email.get('mailHyperlink', ''),
                        'Discipline': email.get('discipline', ''),
                        'Districts': email.get('districts', ''),
                        'Due': email.get('due', ''),
                        'Invoice Date': email.get('invoiceDate', ''),
                        'Invoice No': email.get('invoiceNo', ''),
                        'Payment Due Date': email.get('paymentDueDate', ''),
                        'Payment Type': email.get('paymentType', ''),
                        'Attribute 1': email.get('attributes', ''),
                        'Confidential': email.get('confidential', {}).get('displayText', '')
                    })
                print("Emails Data Exported Successfully!")
        except Exception as e:
            print(f"Error While Write Mails Data in CSV {str(e)}")
