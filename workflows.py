import os
import csv
import requests
from bs4 import BeautifulSoup

class WorkFlows:
    @staticmethod
    def fetch_workflow_data(cookies_dict):
        page_number = 1
        workflow_data = []

        while True:
            data = {
                'isDirty': 'false',
                'isReload': 'false',
                'ACTION': '',
                'REPORT_FORMAT': '',
                'sortDir': 'ASC',
                'selectAllIds': 'false',
                'deselectAllIds': 'false',
                'searchMode': '1',
                'PROJECT_ID': '1342180104',
                'savedSearchId': '',
                'savedSearchType': '',
                'savedsearches': '-1',
                'savedSearchManageSSList': '-1',
                'savedSearchName': '',
                'savedSearchDescription': '',
                'workFlowStatus_listValidator': '',
                'bidi_workFlowStatus_regexp': '',
                '_bidi_workFlowStatus': '',
                'stepStatus_listValidator': '',
                'bidi_stepStatus_regexp': '',
                '_bidi_stepStatus': '',
                'workFlowId': '',
                'reviewOutcome_listValidator': '',
                'bidi_reviewOutcome_regexp': '',
                '_bidi_reviewOutcome': '',
                'workFlowNo': '',
                'docNo': '',
                'initiator_query': '-- Enter search query here --',
                'assignee_query': '-- Enter search query here --',
                'workFlowName': '',
                'stepName': '',
                '-1': '-1',
                '-11': '',
                '-12': '',
                'rawQueryText': '',
                'groupBy': 'NONE',
                'sortField': 'duedate',
                'pageSize': '10000',
                'numTotal': '126303',
                "workFlowPageNo": str(page_number),
                'allIdsInPage': ['1302666192268857669', '1302666192268857671', '1302666192268857683', '1302666192268857680', '1302666192268857681', '1302666192268857682', '1302666192268857626', '1302666192268857618', '1302666192268857635', '1302666192268857628', '1302666192268857646', '1302666192268857641', '1302666192268857649', '1302666192268857647', '1302666192268857648', '1302666192268857650', '1302666192268857651', '1302666192268857654', '1302666192268857661', '1302666192268857658', '1302666192268857676', '1302666192268857665', '1302666192268857686', '1302666192268857684', '1302666192269016810'],
                'numSelectedOnOtherPage': '0',
                'bidi_configColumns': ['workFlowNo', 'workFlowName', 'docNo', 'docRev', 'docVersion', 'docTitle', 'stepName', 'action', 'assignee', 'dateIn', 'dateDue', 'originalDueDate', 'completedDate', 'stepStatus', 'reviewStatus', 'fileName', 'templateName'],
                '_bidi_configColumns': ''
            }

            url = f"https://ksa1.aconex.com/SearchWorkFlow?ACTION=SEARCH_WORKFLOW&workflowPageNo={str(page_number)}"
            headers = {
            "Sec-Ch-Ua": "\"Not A(Brand\";v=\"99\", \"Google Chrome\";v=\"121\", \"Chromium\";v=\"121\"",
            "Accept": "*/*",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Sec-Ch-Ua-Mobile": "?0",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Sec-Ch-Ua-Platform": "\"Windows\"",
            "Origin": "https://ksa1.aconex.com",
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Dest": "empty",
            "Referer": "https://ksa1.aconex.com/SearchWorkFlow?ACTION=VIEW_SEARCH_PAGE&moduleKey=WORKFLOW&projectId=1342180104&organizationId=1342177323",
            "Accept-Encoding": "gzip, deflate, br",
            "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
            "Connection": "close",
        }
            response = requests.post(url, headers=headers, data=data, cookies=cookies_dict)
            soup = BeautifulSoup(response.text, 'html.parser')
            input_tag = soup.find('input', id='totalResults')
            total_results = int(input_tag['value'])

            table = soup.find('table', class_='dataTable')

            if table:
                for row in table.find_all('tr', class_='dataRow'):
                    cells = row.find_all('td')
                    workflow_data.append({
                    'workflow_no': cells[1].text.strip(),
                    'workflow_name': cells[2].text.strip(),
                    'doc_no': cells[3].text.strip(),
                    'doc_revision': cells[4].text.strip(),
                    'doc_version': cells[5].text.strip(),
                    'doc_title': cells[6].text.strip(),
                    'step_name': cells[7].text.strip(),
                    'action': cells[8].text.strip(),
                    'assigned_to': cells[9].text.strip(),
                    'date_in': cells[10].text.strip(),
                    'date_due': cells[11].text.strip(),
                    'original_due_date': cells[12].text.strip(),
                    'date_completed': cells[13].text.strip(),
                    'step_status': cells[14].text.strip(),
                    'step_outcome': cells[15].text.strip(),
                    'file_name': cells[16].text.strip(),
                    'template_name': cells[17].text.strip(),
                })
                print(f"Page {page_number} done, total results: {len(workflow_data)}")
                page_number += 1    
            else:
                print("Table not found. Check if the HTML structure has changed.")
                break
                
            if len(workflow_data) >= total_results:
                break

        return workflow_data
    
    @staticmethod
    def write_to_csv(data, folder_path):
        csv_file = os.path.join(folder_path, 'workflow_data.csv')
        fieldnames = data[0].keys()

        with open(csv_file, mode='w', newline='', encoding='utf-8') as file:    
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            for row in data:
                writer.writerow(row)
        print("Workflow Exported successfully.")