import os
import requests
import pandas as pd
import openpyxl

class Documents:
        
    @staticmethod
    def fetch_docs(cookies_dict, cookies_string, csrf_token):

        all_results = []
        page_num = 1

        while True:
            print(f"Extract Page Number: {page_num}")

            reqUrl = "https://ksa1.aconex.com/SearchControlledDoc?SEARCH_ACTION=7&CORR_EDIT_SCREENMODEL_ID=null&IS_AJAX_REQUEST=true&projectId=1342180104"

            headersList = {
            "Accept": "*/*",
            "Accept-Language": "en-US,en;q=0.9,ar-MA;q=0.8,ar;q=0.7,en-GB;q=0.6,ar-EG;q=0.5",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Cookie": cookies_string,
            "Origin": "https://ksa1.aconex.com",
            "Referer": "https://ksa1.aconex.com/SearchControlledDoc?tab=0&SEARCH_ACTION=0&moduleKey=DOC&projectId=1342180104&organizationId=1342177323",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
            "X-Requested-With": "XMLHttpRequest",
            "sec-ch-ua": '"Google Chrome";v="123", "Not:A-Brand";v="8", "Chromium";v="123"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": "Linux" 
            }

            payload = f"isDirty=false&isReload=false&csrfToken={csrf_token}&PROJECT_ID=1342180104&SEARCH_ACTION=0&searchPreferenceValue=true&searchMode=1&pageNo={page_num}&newDocSearchPage=true&tab=0&sortDir=&drawing=&selectAllIds=false&deselectAllIds=false&selectPageIds=false&attachMode=NONE&EDIT_ACTION=&Correspondence_correspondenceTypeID=&SELECT_ALL_FOR_AUTO_UPDATE=false&COMMAND=&savedSearchId=null&savedSearchType=&LIST_ACTION=&SESSION_SAVE_SELECTED_TAB=0&REGISTER_ACTION=&SRC_PAGE=&SESSION_SAVE_SELECTED_TAB=&PRINT_ACTION=&USE_SELECTED=false&ACTION=&ACTIVE_FOLDER_ID=&ITEM_TYPE=&RETURN_URL=&ITEM_ID=&BROWSER=&VERSION=&CORR_EDIT_SCREENMODEL_ID=null&selectedDocumentId=&numSelectedOnOtherPage=0&numTotal=22065&newSearchEnabled=true&docno=&title=+&doctype=&discipline=&docstatus=&revision=+&author=+&reference=+&vdrcode=&category=&selectList1=&selectList2=&selectList3=&selectList4=&selectList5=&selectList6=&selectList7=&selectList8=&selectList9=&bidi_reviewstatus_regexp=&_bidi_reviewstatus=&-1=-1&-11=&-12=&searchQuery=&projectFields=%5B%5D&sortBy=docno&resultSize=10000&savedSearchName=&savedSearchDescription=&savedSearchSharedFlag=none&ssEditId=-1&ssEditName=&ssEditDescription=%24fld.value&_selectedRegColumns=&_selectedUnregColumns=&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192576188502&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192466743261&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192506101640&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192535975851&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192568976503&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192568976504&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192436230351&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192438926771&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192439091295&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192439099561&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192436193423&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192463838300&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192499248093&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192500152214&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192522751868&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192522752663&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192522752339&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192522752474&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192495726014&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192566981093&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192567012715&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192567006386&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192567006452&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192567006469&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192567006491&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192567006543&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192567006569&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192567006647&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192567006715&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192567006786&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192567006896&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192566344774&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192564634311&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192486941110&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192527503822&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192566397041&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192562774503&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192562774540&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192562774619&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192562774624&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192562774638&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192562774666&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192562774691&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192562774706&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192562774769&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192562774835&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192562774908&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192562774984&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192555244847&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192555245041&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192555245125&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192555245157&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192555245191&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192555245236&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192555245279&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192555245360&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192555245389&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192555245479&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192464243193&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192475544595&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192471363286&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192561453451&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192565078638&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192576362220&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192475544607&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192576362234&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192567051821&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192463863714&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192463864012&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192463864367&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192463864699&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192463865625&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192463866177&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192463866480&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192463866770&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192471740290&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192504872351&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192510688396&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192514650083&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192514650862&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192519274513&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192514649381&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192514649408&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192514649440&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192514649471&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192514649496&HAS_FILE=&IS_SHARED=&allIdsInPage=1302666192514649518&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192459713360&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192466617510&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192475914343&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192476598249&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192476947331&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192479005622&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192481689511&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192482841225&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192485686202&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192486557450&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192488830930&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192489249717&HAS_FILE=yes&IS_SHARED=&allIdsInPage=1302666192490131884"

            response = requests.request("POST", reqUrl, data=payload,  headers=headersList)
            data = response.json()['searchResults']
            total = response.json()['searchResultSummary']['totalResults']

            all_results.extend(data)
                        
            if len(all_results) >= total:
                break
            page_num += 1

        return all_results

    @staticmethod
    def write_docs(data, folder_path, filename):
        column = {
            'Packages': 'numberOfPackages',
            'Project': 'selectList2',
            'Document No.': 'docno',
            'Revision': 'revision',
            'Version': 'version',
            'Title': 'title',
            'Type': 'doctype',
            'Status': 'docstatus',
            'Discipline': 'discipline',
            'Portfolio': 'vdrcode',
            'Created By': 'author',
            'Revision Date': 'revisiondate',
            'Date Modified': 'registered',
            'Asset Class': 'selectList4',
            'Asset/Building': 'selectList3',
            'Comments': 'comments',
            'Confidential': 'confidential',
            'Current': 'cversion',
            'Date Created': 'received',
            'District': 'selectList1',
            'File Name': 'filename',
            'Hyperlink': 'dochyperlink',
            'Level(Drawing/DDE)': 'selectList7',
            'Originator Organisation': 'selectList9',
            'Package/Contrsct Name': 'selectList8',
            'Stage': 'selectList5',
            'Sub-Portfolio': 'category',
            'Type(Drawing/DDE)': 'transmittalin',
            'Transmited in': 'transmitted',
            'Transmitted': 'selectList6',
            'WBS': 'attribute4',
            'Dock Id': 'dockId'
        }

        docs_data = []
        for project in data:
            inner_data = project['values']
            doc = {}
            for key, value in column.items():
                doc[key] = inner_data.get(value, '')
            docs_data.append(doc)

        df = pd.DataFrame(docs_data)
        
        df = df.map(lambda x: x.encode('unicode_escape').decode('utf-8') if isinstance(x, str) else x)

        file_path = os.path.join(folder_path, filename)
        
        try:
            # Attempt to write to Excel file
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            print("Documents Exported Successfully!")
        except openpyxl.utils.exceptions.IllegalCharacterError as e:
            print("Error:", e)
