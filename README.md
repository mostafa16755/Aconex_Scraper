🏗️ Aconex Data Extraction Pipeline

A Python project to automate the extraction of emails, documents, and workflows from the Aconex project collaboration platform. This tool is designed for project teams and analysts to collect structured project data and export it to Excel or CSV files.

Features

Automates login to Aconex using Selenium.

Fetches emails, including attachments, recipients, and metadata.

Extracts controlled documents with detailed metadata (title, revision, status, portfolio, etc.).

Collects workflow data with step status, assignee, and completion dates.

Exports emails to CSV and documents/workflows to Excel/CSV.

Handles multiple pages and large datasets automatically.

Requirements

Python 3.10+

Chrome browser + ChromeDriver

Libraries:

pip install requests pandas openpyxl selenium beautifulsoup4

File Structure
aconex_scraper/
├─ emails.py        # Handles login and email extraction
├─ doc.py           # Handles document extraction
├─ workflows.py     # Handles workflow extraction
├─ main.py          # Main script to execute all extractions
├─ README.md        # Project documentation

Usage
1. Set up credentials and paths

In main.py, set your Aconex username, password, and export folder path:

user_name = "your_username"
password = "your_password"
folder_path = "path_to_save_files"
emails_file_name = "emails.csv"
docs_file_name = "docs.xlsx"

2. Run the script
python main.py


The script will:

Open a Chrome browser to login (automated by Selenium).

Extract emails and save them to emails.csv.

Extract controlled documents and save them to docs.xlsx.

Extract workflow data and save it to workflow_data.csv.

3. Output

Emails CSV: Fields include Mail No, Subject, Date, From, Recipient, Has Attachments, Confidential, etc.

Documents Excel: Columns include Document No, Revision, Title, Type, Status, Author, Portfolio, Hyperlink, etc.

Workflow CSV: Includes Workflow No, Workflow Name, Document No, Step Name, Action, Assigned To, Dates, Step Status, etc.

How it Works

Login via Selenium

Automates inputting username and password.

Collects cookies and CSRF token for API requests.

Fetch Emails

Sends GET requests using cookies and CSRF token.

Loops through all pages to collect emails.

Fetch Documents

Sends POST requests to the document search endpoint.

Loops through pages to collect all documents.

Maps raw JSON data to structured Excel columns.

Fetch Workflow Data

Sends POST requests to workflow search endpoint.

Parses HTML tables using BeautifulSoup.

Extracts workflow information and exports to CSV.

