# 🏗️ Aconex Data Extraction Pipeline

This Python project automates the extraction, transformation, 
and loading (ETL) of project data from Aconex into structured formats like CSV and Excel, 
enabling analysts and project teams to efficiently analyze project information.

### This tool automates the extraction process by:
- Logging in securely using Selenium to collect session cookies and CSRF tokens.
- Fetching emails, documents, and workflows via Aconex endpoints.
- Transforming raw data into structured formats (CSV/Excel) with consistent columns and metadata.
- Exporting datasets for further analysis in BI tools or Excel.

### The project follows the ETL architecture:
- Extract: Collect emails, documents, and workflow data from Aconex using API calls and web scraping.
- Transform: Map JSON/HTML data to structured fields, clean strings, and handle missing values.
- Load: Save datasets to CSV or Excel for easy analysis and reporting.

### Key Features
- Automated Login: Uses Selenium to handle Aconex login securely.
- Email Extraction: Retrieves emails, attachments, recipients, status, and confidentiality flags.
- Document Extraction: Captures document metadata (title, revision, status, author, portfolio) and maps it to Excel columns.
- Workflow Extraction: Collects workflow information including step name, action, assignee, dates, and completion status.
- Multi-Page Handling: Automatically loops through pages to gather large datasets.
- ETL Compliance: Structured pipeline ensures data is collected, transformed, and loaded efficiently.<hr><hr>
- Error Handling: Detects HTML or JSON structure changes and reports issues for troubleshooting.

aconex_scraper/
├─ emails.py        # Handles login and email extraction
├─ doc.py           # Handles document extraction
├─ workflows.py     # Handles workflow extraction
├─ main.py          # Main script to execute all extractions
├─ README.md        # Project documentation

### Usage:
- Set credentials and paths in main.py:
  `user_name = "your_username" `
  `password = "your_password"`
  `folder_path = "path_to_save_files"`
  `emails_file_name = "emails.csv"`
  `docs_file_name = "docs.xlsx"`

- Run the script:
  `python main.py`


### Output:
- Emails CSV: Includes Mail No, Subject, Date, From, Recipient, Attachments, Confidential, etc.
- Documents Excel: Columns include Document No, Revision, Title, Type, Status, Author, Portfolio, Hyperlink, etc.
- Workflow CSV: Workflow No, Workflow Name, Document No, Step Name, Action, Assigned To, Dates, Step Status, etc.

### How It Works:
- Login via Selenium
    Automates username and password input.
    Collects cookies and CSRF tokens for API requests.
- Fetch Emails
    Sends GET requests using cookies and CSRF token.
    Loops through all pages to retrieve email metadata and attachments info.
- Fetch Documents
    Sends POST requests to the document search API.
    Loops through all pages, maps JSON data to Excel columns.
- Fetch Workflow Data
    Sends POST requests to workflow API.
    Parses HTML tables using BeautifulSoup.
    Extracts workflow steps, assignees, actions, and completion dates.

### Data Loading
  Emails → CSV
  Documents → Excel
  Workflows → CSV
