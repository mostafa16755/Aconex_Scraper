from emails import Emails
from doc import Documents
from workflows import WorkFlows

if __name__ == "__main__":
    user_name = ""
    password = ""

    folder_path = ""

    emails_file_name = "emails.csv"
    docs_file_name = "docs.xlsx"

    # Extract Cookies
    cookies_dict, cookies_string, csrf_token = Emails.login(user_name, password)
    
    # Call Emails
    emails = Emails.fetch_emails(cookies_string, csrf_token)
    Emails.write_emails_to_csv(emails, folder_path, emails_file_name)

    # Call Documents
    docs = Documents.fetch_docs(cookies_dict, cookies_string, csrf_token)
    Documents.write_docs(docs, folder_path, docs_file_name)

    # WorkFlow
    workflows = WorkFlows.fetch_workflow_data(cookies_dict)
    WorkFlows.write_to_csv(workflows, folder_path)