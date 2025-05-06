import datetime
import webbrowser
import win32com.client


def send_email(subject: str, body: str):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = "<your.email.name>@outlook.com"  # Change this to your email info
        mail.Subject = subject
        mail.Body = body
        mail.Send()
    except Exception as e:
        log_to_file(f"Outlook COM has failed, using fallback method. Error: {e}\n")
        fallback_send_email(subject, body)


def fallback_send_email(subject: str, body: str):
    import urllib.parse

    to = "<your.email.name>@outlook.com"  # Change this to your email info
    subject_encoded = urllib.parse.quote(subject)
    body_encoded = urllib.parse.quote(body)
    mailto_url = f"mailto:{to}?subject={subject_encoded}&body={body_encoded}"
    webbrowser.open(mailto_url)


def log_to_file(message: str):
    LOG_FILE = "automation_log.txt"
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        timeStamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log.write(f"[{timeStamp}] {message}\n")
