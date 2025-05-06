import datetime
import webbrowser
import win32com.client
import pythoncom
import psutil


def send_email(subject: str, body: str):

    if not is_old_outlook_running():
        log_to_file(
            "Old Outlook is not running - skipping COM. Using Fallback function for sending email"
        )
        fallback_send_email(subject, body)
        return

    try:
        pythoncom.CoInitialize()  # Ensuring that the COM thread is initialized

        outlook = win32com.client.Dispatch("Outlook.Application")

        # Checking if session is available
        if outlook.Session.Accounts.Count == 0:
            raise RuntimeError("No Outlook account is configured")

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

    try:
        to = "<your.email.name>@outlook.com"  # Change this to your email info
        subject_encoded = urllib.parse.quote(subject)
        body_encoded = urllib.parse.quote(body)
        mailto_url = f"mailto:{to}?subject={subject_encoded}&body={body_encoded}"
        webbrowser.open(mailto_url)
        log_to_file("Fallback: Opened default mail client with mailto link.")
    except Exception as e:
        log_to_file(f"Failed to open mailto link: {e}")


def log_to_file(message: str):
    LOG_FILE = "automation_log.txt"
    with open(LOG_FILE, "a", encoding="utf-8") as log:
        timeStamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log.write(f"[{timeStamp}] {message}\n")


def is_old_outlook_running():
    return any("outlook.exe" in proc.name().lower() for proc in psutil.process_iter())
