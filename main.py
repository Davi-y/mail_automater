# main.py
import json
import time
import schedule
import win32com.client as win32
from tkinter import Tk, filedialog, messagebox

# --- Load config ---
with open('config.json', 'r', encoding='utf-8') as f:
    cfg = json.load(f)

TO = cfg.get('to', [])
CC = cfg.get('cc', [])
SUBJECT = cfg.get('subject', 'Daily Report')
BODY = cfg.get('body', '')
SCHEDULE_TIME = cfg.get('schedule_time', '09:00')

# --- Helpers for dialogs ---
def ask_yes_no(title, message):
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    answer = messagebox.askyesno(title, message)
    root.destroy()
    return answer

def pick_file():
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_path = filedialog.askopenfilename(title='Select report to attach')
    root.destroy()
    return file_path

# --- Send mail via Outlook ---
def send_outlook_email(attachment_path=None):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)  # 0: olMailItem
        if TO:
            mail.To = ';'.join(TO)
        if CC:
            mail.CC = ';'.join(CC)
        mail.Subject = SUBJECT
        mail.Body = BODY
        if attachment_path:
            mail.Attachments.Add(attachment_path)
        mail.Send()
        print('Email sent.')
        return True
    except Exception as e:
        print('Error sending email:', e)
        return False

# --- Job run at scheduled time ---
def job():
    print('Reminder popped up...')
    want = ask_yes_no('Send daily report?', f"Do you want to send: \"{SUBJECT}\" now?")
    if not want:
        print('User chose not to send.')
        return

    file_path = pick_file()
    if not file_path:
        print('No file selected; aborting send.')
        return

    success = send_outlook_email(file_path)
    if success:
        ask_yes_no('Done', 'Email sent successfully.')
    else:
        ask_yes_no('Failed', 'Sending failed. Check the console for error details.')

# --- Scheduling ---
schedule.every().day.at(SCHEDULE_TIME).do(job)

if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1 and sys.argv[1].lower() in ('sendnow', 'now'):
        # quick test: trigger the job immediately
        job()
        sys.exit(0)

    print(f"Scheduled daily reminder at {SCHEDULE_TIME}.")
    print("To test immediately run: python main.py sendnow")
    try:
        while True:
            schedule.run_pending()
            time.sleep(1)
    except KeyboardInterrupt:
        print('Stopped by user.')
