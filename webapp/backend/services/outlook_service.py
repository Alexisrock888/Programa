import win32com.client
import os
from datetime import datetime, timedelta
from .file_service import clean_filename


def get_sent_items_folder():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    return ns.GetDefaultFolder(5), outlook


def find_email(sent_folder, recipient: str, subject: str, target_date, target_time, tolerance_minutes: int = 1):
    recipient = recipient.replace("'", "").strip()

    if isinstance(target_date, datetime):
        target_date = target_date.date()
    if isinstance(target_time, datetime):
        target_time = target_time.time()

    target_dt = datetime.combine(target_date, target_time)
    time_from = target_dt - timedelta(minutes=tolerance_minutes)
    time_to = target_dt + timedelta(minutes=tolerance_minutes)

    for mail in sent_folder.Items:
        try:
            if mail.Class != 43:
                continue

            sent_on = mail.SentOn

            if sent_on.date() != target_date:
                continue

            if not (time_from.time() <= sent_on.time() <= time_to.time()):
                continue

            if recipient.lower() not in mail.To.lower():
                continue

            if mail.Subject != subject:
                continue

            return mail

        except Exception:
            continue

    return None


def save_email_as_pdf(mail, output_folder: str, filename: str) -> tuple:
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        clean_name = clean_filename(filename)
        mht_path = os.path.join(output_folder, f"{clean_name}.mht")
        pdf_path = os.path.join(output_folder, f"{clean_name}.pdf")

        counter = 1
        while os.path.exists(mht_path):
            mht_path = os.path.join(output_folder, f"{clean_name}_{counter}.mht")
            pdf_path = os.path.join(output_folder, f"{clean_name}_{counter}.pdf")
            counter += 1

        mail.SaveAs(mht_path, 10)

        doc = word.Documents.Open(mht_path, False, True)
        doc.ExportAsFixedFormat(pdf_path, 17)
        doc.Close(False)
        word.Quit()

        if os.path.exists(mht_path):
            os.remove(mht_path)

        if os.path.exists(pdf_path):
            return True, pdf_path, None
        else:
            return False, None, "No se pudo generar el PDF"

    except Exception as e:
        try:
            word.Quit()
        except:
            pass
        return False, None, str(e)


def send_email(to: str, subject: str, html_body: str, attachments: list) -> tuple:
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = to
        mail.Subject = subject
        mail.HTMLBody = html_body

        for attachment_path in attachments:
            if os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
            else:
                return False, f"Adjunto no encontrado: {attachment_path}"

        mail.Send()
        return True, None

    except Exception as e:
        return False, str(e)
