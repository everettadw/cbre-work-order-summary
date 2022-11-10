# === GMAIL ===
# import smtplib
# import ssl


# PORT = 465
# PASSWORD = "fwmfbagpzsgyfuxi"

# context = ssl.create_default_context()


# def main():
#     with smtplib.SMTP_SSL("smtp.gmail.com", PORT, context=context) as server:
#         server.login("business.eadw@gmail.com", PASSWORD)
#         server.sendmail("business.eadw@gmail.com", "everettdw13@gmail.com",
#                         """Subject: Testing purposes\n\nThis should be the body.""")


# if __name__ == "__main__":
#     main()

# === OUTLOOK ===
import win32com.client as win32


class OutlookClient:
    def __init__(self):
        self.outlook = win32.Dispatch('outlook.application')

    def send_mail(self, to, subject, body=None, html_body=None):
        mail = self.outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject

        if html_body is not None:
            mail.HTMLBody = html_body
        elif body is not None:
            mail.Body = body

        if html_body is None and body is None:
            return False

        mail.Send()
