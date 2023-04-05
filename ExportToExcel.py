from googleapiclient.http import MediaFileUpload

import openpyxl
import csv
from datetime import datetime
import requests
import json
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
from PolicyChangeEncoder import PolicyChangeEncoder
from google.oauth2.service_account import Credentials
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build

to = ['df-rtps-emailsecurityteam@randstad.com']
cc = []
emailFrom = 'dmarc-no-reply@randstad.com'
report_date = datetime.now().strftime('%A, %d. %B %Y %R')
subject = f"Dmarc Data Ingestion - OpCo, SPF & DMARC Lookup: {report_date}"
body = None

SPREADSHEET_ID = '1RWhIYQTnqLGteYgmfrPZkbgZ8SUq_nD4iuBcpXINxyE'

# Replace with the path to your service account JSON key file
SERVICE_ACCOUNT_FILE = 'testprojectdmarc-034ca8acbb1f.json'

# Authenticate and build the service object
creds = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
)

service = build('drive', 'v3', credentials=creds)


class ExportToExcel:
    def export(self, list, Dmarc_Policy_Change):
        file_path = "output.xlsx"
        workbook = openpyxl.load_workbook(filename=file_path)
        sheet = workbook['Sheet']
        rows = sheet.max_row
        # print(f'{rows} is the DMARC Policy')
        for r in range(0, rows - 1):
            row = sheet[r + 2]

            for c in range(5, 9):
                cell = row[c]
                if c == 6:
                    cell.value = list[r].spf_policy
                elif c == 7:
                    cell.value = list[r].dmarc_policy
                elif c == 8:
                    cell.value = list[r].spf_count

        try:
            workbook.save(filename=file_path)
            file_path = file_path
            workbook.save(filename=file_path)
            print("output.xlsx has been created successfully")
            file_metadata = {'name': 'output'}
            media = MediaFileUpload('output.xlsx',
                                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                    resumable=True)

            updated_file = service.files().update(fileId=SPREADSHEET_ID, body=file_metadata, media_body=media,
                                                  fields='id').execute()
            print('Modified Google Sheet uploaded to Google Drive')
            wb = openpyxl.load_workbook('output.xlsx')

            # Select the active worksheet
            ws = wb.active

            # Get the header row
            header = [cell.value for cell in ws[1]]

            # Get the data rows
            data = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                data.append(dict(zip(header, row)))

            # Convert the data to JSON format
            json_data = json.dumps(data, indent=4)

            # Write the JSON data to a file
            with open('Domain_info.json', 'w') as f:
                f.write(json_data)

            print("JSON data saved to Domain_info.json")
            now = datetime.now()

            current_time = now.strftime("%H:%M:%S")
            print("End Time =", current_time)

            # uri = "https://http-inputs-randstad.splunkcloud.com/services/collector"
            # # Enter the original token ->35B8DBEA-2E3C-41DF-A318-EE46154B7AEA
            # splunk_token = "35B8DBEA-2E3C-41DF-A318-EE46154B7AEA"
            # out_path = "Domain_info.json"
            # host_name = "dmarc.orgV2"
            # src = "dmarc:lookupV2"
            #
            # with open(out_path, "r") as f:
            #     input_objects = json.load(f)
            #
            # # for input_object in input_objects:
            # #     body = json.dumps({"event": input_object, "host": host_name, "source": src})
            # #     # print("Body:", body)
            # #     headers = {"Authorization": f"Splunk {splunk_token}"}
            # #     response = requests.post(uri, data=body, headers=headers)
            # #
            # #     if response.text != "Success":
            # #         print(response.text)
            #
            #
            # #  Code foe finding the previous day data with present day data. It will only be possible if we store the output excel sheet, otherwise not possible
            # # Finding what all change have happened previous day+
            # # Total_Q = 0
            # # Total_R = 0
            # # Total_downgrade = 0
            # # Change_spf = 0
            # # transformation_change = 0
            #
            # # for a in range(0, len(Dmarc_Policy_Change)):
            # #     if Dmarc_Policy_Change[a].dmarc_policy_old == "none" and Dmarc_Policy_Change[
            # #         a].dmarc_policy_new == "quarantine":
            # #         Total_Q += 1
            # #     if Dmarc_Policy_Change[a].dmarc_policy_old == "none" and Dmarc_Policy_Change[
            # #         a].dmarc_policy_new == "reject":
            # #         Total_R += 1
            # #     if Dmarc_Policy_Change[a].dmarc_policy_old == "quarantine" and Dmarc_Policy_Change[
            # #         a].dmarc_policy_new == "reject":
            # #         Total_R += 1
            # #     if (Dmarc_Policy_Change[a].dmarc_policy_old == "reject" and Dmarc_Policy_Change[
            # #         a].dmarc_policy_new == "quarantine") \
            # #             or (Dmarc_Policy_Change[a].dmarc_policy_old == "reject" and Dmarc_Policy_Change[
            # #         a].dmarc_policy_new == "none") or (
            # #             Dmarc_Policy_Change[a].dmarc_policy_old == "quarantine" and Dmarc_Policy_Change[
            # #         a].dmarc_policy_new == "none"):
            # #         Total_downgrade += 1
            # #
            # #     if (Dmarc_Policy_Change[a].spf_policy_change_old == None and Dmarc_Policy_Change[
            # #         a].spf_policy_change_new != None) or (
            # #             Dmarc_Policy_Change[a].spf_policy_change_old != Dmarc_Policy_Change[a].spf_policy_change_new):
            # #         Change_spf += 1
            #
            # # print(f'change in Reject - > {Total_R}')
            # # print(f'Change in Qurantine -> {Total_Q}')
            # # print(f'Change in SPF - > {Change_spf}')
            # # json_string = json.dumps(Dmarc_Policy_Change, cls=PolicyChangeEncoder)
            # # with open('Change_Details.json', 'w') as f:
            # #     # Write the JSON string to the file
            # #     f.write(json_string)
            #
            # # create email message
            # # body = "The attached JSON file that includes the details of OpCo, SPF, and DMARC policy has been successfully ingested into Splunk."
            # # msg = MIMEMultipart()
            # # msg['To'] = ', '.join(to)
            # # msg['Cc'] = ', '.join(cc)
            # # msg['Subject'] = subject
            # # msg['From'] = emailFrom
            # # msg.attach(MIMEText(body))
            # #
            # # # with open(file_path2, 'rb') as f:
            # # #     attachment = MIMEBase('application', 'octet-stream')
            # # #     attachment.set_payload(f.read())
            # # #     encoders.encode_base64(attachment)
            # # #     attachment.add_header('Content-Disposition', 'attachment', filename=filename)
            # # #     msg.attach(attachment)
            # # with open('Domain_info.json', 'rb') as f:
            # #     attachment = MIMEBase('application', 'octet-stream')
            # #     attachment.set_payload(f.read())
            # #     encoders.encode_base64(attachment)
            # #     attachment.add_header('Content-Disposition', 'attachment', filename='Total Domain Information.json')
            # #     msg.attach(attachment)
            # #     try:
            # #         server = smtplib.SMTP('smtpeu.randstad.gis', 25)
            # #         server.ehlo()
            # #         server.starttls()
            # #         # not req
            # #         # server.login(username, password)
            # #         server.sendmail(emailFrom, to + cc, msg.as_string())
            # #         server.close()
            # #         print("Email sent successfully")
            # #     except Exception as e:
            # #         print(f'Error: {str(e)}')
            #
            # smtpServer = 'smtpeuext.randstadgis.com'
            # smtpServerPort = 25
            # emailTo = 'df-rtps-emailsecurityteam@randstad.com'
            # emailFrom = 'dmarc-no-reply@randstad.com'
            # reportDate = datetime.now().strftime("%A, %d. %B %Y %R")
            # emailSubject = f"Dmarc Data Ingestion - OpCo, SPF & DMARC Lookup: {reportDate}"
            # emailMessage = MIMEMultipart()
            # emailMessage['From'] = emailFrom
            # emailMessage['To'] = emailTo
            # emailMessage['Subject'] = emailSubject
            # emailMessage.attach(MIMEText(
            #     "The attached JSON file that includes the details of OpCo, SPF, and DMARC policy has been successfully ingested into Splunk."))
            #
            # out = "Domain_info.json"
            # attachment = MIMEApplication(open(out, "rb").read(), _subtype="json")
            # attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(out))
            # emailMessage.attach(attachment)
            # try:
            #     smtpClient = smtplib.SMTP(smtpServer, smtpServerPort)
            #     # smtpClient.starttls()
            #     smtpClient.sendmail(emailFrom, emailTo, emailMessage.as_string())
            #     smtpClient.quit()
            # except Exception as e:
            #     print(f'Error Mail Sending:- {str(e)}')


        except Exception as e:
            print(e)
        finally:
            workbook.close()
