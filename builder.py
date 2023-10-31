from sqlalchemy import *
import pandas as pd
from urllib.parse import quote
# from datetime import datetime, timedelta
from openpyxl import *
import win32com.client # pywin32
import os

class DISB_TRACKING:
    def __init__(self):
        self.cnx_analytics = 'mysql+pymysql://analytics:%s@34.93.197.76:61306/analytics' % quote('Secure@123')
        #self.cnx_analytics = 'mysql+pymysql://prakesh:%s@10.10.0.33:61306/analytics' % quote('Prakesh_033')

    def get_data_analytics(self, query):
        cnx = create_engine(self.cnx_analytics)
        data = pd.read_sql(text(query), cnx, index_col=None)
        cnx.dispose()
        return data

    def add_total_row(self, df, columnname_for_total):
        total_row = df.iloc[:, 1:].sum(axis=0)
        # Convert the total row to a DataFrame and transpose it
        total_df = pd.DataFrame(total_row).T
        total_df[f'{columnname_for_total}'] = 'Total'
        # Append the total row to the original DataFrame
        df_with_total = df._append(total_df, ignore_index=True)
        return df_with_total

    def export_data_to_excel(self, raw_data, dte):
        wb = load_workbook('template.xlsx')
        writer = pd.ExcelWriter(f'JLG_topup_Pre_Approved_collection_tracking_Report_as_on_{dte}.xlsx', engine='openpyxl')
        writer.book = wb
        #writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)
        # summary_data.to_excel(writer, sheet_name='Summary', header=False, index=False, startcol=0, startrow=2)
        raw_data.to_excel(writer, sheet_name='Raw_data', header=True, index=False, startcol=0, startrow=0)
        writer.save()


    def email(self, filename, dte):
        # e-mail
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        #mail.To = 'prakash.r@dvarakgfs.com; prakash.r@dvarakgfs.com'

        mail.To = 'Radha.kumari@dvarakgfs.com; anoop.p@dvarakgfs.com; Analytics_kgfs@dvarakgfs.com'
        # mail.CC = 'Analytics_kgfs@dvarakgfs.com'
        # ----
        mail.Subject = f'JLG Topup Pre Approved disb accounts collection tracking details as on  {dte}'


        html_body = f"""
        <html>
        <head>
        </head>

        <body>
           <p>Dear All,</p>
           <p>PFA the JLG Topup Pre Approved Loan, sep month disbursed accounts collection tracking details as of {dte}. </p>
           
           <br><br>
           <p>Regards,</p>
           <p>Analytics Team.,</p>
        </body>
        </html>
        """


        mail.HTMLBody = html_body

        cwd = os.getcwd()
        mail.Attachments.Add(f'{cwd}\{filename}.xlsx')
        #mail.Save()
        mail.Send()
        print('send email.....')

    def dtype_data_correction(self, result):
        result['Demand_date_topup'].fillna('NA', inplace=True)
        result['Demand_Amount_topup'].fillna(0, inplace=True)
        result['amount_paid_topup'].fillna(0, inplace=True)

        result['Demand_date_JLG'].fillna('NA', inplace=True)
        result['Demand_Amount_JLG'].fillna(0, inplace=True)
        result['amount_paid_JLG'].fillna(0, inplace=True)
        return result

    def categorize_status_topup(self, row):
        demand_amount = row['Demand_Amount_topup']
        amount_paid = row['amount_paid_topup']
        if demand_amount == 0:
            return "No Demand"
        elif amount_paid == 0:
            return "Not Paid"
        elif amount_paid < demand_amount:
            return "Partially Paid"
        else:
            return "Paid"

    def categorize_status_jlg(self, row):
        demand_amount = row['Demand_Amount_JLG']
        amount_paid = row['amount_paid_JLG']
        if demand_amount == 0:
            return "No Demand"
        elif amount_paid == 0:
            return "Not Paid"
        elif amount_paid < demand_amount:
            return "Partially Paid"
        else:
            return "Paid"

    def demand_generated_or_not(self, row):
        JLG_Topup_status = row['JLG_Topup_status']
        if JLG_Topup_status == "No Demand":
            return 0
        else:
            return 1







