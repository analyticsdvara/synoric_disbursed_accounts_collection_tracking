import pandas as pd
from builder import DISB_TRACKING
import datetime
from datetime import timedelta
from dateutil.relativedelta import relativedelta
import warnings
from log import data_log
warnings.filterwarnings("ignore")


# Automation Declaration
# Automation Date and time
frm = (datetime.date.today() - relativedelta(days=1)).strftime('%Y-%m-01')
to = (datetime.date.today() - relativedelta(days=1)).strftime('%Y-%m-%d')

last_day_of_curr_mnth = datetime.date.today().replace(day=1) + relativedelta(day=31)

check = ((datetime.date.today() - relativedelta(days=1)).strftime('%Y-%m-01'))
date_obj = datetime.datetime.strptime(check, '%Y-%m-%d').date()
to_dte = (date_obj - relativedelta(days=1)).strftime('%d%b%Y')
from_dte = (date_obj - relativedelta(days=1)).strftime('01%b%Y')


# Manuall Date Declaration
print(f'frm: {frm}')
print(f'to: {to}')
print(f'to_dte: {to_dte}')
print(f'from_dte: {from_dte}')
print(f'last_day_of_curr_mnth:{last_day_of_curr_mnth}')

d_t = DISB_TRACKING()

disb_data = pd.read_excel('input_disbursed_cases.xlsx', sheet_name='disb_data', usecols=['URN', 'AccountNumber_topup', 'branch', 'mobile_number', 'customer_name',  'Zone', 'Region', 'disb_date'], dtype={'AccountNumber_topup':str, 'URN':str})
print('Loaded Tracking Details')

# identify JLG Acounts
original_jlg_acounts_query=f"""
select URN, AccountNumber as AccountNumber_JLG from perdix_cdr.quick_cbs_loan_dump_{from_dte}to{to_dte}
where product_code<>'K216'
and product like '%jlg%'
and account_status like '%open%'
"""
original_jlg_acounts = d_t.get_data_analytics(original_jlg_acounts_query)
disb_data = pd.merge(disb_data, original_jlg_acounts, how='left', left_on='URN', right_on='URN')

# 2. Demand Data
demand_query = f"""
select ACCOUNT_NO as AccountNumber, min(INSTALLMENT_DATE) as Demand_date,  round(sum(INSTALLMENT_AMOUNT), 0) as Demand_Amount
 from perdix_cdr.fut__{to_dte.upper()}
where INSTALLMENT_DATE between '{frm}' and '{to}'
group by ACCOUNT_NO
"""
demand_data = d_t.get_data_analytics(demand_query)
demand_data['AccountNumber'] = demand_data['AccountNumber'].astype(str)

# Demand for Topup Loans
result = pd.merge(disb_data, demand_data, how='left', left_on='AccountNumber_topup', right_on='AccountNumber')
result.drop('AccountNumber', axis=1, inplace=True)
result = result.rename(columns={'Demand_Amount': 'Demand_Amount_topup', 'Demand_date':'Demand_date_topup'})

print('Demand data fetched for topup loans..')
# Demand for Topup Loans
result = pd.merge(result, demand_data, how='left', left_on='AccountNumber_JLG', right_on='AccountNumber')
result.drop('AccountNumber', axis=1, inplace=True)
result = result.rename(columns={'Demand_Amount': 'Demand_Amount_JLG', 'Demand_date':'Demand_date_JLG'})
print('Demand data fetched for jlg loans..')

# 3. Collection Data
coll_query = f"""
select account_no as AccountNumber, sum(amount_paid) as amount_paid
from assist_cdr.assist_loan_repayments
where date(created_at) between "{frm}" and "{last_day_of_curr_mnth}"
group by account_no
"""
coll_data = d_t.get_data_analytics(coll_query)
coll_data['AccountNumber'] = coll_data['AccountNumber'].astype(str)

# collection JLG Topup
result = pd.merge(result, coll_data, how='left', left_on='AccountNumber_topup', right_on='AccountNumber')
result.drop('AccountNumber', axis=1, inplace=True)
result = result.rename(columns={'amount_paid': 'amount_paid_topup'})
print('Collection data fetched for JLG Topup')

# collection JLG
result = pd.merge(result, coll_data, how='left', left_on='AccountNumber_JLG', right_on='AccountNumber')
result.drop('AccountNumber', axis=1, inplace=True)
result = result.rename(columns={'amount_paid': 'amount_paid_JLG'})
print('Collection data fetched for JLG')

# data correction
result = d_t.dtype_data_correction(result)


result['JLG_Topup_status'] = result.apply(d_t.categorize_status_topup, axis=1)
print('JLG_Topup_status is completed...')

result['JLG_status'] = result.apply(d_t.categorize_status_jlg, axis=1)
print('JLG_status is completed...')

result['jlg_topup_demand_generated'] = result.apply(d_t.demand_generated_or_not, axis=1)

# arrange column
result = result[['URN',	'AccountNumber_topup',	'branch',	'mobile_number',	'customer_name', 'Region', 'disb_date',	'Zone',	'AccountNumber_JLG',	'Demand_date_topup',	'Demand_Amount_topup',	'amount_paid_topup',	'JLG_Topup_status',	'Demand_date_JLG',	'Demand_Amount_JLG',	'amount_paid_JLG',	'JLG_status', 'jlg_topup_demand_generated']]


d_t.export_data_to_excel(result, to)
print('Exported to excel...')

# ----------------

# # mail
filename = f"JLG_topup_Pre_Approved_collection_tracking_Report_as_on_{to}"
d_t.email(filename, to)
#
# # log
data_log("Synoric Disb Collection Tacking")