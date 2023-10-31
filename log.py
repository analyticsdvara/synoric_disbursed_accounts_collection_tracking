import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine, text, NVARCHAR, DATE
from urllib.parse import quote
def data_log(report_name):
    data_log = pd.DataFrame({
        "report_name": [f'{report_name}'],
        "published": [datetime.now()],
        "published_date": [datetime.now().date()]})
    sql_string = 'mysql+pymysql://analytics:%s@34.93.197.76:61306/analytics' % quote('Secure@123')
    cnx = create_engine(sql_string)
    data_log.to_sql(con=cnx, name='automation_tracker', if_exists='append', index=False)
    print('Log Created....')

