import EmailNLDC 
from datetime import datetime, timedelta
import os
import time
from datetime import datetime,timedelta 

# server='smtp.office365.com'
user = 'nltt@nldc.evn.vn'         # Email userID
password = 'Abc@1234567'      # Email password
from_addr = 'nltt@nldc.evn.vn'
# recipients_addr = ['trinhtuantu98@gmail.com','hungtv@nldc.evn.vn', 'linhbd@nldc.evn.vn']
recipients_addr = ['tomas.hruska@solargis.com','michal.strba@solargis.com','artur.skoczek@solargis.com','daniele.fucini@solargis.com','trinhtuantu98@gmail.com','hungtv@nldc.evn.vn']
subject = 'EVNNLDC RE forecast daily report'
body_vinhtan = '''Dear Solargis team, please check historiscal data from Vinh Tan solar plant as attachment file.
'''
body_songluy = '''Dear Solargis team, please check historiscal data from Song Luy solar plant as attachment file.
'''
subject_songluy='EVNNLDC historical data for ID_NM 491'
subject_vinhtan='EVNNLDC historical data for ID_NM 395'

path_ls = r"Z:\LUU"

def job():
    ngayluachon = datetime.today() - timedelta(days=1)
    path_common = r"P:\1. Du bao\7. Du lieu DMT Bx\SongLuy-VinhTan"
    vinhtan_subpath = f"\ID_NM_395_{ngayluachon.strftime('%d_%m_%Y')}.csv"
    total_path_vt = path_common + vinhtan_subpath
    songluy_subpath = f"\ID_NM_491_{ngayluachon.strftime('%d_%m_%Y')}.csv"
    total_path_sl = path_common + songluy_subpath     
    EmailNLDC.send_email(user, password, from_addr, recipients_addr, subject=subject_vinhtan, body=body_vinhtan, files_path=total_path_vt)
    EmailNLDC.send_email(user, password, from_addr, recipients_addr, subject=subject_songluy, body=body_songluy, files_path=total_path_sl)
    print('Done send email')
# import schedule
# print('Start checking ... D+2')
# schedule.every().day.at("00:15").do(job)
# schedule.every().day.at("08:10").do(job)


# while True:
#     schedule.run_pending()
#     time.sleep(1) 
job ()