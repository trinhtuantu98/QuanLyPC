import os 
import sys
sys.path.append(r'P:\7. User\Tu')
import pandas as pd 
import numpy as np 
from datetime import datetime ,timedelta
from db import Database_Main
from db_SQL import DB_BCSX_INTERNET
import SendMail

user = 'nltt@nldc.evn.vn'         # Email userID
password = 'Abc@1234567'      # Email password
from_addr = 'nltt@nldc.evn.vn'


df_PC = pd.read_excel("Danh sach Cty DL.xlsx")
df_1 = df_PC[['PC_ID','Ten_PC']]
df_1.set_index('PC_ID', inplace= True)
dict_pc = df_1.to_dict()
df_PC = df_PC[~df_PC.TCT_TRUCTHUOC.isnull()]
listTCT = df_PC['TCT_TRUCTHUOC'].unique().tolist()

dict_cols = {"PTAI_DNGUON" :"Phụ tải đầu nguồn",	
"CS_INTERVER_MD" : "Công suất inverter DMT mặt đất",
"CS_CATGIAM_MD" : "Công suất DMT mặt đất cắt giảm",
"CS_INTERVER_TA" : "Công suất inverter trung áp",
"CS_BAN_TA" : "Công suất bán trung áp",
"CS_CATGIAM_TA" : "Công suất cắt giảm trung áp",
"CS_INTERVER_HA" : "Công suất inverter hạ áp",
"CS_BAN_HA" : "Công suất bán hạ áp",
"CS_CATGIAM_HA" : "Công suất cắt giảm hạ áp"}


def main(target_TCT):
    global dict_pc
    global listTCT
    # tao ra check list 
    ngaybd = datetime.today()
    ngaykt = ngaybd - timedelta(days= 1)
    db = DB_BCSX_INTERNET()
    rawData = db.get_BCSX_PC_NGAY(ngaykt.strftime('%Y%m%d'), ngaybd.strftime('%Y%m%d'))
    df_BCSX = pd.DataFrame.from_records(rawData,columns= ['NGAY','PC_ID',"CHU_KY"
                            ,"PTAI_DNGUON"
                            ,"CS_INTERVER_MD"
                            ,"CS_CATGIAM_MD"
                            ,"CS_INTERVER_TA"
                            ,"CS_BAN_TA"
                            ,"CS_CATGIAM_TA"
                            ,"CS_INTERVER_HA"
                            ,"CS_BAN_HA"
                            ,"CS_CATGIAM_HA"])
    list_PC = df_PC[df_PC.TCT_TRUCTHUOC == target_TCT]['PC_ID'].tolist()
    if target_TCT == 'SPC':
        list_PC.append(5)
    df_sub = df_BCSX[df_BCSX['PC_ID'].isin(list_PC)]
    df_sub_1 = df_sub.copy()
    df_sub['Check'] = np.where((~df_sub['PTAI_DNGUON'].isnull()) | (~df_sub['CS_BAN_HA'].isnull()),1,0)
    list_PC_COSO = df_sub[(df_sub.CHU_KY == 24) & (df_sub.Check == 1)]['PC_ID'].tolist()
    # tab check 
    df_check = pd.DataFrame(list_PC, columns = ['DANHSACH_DL'])
    df_check['TCT_DL'] = target_TCT
    df_check['CHECK'] = np.where(df_check['DANHSACH_DL'].isin(list_PC_COSO),'x',None)
    df_check['DANHSACH_DL'] = df_check['DANHSACH_DL'].replace(dict_pc['Ten_PC'])

    # tab BCSX
    needed_cols = df_sub_1.columns[3:]
    df_2 = pd.melt(df_sub_1, id_vars=["NGAY",'CHU_KY','PC_ID'],
                        value_vars=needed_cols,
                        var_name = 'BAOCAO',
                        value_name="GIATRI")
    df_2['GIATRI'] = df_2['GIATRI'].astype(float)
    df_SoLieu = df_2.pivot_table(columns = ['CHU_KY'] , index = ['NGAY','BAOCAO','PC_ID'], values = ['GIATRI'])
    df_SoLieu.reset_index(inplace = True)
    df_SoLieu.columns = ['NGAY','BAOCAO','PC_ID','CK1','CK2','CK3','CK4','CK5','CK6','CK7','CK8','CK9','CK10','CK11','CK12','CK13','CK14','CK15','CK16','CK17','CK18','CK19','CK20','CK21','CK22','CK23','CK24','CK25','CK26','CK27','CK28','CK29','CK30','CK31','CK32','CK33','CK34','CK35','CK36','CK37','CK38','CK39','CK40','CK41','CK42','CK43','CK44','CK45','CK46','CK47','CK48','SANLUONG']
    df_SoLieu['NGAY'] = df_SoLieu['NGAY'].dt.strftime('%Y%m%d')
    df_SoLieu['PC_ID'] = df_SoLieu['PC_ID'].replace(dict_pc['Ten_PC'])
    df_SoLieu['BAOCAO'] = df_SoLieu['BAOCAO'].replace(dict_cols) 

    # ghi ra file excel
    with pd.ExcelWriter(f"{target_TCT}_{(ngaybd- timedelta(days =1)).strftime('%Y%m%d')}.xlsx") as writer:  # doctest: +SKIP
        df_check.to_excel(writer, sheet_name='CHECK',index= False)
        df_SoLieu.to_excel(writer, sheet_name='BCSX',index= False)


def GuiMail():
    ngaybd = datetime.today()
    for i in listTCT :
        print(f'Bat dau {i}')
        main(i)
        SendMail.Send_email(f"F:\Tu\Tu\QuanLyPC\{i}_{(ngaybd - timedelta(days = 1)).strftime('%Y%m%d')}.xlsx",i)
        print(f"Done {i}")

# if __name__ == '__main__':
#     date = datetime.today()
    # for i in range(4):
    #     GuiMail(date - timedelta(days = i))
    # GuiMail()
    # SendMail.Send_email(f"F:\Tu\Tu\QuanLyPC\CPC_{datetime.today().strftime('%Y%m%d')}.xlsx")


# schedule

import schedule
import time 

print("Bat dau chay tu dong gui mail cho don vi PC")
schedule.every().day.at('12:00').do(GuiMail)

while True:
    schedule.run_pending()
    time.sleep(1)