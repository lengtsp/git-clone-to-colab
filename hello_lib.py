


import datetime, pytz
tz = pytz.timezone('Asia/Bangkok')

import pandas as pd
import numpy as np
def now():
    now1 = datetime.datetime.now(tz)
    month_name = 'x มกราคม กุมภาพันธ์ มีนาคม เมษายน พฤษภาคม มิถุนายน กรกฎาคม สิงหาคม กันยายน ตุลาคม พฤศจิกายน ธันวาคม'.split()[now1.month]
    thai_year = now1.year + 543
    time_str = now1.strftime('%H:%M:%S')
    return "%d %s %d %s"%(now1.day, month_name, thai_year, time_str) # 30 ตุลาคม 2560 20:45:30
  
   

def import_gg_xlsx(file_id, file_name):
    

    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive 
    from google.colab import auth 
    from oauth2client.client import GoogleCredentials
    auth.authenticate_user()
    gauth = GoogleAuth()
    gauth.credentials = GoogleCredentials.get_application_default()
    drive = GoogleDrive(gauth)

    downloaded = drive.CreateFile({'id': file_id})
    downloaded.GetContentFile(file_name)
    df_new          = pd.read_excel(file_name)
    df_new = df_new.fillna("")
    return df_new



def import_gg_csv(file_id, file_name):
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive 
    from google.colab import auth 
    from oauth2client.client import GoogleCredentials
    auth.authenticate_user()
    gauth = GoogleAuth()
    gauth.credentials = GoogleCredentials.get_application_default()
    drive = GoogleDrive(gauth)

    downloaded = drive.CreateFile({'id': file_id})
    downloaded.GetContentFile(file_name)
    df_new = pd.read_csv(file_name ,  sep='\^'       )
    df_new = df_new.fillna("")
    return df_new

def import_gg_csv_comma(file_id, file_name):
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive 
    from google.colab import auth 
    from oauth2client.client import GoogleCredentials
    auth.authenticate_user()
    gauth = GoogleAuth()
    gauth.credentials = GoogleCredentials.get_application_default()
    drive = GoogleDrive(gauth)

    downloaded = drive.CreateFile({'id': file_id})
    downloaded.GetContentFile(file_name)
    df_new = pd.read_csv(file_name ,  sep=','       )
    df_new = df_new.fillna("")
    return df_new


#พ่นออกมาที่ชีทนี้
#https://docs.google.com/spreadsheets/d/1EaVFbZrNsiE9fT_DXrCoSN9PNZmzXte5WgB0SVT72PI/edit#gid=1020847031
def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list)+1)


def add_log(
    link_log,
    cell1,
    cell2,
    cell3,
    cell4,
    cell5,
    cell6

):
    
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive 
    from google.colab import auth 
    from oauth2client.client import GoogleCredentials

    auth.authenticate_user()
    gauth = GoogleAuth()
    gauth.credentials = GoogleCredentials.get_application_default()
    drive = GoogleDrive(gauth)

    #สำหรับจัดการ Google sheet
    import gspread
    from gspread_dataframe import set_with_dataframe
    from gspread_dataframe import get_as_dataframe


    from google.auth import default
    creds, _ = default()

    gc = gspread.authorize(creds)
    # gc = gspread.authorize(GoogleCredentials.get_application_default())
    #สำหรับจัดการ Google sheet---------------------------------


    
    
    
    
    wb_logfile= gc.open_by_url(link_log)
    sheet1 = wb_logfile.worksheet('Logfile Report')
    next_row = next_available_row(sheet1)

    sheet1.update_acell("A{}".format(next_row), cell1)
    sheet1.update_acell("B{}".format(next_row), cell2)
    sheet1.update_acell("C{}".format(next_row), cell3)
    sheet1.update_acell("D{}".format(next_row), cell4)
    sheet1.update_acell("E{}".format(next_row), cell5)
    sheet1.update_acell("F{}".format(next_row), cell6)

    
def df_rename_column(df, col_index, col_name ):
    df         = df.rename(columns={df.columns[col_index] : col_name})      
    return df

def df_cutheader_secondrow(df, col_index, col_name, cut_name ):
    #     df         = df.rename(columns={df.columns[col_index] : col_name})      
    cut_name_list = [cut_name]
    df         = df[  (~df[col_name].isin(cut_name_list))   ]
    return df


# def convert_field_type_int(df, col_name):
#     df[col_name] = df[col_name].astype(int)
#     return df


# def convert_field_type_int(col_value):
#     return str(col_value)

def mapping_base_info_from_main(df_m, df, col_name):
    df['empl_status'] = df[col_name].map(df_m.set_index(col_name)['empl_status'])
    df['empl_class'] = df[col_name].map(df_m.set_index(col_name)['empl_class'])
    df['descr_c'] = df[col_name].map(df_m.set_index(col_name)['descr_c'])
    df['descr_m'] = df[col_name].map(df_m.set_index(col_name)['descr_m'])
    df['effdt_en'] = df[col_name].map(df_m.set_index(col_name)['effdt_en'])
    df['effdt_en_yyyymm'] = df[col_name].map(df_m.set_index(col_name)['effdt_en_yyyymm'])
    df['merge_hiredate_en'] = df[col_name].map(df_m.set_index(col_name)['merge_hiredate_en'])
    df['merge_hiredate_en_yyyymm'] = df[col_name].map(df_m.set_index(col_name)['merge_hiredate_en_yyyymm'])
    return df



def mapping_ey_from_main(df_m, df, col_name, today_date):
    
    df_m['today_date'] = today_date

    df_m['พนักงานเข้าใหม่ยังไม่ถึงปี'] = ''
    df_m.loc[(df_m['empl_status'].isin(['A'])  ) &
                 (pd.to_datetime(df_m['today_date'], format='%d/%m/%Y') - pd.to_datetime(df_m['hire_date_en'], format='%d/%m/%Y')  < '365 days' ), 'พนักงานเข้าใหม่ยังไม่ถึงปี']  = 'พนักงานเข้าใหม่ยังไม่ถึงปี'

    df_m.loc[(df_m['action'] == 'REH') , 'พนักงานเข้าใหม่ยังไม่ถึงปี']  = ''


#     df_m['วันสิ้นสภาพ'] = ''
    df_m.loc[(df_m['empl_status'].isin(['T','D'])  ), 'วันสิ้นสภาพ']  = df_m['effdt_en']
    
    
    
    
    df['rc_code']                         = df[col_name].map(df_m.set_index(col_name)['rc_code'] )
    df['descr_rc_code']                   = df[col_name].map(df_m.set_index(col_name)['descr_rc_code'] )
    df['descr_c']                         = df[col_name].map(df_m.set_index(col_name)['descr_c'] )
    df['1. กำกับสายงาน']                   = df[col_name].map(df_m.set_index(col_name)['1. กำกับสายงาน'] )
    df['2. สาย']                          = df[col_name].map(df_m.set_index(col_name)['2. สาย'] )
    df['3. กลุ่ม']                          = df[col_name].map(df_m.set_index(col_name)['3. กลุ่ม'] )
    df['merge_hiredate_en']               = df[col_name].map(df_m.set_index(col_name)['merge_hiredate_en'] )
    df['age']                             = df[col_name].map(df_m.set_index(col_name)['age'] )
    df['sex']                             = df[col_name].map(df_m.set_index(col_name)['sex'] )
    df['sys_retireyear']                  = df[col_name].map(df_m.set_index(col_name)['sys_retireyear'] )
    df['empl_class']                      = df[col_name].map(df_m.set_index(col_name)['empl_class'] )
    df['effdt']                           = df[col_name].map(df_m.set_index(col_name)['effdt'] )
    df['action']                          = df[col_name].map(df_m.set_index(col_name)['action'] )
    df['action_reason']                   = df[col_name].map(df_m.set_index(col_name)['action_reason'] )
    df['action_reason_descr']             = df[col_name].map(df_m.set_index(col_name)['action_reason_descr'] )
    df['พนักงานเข้าใหม่ยังไม่ถึงปี']             = df[col_name].map(df_m.set_index(col_name)['พนักงานเข้าใหม่ยังไม่ถึงปี'] )
    df['วันสิ้นสภาพ']                        = df[col_name].map(df_m.set_index(col_name)['วันสิ้นสภาพ'] )
    df['เรียงลำดับสายงาน']                   = df[col_name].map(df_m.set_index(col_name)['เรียงลำดับสายงาน'] )
    df['สายงานภายใต้การรับผิดชอบของ BP ณ ปัจจุบัน (Fix Code กรณีมีการแก้ไขต้องปรับเปลี่ยนช่องด้วย)'] = df[col_name].map(df_m.set_index(col_name)['สายงานภายใต้การรับผิดชอบของ BP ณ ปัจจุบัน (Fix Code กรณีมีการแก้ไขต้องปรับเปลี่ยนช่องด้วย)'] )
    
    
    a = ['CON',  'RLS',  'WHI']
    df = df[~df.action_reason.isin(a)]


    return df




def mapping_ey_from_main_EY(df_m, df, col_name, today_date, field11,field12,field13,    field21,field22,field23,    field31,field32,field33  ,  
                            field12_3m, field22_12m  , field32_12m 
                           ):
    
    df_m['today_date'] = today_date

    df_m['พนักงานเข้าใหม่ยังไม่ถึงปี'] = ''
    df_m.loc[(df_m['empl_status'].isin(['A'])  ) &
                 (pd.to_datetime(df_m['today_date'], format='%d/%m/%Y') - pd.to_datetime(df_m['hire_date_en'], format='%d/%m/%Y')  < '365 days' ), 'พนักงานเข้าใหม่ยังไม่ถึงปี']  = 'พนักงานเข้าใหม่ยังไม่ถึงปี'

    df_m.loc[(df_m['action'] == 'REH') , 'พนักงานเข้าใหม่ยังไม่ถึงปี']  = ''


#     df_m['วันสิ้นสภาพ'] = ''
    df_m.loc[(df_m['empl_status'].isin(['T','D'])  ), 'วันสิ้นสภาพ']  = df_m['effdt_en']
    

#     df = df.drop(['สายงาน'],axis =1)
    
    df['รหัสสังกัด']                         = df[col_name].map(df_m.set_index(col_name)['rc_code'] )
    df['สังกัด']                   = df[col_name].map(df_m.set_index(col_name)['descr_rc_code'] )
    df['corporate']                         = df[col_name].map(df_m.set_index(col_name)['descr_c'] )
    df['กำกับสายงาน']                   = df[col_name].map(df_m.set_index(col_name)['1. กำกับสายงาน'] )
    df['สาย']                          = df[col_name].map(df_m.set_index(col_name)['2. สาย'] )
    df['กลุ่ม']                          = df[col_name].map(df_m.set_index(col_name)['3. กลุ่ม'] )
    df['merge_hiredate_en']               = df[col_name].map(df_m.set_index(col_name)['merge_hiredate_en'] )
    df['age']                             = df[col_name].map(df_m.set_index(col_name)['age'] )
    df['sex']                             = df[col_name].map(df_m.set_index(col_name)['sex'] )
    df['ปีเกษียณ ค.ศ.']                  = df[col_name].map(df_m.set_index(col_name)['sys_retireyear'] )
    df['empl_class']                      = df[col_name].map(df_m.set_index(col_name)['empl_class'] )
    df['effdt']                           = df[col_name].map(df_m.set_index(col_name)['effdt'] )
    df['action']                          = df[col_name].map(df_m.set_index(col_name)['action'] )
    df['action_reason']                   = df[col_name].map(df_m.set_index(col_name)['action_reason'] )
    df['action_reason_descr']             = df[col_name].map(df_m.set_index(col_name)['action_reason_descr'] )
    df['พนักงานเข้าใหม่ยังไม่ถึงปี']             = df[col_name].map(df_m.set_index(col_name)['พนักงานเข้าใหม่ยังไม่ถึงปี'] )
    df['วันสิ้นสภาพ']                        = df[col_name].map(df_m.set_index(col_name)['วันสิ้นสภาพ'] )
    df['เรียงลำดับสายงาน']                   = df[col_name].map(df_m.set_index(col_name)['เรียงลำดับสายงาน'] )
    df['หน่วยงานที่ดูแล'] = df[col_name].map(df_m.set_index(col_name)['สายงานภายใต้การรับผิดชอบของ BP ณ ปัจจุบัน (Fix Code กรณีมีการแก้ไขต้องปรับเปลี่ยนช่องด้วย)'] )
    
    
    a = ['CON',  'RLS',  'WHI']
    df = df[~df.action_reason.isin(a)]


#     this_year_str = str(this_year)
#     this_year_before = str(this_year-1)
#     field11 = '1.1 สิทธิลาพักผ่อนยกมาจาก ' + this_year_before + ' ใช้ภายใน 31/03/'+ this_year_str
#     field12 = '1.2 ยอดใช้ไปลาพักผ่อนยกมาจาก ' + this_year_before
#     field13 = '1.3 คงเหลือลาพักผ่อนยกมาจาก ' + this_year_before
    
#     field21 = '2.1 สิทธิลาพักผ่อนประจำปี ' + this_year_str
#     field22 = '2.2 ยอดใช้ไปลาพักผ่อนประจำปี ' + this_year_str
#     field23 = '2.3 คงเหลือลาพักผ่อนปี ' + this_year_str
    
#     field31 = '3.1 สิทธิลาพักผ่อนสะสม'
#     field32 = '3.2 ยอดใช้ไปลาพักผ่อนสะสม'
#     field33 = '3.3 คงเหลือลาพักผ่อนสะสม'
    
    

    df = df.rename(columns={
                            'empl_status' : 'สถานะ',
                            'ABSV_DYS_CARRYOVER' : field11,
                            'KTB_ABS_TAKE_FW' : field12,
                            'KTB_ABS_TAKE_FW_3M' : field12_3m,

                            'ABSV_DYS_EARN_YTD' : field21,
                            'ABSV_DYS_TAKE_YTD' : field22,
                            'ABSV_DYS_TAKE_YTD_12M' : field22_12m,
        
                            'sum ตารางยกยอด 3 ปีจริงที่เริ่มรันตอนสิ้นปี 2564 ABSV_DYS_CARRYOVER' : field31,
                            'sum ตารางยกยอด 3 ปีจริงที่เริ่มรันตอนสิ้นปี 2564 ABSV_DYS_TAKE_YTD' : field32,
                            'sum ตารางยกยอด 3 ปีจริงที่เริ่มรันตอนสิ้นปี 2564 ABSV_DYS_TAKE_YTD_ALL' : field32_12m,
        

    })

    
    df = df[(df['สถานะ'].isin(['A','L','S']))]
    df['ปีเกษียณ ค.ศ.'] = df['ปีเกษียณ ค.ศ.'] - 543
    return df
    
    





def fnc_senddata_to_googlesheet(df , googlesheet_url, sheet_name, column_list, sheet_range):

#     !ls drive
#     !pip install pydrive
#     !pip install xlsxwriter
    import pandas as pd

    #add column
    if (len(column_list)>0):
        df = df[column_list]
        

    # these classes allow you to request the Google drive API
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive 
    from google.colab import auth 
    from oauth2client.client import GoogleCredentials

    auth.authenticate_user()
    gauth = GoogleAuth()
    gauth.credentials = GoogleCredentials.get_application_default()
    drive = GoogleDrive(gauth)

    #สำหรับจัดการ Google sheet
    import gspread
    from gspread_dataframe import set_with_dataframe
    from gspread_dataframe import get_as_dataframe


    from google.auth import default
    creds, _ = default()

    gc = gspread.authorize(creds)
    # gc = gspread.authorize(GoogleCredentials.get_application_default())
    #สำหรับจัดการ Google sheet---------------------------------


    
    
    
    
    #wb2 = gc.open_by_url('https://docs.google.com/spreadsheets/d/{}/edit#gid=0'.format(val_4_2))
    wb2 = gc.open_by_url(googlesheet_url)

    
    
    try:
            sheet1 = wb2.worksheet(sheet_name)
#             wb2.del_worksheet(sheet1)

#             worksheet1 = wb2.add_worksheet(title=sheet_name, rows="100", cols="20")
    except:
            worksheet1 = wb2.add_worksheet(title=sheet_name, rows="100", cols="20")
    
    
    #worksheet1 เสมือน sheet_destination_sheet
    sheet_destination_sheet = wb2.worksheet(sheet_name)
    wb2.values_clear("'{}'!{}".format(sheet_name,sheet_range))
    set_with_dataframe(sheet_destination_sheet, df, row=10, include_column_header=True) 
    
    
    

    
    

def fnc_senddata_to_googlesheet_change_startrow(df , googlesheet_url, sheet_name, column_list, sheet_range, change_startrow):

#     !ls drive
#     !pip install pydrive
#     !pip install xlsxwriter
    import pandas as pd

    #add column
    if (len(column_list)>0):
        df = df[column_list]
        

    # these classes allow you to request the Google drive API
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive 
    from google.colab import auth 
    from oauth2client.client import GoogleCredentials

    auth.authenticate_user()
    gauth = GoogleAuth()
    gauth.credentials = GoogleCredentials.get_application_default()
    drive = GoogleDrive(gauth)

    #สำหรับจัดการ Google sheet
    import gspread
    from gspread_dataframe import set_with_dataframe
    from gspread_dataframe import get_as_dataframe


    from google.auth import default
    creds, _ = default()

    gc = gspread.authorize(creds)
    # gc = gspread.authorize(GoogleCredentials.get_application_default())
    #สำหรับจัดการ Google sheet---------------------------------


    
    
    
    
    #wb2 = gc.open_by_url('https://docs.google.com/spreadsheets/d/{}/edit#gid=0'.format(val_4_2))
    wb2 = gc.open_by_url(googlesheet_url)

    
    
    try:
            sheet1 = wb2.worksheet(sheet_name)
#             wb2.del_worksheet(sheet1)

#             worksheet1 = wb2.add_worksheet(title=sheet_name, rows="100", cols="20")
    except:
            worksheet1 = wb2.add_worksheet(title=sheet_name, rows="100", cols="20")
    
    
    #worksheet1 เสมือน sheet_destination_sheet
    sheet_destination_sheet = wb2.worksheet(sheet_name)
    wb2.values_clear("'{}'!{}".format(sheet_name,sheet_range))
    set_with_dataframe(sheet_destination_sheet, df, row=change_startrow, include_column_header=True) 
    
    
    
    
def grouptake1_12month(df,  filter_type, key_type_value):
#     listofmonth_int = ['01','02','03','04','05','06','07','08','09','10','11','12']
    listofmonth_int = ['1','2','3','4','5','6','7','8','9','10','11','12']
    key_mapping = 'emplid'
    key_sum = 'DURATION_DAYS'
    key_m = 'ABSENCE_DATE_only_month'
    key_type = 'ประเภทการลาพักผ่อน 3 กลุ่ม'
    
    key_type_value = key_type_value.replace('1.2 ','').replace('2.2 ','').replace('3.2 ','')
    
    newlist = []
    count = 1
    # for x in range(1, 10):
    #     field_rename = 'สิทธิลาพักผ่อนประจำปี : รวมวันลาใช้ไป ณ เดือน ' + str(x)
    for x1 in listofmonth_int:
        newlist.append(x1)
        
        field_rename = key_type_value +  ' ณ เดือน ' + x1 #+ ' (หน่วย:วัน)' 

        
#         print(newlist)
#         print(field_rename)

        #step1 group value with month
#         df_group           = df[  (df[key_type]).isin([key_type_value]) & (df[key_m]).isin(newlist)  ] \
#                                     .groupby(key_mapping).agg({key_sum: 'sum'}).reset_index().rename(columns={key_sum:field_rename})
        
        if filter_type == 'ลาพักผ่อนสะสม':
                df_group = df[      (
                                        (df['ประเภทการลาพักผ่อน 3 กลุ่ม'] == filter_type) & 
                                        (df['ABSENCE_DATE_only_month'].isin(newlist))  
                                    )   |
                                        (df['ประเภทการลาพักผ่อน 3 กลุ่ม'] == 'ลาพักผ่อนสะสมปีก่อนหน้าปัจจุบัน') 
                             ] \
                    .groupby('emplid').agg({'DURATION_DAYS': 'sum'}).reset_index().rename(columns={'DURATION_DAYS':field_rename})
                
        else:
                df_group = df[  (df['ประเภทการลาพักผ่อน 3 กลุ่ม'] == filter_type) & (df['ABSENCE_DATE_only_month'].isin(newlist))  ] \
                    .groupby('emplid').agg({'DURATION_DAYS': 'sum'}).reset_index().rename(columns={'DURATION_DAYS':field_rename})
        
        
        
#         print( 'key_type_value ' + key_type_value)
#         print( 'ABSENCE_DATE_only_month ' + newlist)
#         print( 'field_rename ' + field_rename)
        #step2 mapping data
        df[key_mapping] = df[key_mapping].astype(int)
        df_group[key_mapping] = df_group[key_mapping].astype(int)
        df[field_rename]       = df[key_mapping].map(df_group.set_index(key_mapping)[field_rename])

#         print("----------------df group field_rename--------------")
#         print(df_group.head())
        
        #step3 fillna, replace fill
        df.loc[(      df[field_rename].isnull() ), field_rename]  = 0.0
        count                       = count +1
#         print(field_rename)

        
        
        
        
def assign_field_to_float(df, f):
    df[f] = df[f].fillna('0.0')
    df.loc[(df[f] == '' ),f]='0.0'
    df[f] = df[f].astype(float)
    return df



def parameter_this_year(this_year):

    return this_year


def upload_file_split_sheet(df, id_googlesheet, 
                            this_year 
                           
                           ,field11, field12, field13
                           ,field21, field22, field23
                           ,field31, field32, field33,
                            
                            
                            f1_value, f2, f3, f3_value, my_sheet, my_clean_range
                           ):
    

    
    df_new = df[   df['หน่วยงานที่ดูแล'].isin([f1_value])  & df[f2].isin(['A'])  &  df[f3].isin([f3_value])          ]

    
#     this_year_str = str(this_year)
#     this_year_before = str(this_year-1)
#     field11 = '1.1 สิทธิลาพักผ่อนยกมาจาก ' + this_year_before + ' ใช้ภายใน 31/03/'+ this_year_str
#     field12 = '1.2 ยอดใช้ไปลาพักผ่อนยกมาจาก ' + this_year_before
#     field13 = '1.3 คงเหลือลาพักผ่อนยกมาจาก ' + this_year_before
    
#     field21 = '2.1 สิทธิลาพักผ่อนประจำปี ' + this_year_str
#     field22 = '2.2 ยอดใช้ไปลาพักผ่อนประจำปี ' + this_year_str
#     field23 = '2.3 คงเหลือลาพักผ่อนปี ' + this_year_str
    
#     field31 = '3.1 สิทธิลาพักผ่อนสะสม'
#     field32 = '3.2 ยอดใช้ไปลาพักผ่อนสะสม'
#     field33 = '3.3 คงเหลือลาพักผ่อนสะสม'
    
        
    df_new = df_new.rename(columns={    '3. กลุ่ม' : 'กลุ่ม',
                                        'rc_code' : 'รหัสสังกัด',
                                        'descr_rc_code' : 'สังกัด',
                                        '1.1 สิทธิการลาพักผ่อนประจำปียกมาจากปีที่แล้ว' : field11,
                                        '1.2 ยอดใช้ไป - สิทธิการลาพักผ่อนประจำปียกมาจากปีที่แล้ว' : field12,                          
                                        '1.3 คงเหลือ - สิทธิการลาพักผ่อนประจำปียกมาจากปีที่แล้ว' : field13, 
                                    
                                    
                                        '2.1 สิทธิการลาพักผ่อนประจำปีปัจจุบัน' : field21,
                                        '2.2 ยอดใช้ไป - สิทธิการลาพักผ่อนประจำปีปัจจุบัน' : field22,                          
                                        '2.3 คงเหลือ - สิทธิการลาพักผ่อนประจำปีปัจจุบัน' : field23,     
                                    
                                        '3.1 สิทธิการลาพักผ่อนสะสมใช้ได้ 3 ปี' : field31,
                                        '3.2 ยอดใช้ไป - สิทธิการลาพักผ่อนสะสมใช้ได้ 3 ปี' : field32,                          
                                        '3.3 คงเหลือ - สิทธิการลาพักผ่อนสะสมใช้ได้ 3 ปี' : field33,     
                                    
    })


    df_new.sort_values([field33], ascending=[False],  na_position ='first', inplace=True)
    column_BP = [                
                                'emplid',
                                'name',
                                'ตำแหน่ง',
                                'กลุ่ม',
                                'รหัสสังกัด',
                                'สังกัด',
                                'กลุ่มที่ต้อง adjust วันลา',


                                field31,
                                field32,
                                field33,

        
                                field21,
                                field22,
                                field23,

                                field11,
                                field12,
                                field13,
             
        
        
#                                 'หน่วยงานที่ดูแล'
                              ]
    df_new = df_new.drop(['หน่วยงานที่ดูแล'],axis =1)
    fnc_senddata_to_googlesheet(df_new, id_googlesheet, my_sheet  , column_BP, my_clean_range) 
    
    
    
    
def run_e_status(df,e):
    df[e] = 0
    df.loc[(df[e] == 'A' ),e]=1
    df.loc[(df[e] == 'L' ),e]=2
    df.loc[(df[e] == 'S' ),e]=3
    df.loc[(df[e] == 'T' ),e]=4
    df.loc[(df[e] == 'D' ),e]=5
    return df
    
    
def fnc_find_this_year_match_function(today_date_month, my_text):
    my_column_name = ''
    if(today_date_month =='1'):
        my_column_name = my_text + '1'
    elif(today_date_month =='2'):
        my_column_name = my_text + '2'
    elif(today_date_month =='3'):
        my_column_name = my_text + '3'
    elif(today_date_month =='4'):
        my_column_name = my_text + '4'
    elif(today_date_month =='5'):
        my_column_name = my_text + '5'
    elif(today_date_month =='6'):
        my_column_name = my_text + '6'
    elif(today_date_month =='7'):
        my_column_name = my_text + '7'
    elif(today_date_month =='8'):
        my_column_name = my_text + '8'
    elif(today_date_month =='9'):
        my_column_name = my_text + '9'
    elif(today_date_month =='10'):
        my_column_name = my_text + '10'
    elif(today_date_month =='11'):
        my_column_name = my_text + '11'
    elif(today_date_month =='12'):
        my_column_name = my_text + '12'

     
    return my_column_name
    
    
    
def fnc_find_this_year_match_function_before_month(today_date_month, my_text):

    my_column_name = ''
    if(today_date_month =='1'):
        my_column_name = my_text + '1'
    elif(today_date_month =='2'):
        my_column_name = my_text + '1'
    elif(today_date_month =='3'):
        my_column_name = my_text + '2'
    elif(today_date_month =='4'):
        my_column_name = my_text + '3'
    elif(today_date_month =='5'):
        my_column_name = my_text + '4'
    elif(today_date_month =='6'):
        my_column_name = my_text + '5'
    elif(today_date_month =='7'):
        my_column_name = my_text + '6'
    elif(today_date_month =='8'):
        my_column_name = my_text + '7'
    elif(today_date_month =='9'):
        my_column_name = my_text + '8'
    elif(today_date_month =='10'):
        my_column_name = my_text + '9'
    elif(today_date_month =='11'):
        my_column_name = my_text + '10'
    elif(today_date_month =='12'):
        my_column_name = my_text + '11'

     
    return my_column_name
    
    
    

def leave_convert_ceiling05(df, field1, field2, field3 ):
#     listofmonth_int = ['01','02','03','04','05','06','07','08','09','10','11','12']
    listofmonth_int = ['1','2','3','4','5','6','7','8','9','10','11','12']
    for x in listofmonth_int:
        field_name_avg = field1 + str(x) + ')'
        field_name = field2 + str(int(x))
        df[field_name_avg] = np.ceil((df[field_name] / df[field3])  *2)/2

    return df



def rename_googlesheet(id1, rename_file):
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive 
    from google.colab import auth 
    from oauth2client.client import GoogleCredentials
    auth.authenticate_user()
    gauth = GoogleAuth()
    gauth.credentials = GoogleCredentials.get_application_default()
    drive = GoogleDrive(gauth)


    file1 = drive.CreateFile({'id': id1})
    file1.Upload()                 # Upload new title.
    a=drive.auth.service.files().get(fileId=id1).execute()
    a['title']=rename_file
    update=drive.auth.service.files().update(fileId=id1,body=a).execute()



    
def copy_fileid_to_specific_folder(id1, specific_folderid, rename_file):
    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive 
    from google.colab import auth 
    from oauth2client.client import GoogleCredentials
    auth.authenticate_user()
    gauth = GoogleAuth()
    gauth.credentials = GoogleCredentials.get_application_default()
    drive = GoogleDrive(gauth)

    file1 = drive.CreateFile({'id': id1})
    file1.Upload()                 # Upload new title.

    file_main = drive.auth.service.files().copy(fileId=id1, body={"parents": [{"id": specific_folderid}], 'title': rename_file}).execute()
    return file_main
    
#reference https://medium.com/@simonprdhm/how-to-send-emails-with-gmail-using-python-f4a8bcb6a9cc
def send_email_when_finish(finished_user, finished_pass, finished_send_to,  Folder_querymain_ID, Folder_querymain_for_HRMS, file_date ):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    # me == my email address
    # you == recipient's email address
    me = finished_user
    you = finished_send_to

    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "[Querymain Run Complete] ข้อมูล ณ วันที่ " + file_date
    msg['From'] = me
    msg['To'] = you

    file1 = "https://drive.google.com/drive/folders/" + Folder_querymain_ID
    file2 = "https://drive.google.com/drive/folders/" + Folder_querymain_for_HRMS
    
    #monitor
    file3 = "https://drive.google.com/drive/folders/1rkTLVu8UHJ338bfEuhAiwOVUMEhnsuV5"
    
    
    #สรุป report
    file4 = "https://docs.google.com/spreadsheets/d/1EaVFbZrNsiE9fT_DXrCoSN9PNZmzXte5WgB0SVT72PI/edit"
    
    #ทะเบียนพนักงาน
    file5 = "https://drive.google.com/drive/folders/15qmbj-jy0A4fIttJMjfYnz_XzPwDJh_8"
    
    #file Dashboard
    file6 = "https://drive.google.com/drive/folders/0B8onEo_a5GOlRTBScGRBY0R5MW8?resourcekey=0-cu-OiVCf02Nwj1cmhmxCLQ"
    # Create the body of the message (a plain-text and an HTML version).
#     text = "Hi!\nHow are you?\nHere is the link you wanted:\nhttp://www.python.org"



    html = """\
    <html>
      <head></head>
      <body>
        <p>Hi!<br>
           <b><h1>พี่ดำรันเสร็จแล้วนะ</h1></b>
           <u><h3>Main</h3></u>
           1. เข้าไปตรวจสอบ Querymain รายเดือนได้ที่ <a href="%s">link</a> (สำหรับคนทั่วไป) <br>
           2. Querymain ที่มีครบทุก field ให้เช็คที่นี้ <a href="%s">link</a> (สำหรับ HRMS) <br>
           3. Monitor ทุกวัน <a href="%s">link</a><br>
           
           <br>
           <u><h3>หัวข้อรายเดือน</h3></u>
           4. สรุป report ทีม <a href="%s">link</a><br>
           5. ทะเบียนพนักงาน <a href="%s">link</a><br>
           6. file dashboard <a href="%s">link</a><br>
        </p>
      </body>
    </html>
    """ % (file1, file2, file3, file4, file5, file6)

    # Record the MIME types of both parts - text/plain and text/html.
#     part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')

    # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, in this case
    # the HTML message, is best and preferred.
#     msg.attach(part1)
    msg.attach(part2)
    # Send the message via local SMTP server.
    mail = smtplib.SMTP('smtp.gmail.com', 587)

    mail.ehlo()

    mail.starttls()

    mail.login(me, finished_pass)
    mail.sendmail(me, you, msg.as_string())
    mail.quit()
    
    
    
    

def convert_column_to_date(df, field1):
    df[field1] = pd.to_datetime(df[field1],format='%d/%m/%Y' , errors='coerce')
    
    
    
    
    
def insert_intogooglesheet(url,df, sheet_name, clearspace, my_index_startrow):


    from pydrive.auth import GoogleAuth
    from pydrive.drive import GoogleDrive 
    from google.colab import auth 
    from oauth2client.client import GoogleCredentials


    auth.authenticate_user()
    gauth = GoogleAuth()
    gauth.credentials = GoogleCredentials.get_application_default()
    drive = GoogleDrive(gauth)

    #สำหรับจัดการ Google sheet
    import gspread
    from gspread_dataframe import set_with_dataframe
    from gspread_dataframe import get_as_dataframe
    from google.auth import default
    creds, _ = default()

    gc = gspread.authorize(creds)
    wb2 = gc.open_by_url(url)

    wb2.values_clear(clearspace.format(sheet_name))

    sheet_destination_sheet = wb2.worksheet(sheet_name)
    set_with_dataframe(sheet_destination_sheet, df, row=my_index_startrow, include_column_header=True) 
    
    
def create_list_fromdaterange(sdate, edate):
    from datetime import date, timedelta
    list1 = pd.date_range(sdate,edate-timedelta(days=0),freq='B')
    
    df_list1  = pd.DataFrame(list1)
    return df_list1
    
