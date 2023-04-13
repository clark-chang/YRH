#!/usr/bin/env python
# coding: utf-8

# In[1]:


#1.載入套件
#註記:pip install cx_Oracle

import os
import shutil
import pandas as pd
import cx_Oracle
import binascii
import datetime as dt
from datetime import datetime
import calendar


# In[2]:


#2.確認專案儲存路徑與新增資料夾
#註記:需網路下載安裝oracle instantclient_11_2
#註記:__file__
#os.chdir(r"C:\Users\Administrator")

if(os.path.exists(r"C:\Users\Administrator\YR")):
    shutil.rmtree(r"C:\Users\Administrator\YR")
    os.mkdir(r"C:\Users\Administrator\YR")
    os.chdir(r"C:\Users\Administrator\YR")
    print("已經移除舊資料夾與新增資料夾")
else:
    os.mkdir(r"C:\Users\Administrator\YR")
    os.chdir(r"C:\Users\Administrator\YR")
    print("已經新增資料夾")
print(os.getcwd())
print(os.path.abspath(''))


# In[3]:


#3.年月比對函數
def mappingyear(mapping):
    if mapping == 1:
        sqlyear2 = ''
    elif mapping == 2:
        sqlyear2 = 'B'
    elif mapping == 3:
        sqlyear2 = 'C'
    elif mapping == 4:
        sqlyear2 = 'D'
    elif mapping == 5:
        sqlyear2 = 'E'
    elif mapping == 6:
        sqlyear2 = 'F'    
    elif mapping == 7:
        sqlyear2 = 'G'
    elif mapping == 8:
        sqlyear2 = 'H'
    elif mapping == 9:
        sqlyear2 = 'I'
    elif mapping == 10:
        sqlyear2 = 'J'   
    return sqlyear2

def mappingmonth(mapping):
    if mapping == 1:
        sqlmonth1 = '01'
    elif mapping == 2:
        sqlmonth1 = '02'
    elif mapping == 3:
        sqlmonth1 = '03'
    elif mapping == 4:
        sqlmonth1 = '04'
    elif mapping == 5:
        sqlmonth1 = '05'
    elif mapping == 6:
        sqlmonth1 = '06'    
    elif mapping == 7:
        sqlmonth1 = '07'
    elif mapping == 8:
        sqlmonth1 = '08'
    elif mapping == 9:
        sqlmonth1 = '09'
    elif mapping == 10:
        sqlmonth1 = '10' 
    elif mapping == 11:
        sqlmonth1 = '11'
    else:
        sqlmonth1 = '12'
    return sqlmonth1


# In[4]:


#4.日期比對

#本月指標，如果要計算本月指標，解開下行註解(刪除#)
today = dt.datetime.now()

#上個月指標，如果要計算上個月指標，解開下行註解(刪除#)#today = datetime(2022,8,31,0,0)
#oday = datetime(int(str(datetime.now().year)), int(str(datetime.now().month-1).zfill(2)), int(str(calendar.monthrange(datetime.now().year, datetime.now().month-1)[1])))



yearcycle = ['','A','B','C','D','E','F','G','H','I','J']
adyear = today.year
year = adyear - 1911
quarter = ''
month = mappingmonth(today.month)
day = str(today.day).zfill(2)
sqlyear1 = year % 11
sqlyear9 = str(year)
sqlyear10 = str(year)
sqlyear11 = str(year)
sqlyear12 = str(year)
sqlmonth1 = today.month
sqlday1 = today.day
sqlyear = mappingyear(sqlyear1)
sqlmonth = mappingmonth(sqlmonth1)
sqlmonth7_1 = ""
sqlmonth7_2 = ""
sqlmonth7_3 = ""
sqlmonth8_1 = ""
sqlmonth8_2 = ""
sqlmonth8_3 = ""
sqlmonth9_1 = ""
sqlmonth9_2 = ""
sqlmonth9_3 = ""
sqlmonth10_1 = ""
sqlmonth10_2 = ""
sqlmonth10_3 = ""
sqlmonth11_1 = ""
sqlmonth11_2 = ""
sqlmonth11_3 = ""
sqlmonth12_1 = ""
sqlmonth12_2 = ""
sqlmonth12_3 = ""


print(sqlyear1)
print(sqlyear)
print(month, day)


# In[5]:


#5.SQL指標7字串參數,SQL指標8字串參數,SQL指標9字串參數,SQL指標10字串參數,SQL指標11字串參數，確認季份與月份
if (sqlmonth1  == 1):
    sqlmonth7_1 = "01"
    sqlmonth7_2 = ""
    sqlmonth7_3 = ""
    sqlmonth8_1 = "01"
    sqlmonth8_2 = ""
    sqlmonth8_3 = ""
    sqlmonth9_1 = "01"
    sqlmonth9_2 = "02"
    sqlmonth9_3 = "03"
    sqlmonth10_1 = "01"
    sqlmonth10_2 = "02"
    sqlmonth10_3 = "03"
    sqlmonth11_1 = "01"
    sqlmonth11_2 = "02"
    sqlmonth11_3 = "03"    
    sqlmonth12_1 = "01"
    sqlmonth12_2 = "02"
    sqlmonth12_3 = "03"
    
    quarter = 'Q1'
elif(sqlmonth1  == 2):
    sqlmonth7_1 = "01"
    sqlmonth7_2 = "02"
    sqlmonth7_3 = ""
    sqlmonth8_1 = "01"
    sqlmonth8_2 = "02"
    sqlmonth8_3 = ""
    sqlmonth9_1 = "01"
    sqlmonth9_2 = "02"
    sqlmonth9_3 = "03"
    sqlmonth10_1 = "01"
    sqlmonth10_2 = "02"
    sqlmonth10_3 = "03"
    sqlmonth11_1 = "01"
    sqlmonth11_2 = "02"
    sqlmonth11_3 = "03"    
    sqlmonth12_1 = "01"
    sqlmonth12_2 = "02"
    sqlmonth12_3 = "03"
    quarter = 'Q1'
elif(sqlmonth1  == 3):
    sqlmonth7_1 = "01"
    sqlmonth7_2 = "02"
    sqlmonth7_3 = "03"
    sqlmonth8_1 = "01"
    sqlmonth8_2 = "02"
    sqlmonth8_3 = "03"
    sqlmonth9_1 = "01"
    sqlmonth9_2 = "02"
    sqlmonth9_3 = "03"
    sqlmonth10_1 = "01"
    sqlmonth10_2 = "02"
    sqlmonth10_3 = "03"
    sqlmonth11_1 = "01"
    sqlmonth11_2 = "02"
    sqlmonth11_3 = "03"    
    sqlmonth12_1 = "01"
    sqlmonth12_2 = "02"
    sqlmonth12_3 = "03"
    quarter = 'Q1'
elif(sqlmonth1  == 4):
    sqlmonth7_1 = "04"
    sqlmonth7_2 = ""
    sqlmonth7_3 = ""
    sqlmonth8_1 = "04"
    sqlmonth8_2 = ""
    sqlmonth8_3 = ""
    sqlmonth9_1 = "04"
    sqlmonth9_2 = "05"
    sqlmonth9_3 = "06"
    sqlmonth10_1 = "04"
    sqlmonth10_2 = "05"
    sqlmonth10_3 = "06"
    sqlmonth11_1 = "04"
    sqlmonth11_2 = "05"
    sqlmonth11_3 = "06"    
    sqlmonth12_1 = "04"
    sqlmonth12_2 = "05"
    sqlmonth12_3 = "06"
    quarter = 'Q2'
elif(sqlmonth1  == 5):
    sqlmonth7_1 = "04"
    sqlmonth7_2 = "05"
    sqlmonth7_3 = ""
    sqlmonth8_1 = "04"
    sqlmonth8_2 = "05"
    sqlmonth8_3 = ""
    sqlmonth9_1 = "04"
    sqlmonth9_2 = "05"
    sqlmonth9_3 = "06"
    sqlmonth10_1 = "04"
    sqlmonth10_2 = "05"
    sqlmonth10_3 = "06"
    sqlmonth11_1 = "04"
    sqlmonth11_2 = "05"
    sqlmonth11_3 = "06"    
    sqlmonth12_1 = "04"
    sqlmonth12_2 = "05"
    sqlmonth12_3 = "06"
    quarter = 'Q2'
elif(sqlmonth1  == 6):
    sqlmonth7_1 = "04"
    sqlmonth7_2 = "05"
    sqlmonth7_3 = "06"
    sqlmonth8_1 = "04"
    sqlmonth8_2 = "05"
    sqlmonth8_3 = "06"
    sqlmonth9_1 = "04"
    sqlmonth9_2 = "05"
    sqlmonth9_3 = "06"
    sqlmonth10_1 = "04"
    sqlmonth10_2 = "05"
    sqlmonth10_3 = "06"
    sqlmonth11_1 = "04"
    sqlmonth11_2 = "05"
    sqlmonth11_3 = "06"    
    sqlmonth12_1 = "04"
    sqlmonth12_2 = "05"
    sqlmonth12_3 = "06"
    quarter = 'Q2'
elif(sqlmonth1  == 7):
    sqlmonth7_1 = "07"
    sqlmonth7_2 = ""
    sqlmonth7_3 = ""
    sqlmonth8_1 = "07"
    sqlmonth8_2 = ""
    sqlmonth8_3 = ""
    sqlmonth9_1 = "07"
    sqlmonth9_2 = "08"
    sqlmonth9_3 = "09"
    sqlmonth10_1 = "07"
    sqlmonth10_2 = "08"
    sqlmonth10_3 = "09"
    sqlmonth11_1 = "07"
    sqlmonth11_2 = "08"
    sqlmonth11_3 = "09"    
    sqlmonth12_1 = "07"
    sqlmonth12_2 = "08"
    sqlmonth12_3 = "09"
    quarter = 'Q3'
elif(sqlmonth1  == 8):
    sqlmonth7_1 = "07"
    sqlmonth7_2 = "08"
    sqlmonth7_3 = ""
    sqlmonth8_1 = "07"
    sqlmonth8_2 = "08"
    sqlmonth8_3 = ""
    sqlmonth9_1 = "07"
    sqlmonth9_2 = "08"
    sqlmonth9_3 = "09"
    sqlmonth10_1 = "07"
    sqlmonth10_2 = "08"
    sqlmonth10_3 = "09"
    sqlmonth11_1 = "07"
    sqlmonth11_2 = "08"
    sqlmonth11_3 = "09"    
    sqlmonth12_1 = "07"
    sqlmonth12_2 = "08"
    sqlmonth12_3 = "09"
    quarter = 'Q3'
elif(sqlmonth1  == 9):
    sqlmonth7_1 = "07"
    sqlmonth7_2 = "08"
    sqlmonth7_3 = "09"
    sqlmonth8_1 = "07"
    sqlmonth8_2 = "08"
    sqlmonth8_3 = "09"
    sqlmonth9_1 = "07"
    sqlmonth9_2 = "08"
    sqlmonth9_3 = "09"
    sqlmonth10_1 = "07"
    sqlmonth10_2 = "08"
    sqlmonth10_3 = "09"
    sqlmonth11_1 = "07"
    sqlmonth11_2 = "08"
    sqlmonth11_3 = "09"    
    sqlmonth12_1 = "07"
    sqlmonth12_2 = "08"
    sqlmonth12_3 = "09"
    quarter = 'Q3'
elif(sqlmonth1  == 10):
    sqlmonth7_1 = "10"
    sqlmonth7_2 = ""
    sqlmonth7_3 = ""
    sqlmonth8_1 = "10"
    sqlmonth8_2 = ""
    sqlmonth8_3 = ""
    sqlmonth9_1 = "10"
    sqlmonth9_2 = "11"
    sqlmonth9_3 = "12"
    sqlmonth10_1 = "10"
    sqlmonth10_2 = "11"
    sqlmonth10_3 = "12"
    sqlmonth11_1 = "10"
    sqlmonth11_2 = "11"
    sqlmonth11_3 = "12"    
    sqlmonth12_1 = "10"
    sqlmonth12_2 = "11"
    sqlmonth12_3 = "12"
    quarter = 'Q4'
elif(sqlmonth1  == 11):
    sqlmonth7_1 = "10"
    sqlmonth7_2 = "11"
    sqlmonth7_3 = "" 
    sqlmonth8_1 = "10"
    sqlmonth8_2 = "11"
    sqlmonth8_3 = ""
    sqlmonth9_1 = "10"
    sqlmonth9_2 = "11"
    sqlmonth9_3 = "12"
    sqlmonth10_1 = "10"
    sqlmonth10_2 = "11"
    sqlmonth10_3 = "12"
    sqlmonth11_1 = "10"
    sqlmonth11_2 = "11"
    sqlmonth11_3 = "12"    
    sqlmonth12_1 = "10"
    sqlmonth12_2 = "11"
    sqlmonth12_3 = "12"
    quarter = 'Q4'
elif(sqlmonth1  == 12):
    sqlmonth7_1 = "10"
    sqlmonth7_2 = "11"
    sqlmonth7_3 = "12"
    sqlmonth8_1 = "10"
    sqlmonth8_2 = "11"
    sqlmonth8_3 = "12"
    sqlmonth9_1 = "10"
    sqlmonth9_2 = "11"
    sqlmonth9_3 = "12"
    sqlmonth10_1 = "10"
    sqlmonth10_2 = "11"
    sqlmonth10_3 = "12"
    sqlmonth11_1 = "10"
    sqlmonth11_2 = "11"
    sqlmonth11_3 = "12"    
    sqlmonth12_1 = "10"
    sqlmonth12_2 = "11"
    sqlmonth12_3 = "12"
    quarter = 'Q4'


# In[6]:


#6.健保指標SQL
#6-1.YR申報前門診申報點數
sql1 = '''
SELECT                                                                         --門診申報點數
A.clinic_date,                                                                 --AS "就醫日",
A.chart_no,                                                                    --AS "病歷號",
rawtohex(utl_raw.cast_to_raw(A.pt_name)) AS pt_name,                           --AS "病人姓名",
                                                                               --F.SEQ,F.DISEASE_CODE AS "診斷碼",
rawtohex(utl_raw.cast_to_raw(C.doctor_name)) AS doctor_name,                   --AS "醫師姓名",
rawtohex(utl_raw.cast_to_raw(F.div_full_name)) AS div_full_name,               --AS "科別全名",
B.code,                                                                        --AS "院內碼",
B.nh_code,                                                                     --AS "健保碼",
D.map_onh_acnt_no,                                                             --AS "健保科目",
rawtohex(utl_raw.cast_to_raw(E.acnt_full_name)) AS acnt_full_name,             --AS "科目名稱",
( B.nh_amt_1 + B.nh_amt_2 ),                                                   --AS "健保點數",
A.building_no                                                                  --AS "院區"
FROM   onh{0}{1}.acntopd B
       left join onh{0}{1}.ptopd A
              ON A.chart_no = B.chart_no
                 AND B.clinic_date = a.clinic_date
                 AND A.duplicate_no = B.duplicate_no
       left join mast.doctor C
              ON A.doctor_no = C.doctor_no
       left join mast.acntmap D
              ON B.acnt_no_1 = D.acnt_no
       left join mast.acntname E
              ON D.map_onh_acnt_no = E.acnt_no
       left join mast.div F
              ON C.div_no = F.div_no
WHERE  E.acnt_type = '2'                                                       --費用科目類別 = 2健保門診
       AND A.div_no != '60'                                                    --排除中醫
       AND A.pt_type IN (SELECT pt_type
                         FROM   mast.pttype
                         WHERE  insurance_type = '2'
                                AND supple_flag <> 'Y')
       AND A.merge_flag NOT IN 'D'
                                                                               --排除同療程合併註記->D:合併後刪除資料
       AND A.nh_apply_type IN ( '1' )
       AND B.pricing_flag NOT IN ( 'S', 's', 'Y', 'y',
                                   'h', 'X', 'x', 'Z' )
       AND A.div_no NOT IN ( 'EA', '2B' )
       AND ( B.nh_amt_1 != 0.00
              OR B.nh_amt_2 != 0.00 )
       AND A.nh_apply_flag IN ( 'Y', 'U' )
       AND B.nh_apply_flag IN( 'Y' )                                           --Y:申報 N:轉下月申報 X:不申報
                                                                               --AND A.APPLY_CASE_TYPE NOT IN ('A3','D2','C1')
                                                                               --排除門診血液透析費
       AND ( A.clinic_flag NOT IN ( 'O', 'V', 'H', 'F', 'P' )
              OR ( A.clinic_flag IN ( 'O', 'V', 'H', 'F', 'P' )
                   AND D.map_onh_acnt_no NOT IN ( '010' ) ) )
       AND ( B.ref_no = ' '                                                    --單據號碼是一個空格
              OR ( B.ref_no <> ' '
                   AND B.acnt_no_1 IN (SELECT acnt_no
                                       FROM   mast.acntmap
                                       WHERE  map_prog_acnt_no BETWEEN
                                              '040' AND '048')
                                                                               --明細科目的對應程式科目 040-048
                   AND B.ref_no IN (SELECT apply_no
                                    FROM   xray.ptexamindex
                                    WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                  )
              OR ( B.acnt_no_1 IN (SELECT acnt_no
                                   FROM   mast.acntmap
                                   WHERE  map_prog_acnt_no BETWEEN
                                          '120' AND '128')
                                                                               --明細科目的對應程式科目 120-128
                   AND ( B.ref_no ) IN (SELECT apply_no
                                        FROM   exam.ptexamindex
                                        WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                  )
              OR ( B.ref_no <> ' '
                   AND B.acnt_no_1 IN (SELECT acnt_no
                                       FROM   mast.acntmap
                                       WHERE  map_prog_acnt_no IN (
                                              '050', '056', '058'
                                                                  ))
                                                                               --明細科目的對應程式科目 050, 056, 058
                   AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                   FROM   lab.labx1
                                                   WHERE
                       report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"細菌初步報告表頭"檢驗單號 確認有檢查
                  )
              OR ( B.ref_no <> ' '
                   AND B.acnt_no_1 IN (SELECT acnt_no
                                       FROM   mast.acntmap
                                       WHERE  map_prog_acnt_no IN ( '055' ))
                                                                               --明細科目的對應程式科目 055
                   AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                   FROM   lam.labx1
                                                   WHERE
                       report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"??? LAM ?"檢驗單號 確認有檢查
                  ) ) 
'''
print(sql1)


# In[7]:


#6-2.YR代辦案件預估
sql2 = '''
SELECT Substr(a.clinic_date, 0, 5),--"年月",
       a.clinic_date              ,--"就醫日期",
       d.clinic_apn               ,--"午別",
       d.clinic_flag              ,--"診別",
       rawtohex(utl_raw.cast_to_raw(d.clinic_no)),                --"診間號",
       a.chart_no                 ,--"病歷號",
       a.code                     ,--"院內碼",
       rawtohex(utl_raw.cast_to_raw(F.pt_type_full_name)),         --"身份別",
       a.doctor_no                ,--"醫師代碼",
       rawtohex(utl_raw.cast_to_raw(c.doctor_name)),              --"醫師姓名",
       rawtohex(utl_raw.cast_to_raw(b.div_full_name)),               --"科別名稱",
       d.nh_clinic_seq            ,--"健保就醫序號",
       A.total_qty                ,--"開立總量",
       ( A.nh_amt_1 )             ,--"健保金額"
                                   --主鍵對應 就醫日+午別+診間號+診別,
       CASE A.nh_apply_flag
          WHEN 'Y' THEN 'A5D3B3F8'
          WHEN 'N' THEN 'C2E0A455A4EBA5D3B3F8'
          WHEN 'X' THEN 'A4A3A5D3B3F8'
       END                         nh_apply_flag,
       rawtohex(utl_raw.cast_to_raw(H.acnt_full_name))      ,--"收據科目",
       H.acnt_full_name           ,--"收據科目名稱",
       D.apply_case_type          ,--"案件分類",
       D.building_no              --"院區"
                                   --,A.NH_APPLY_FLAG,(A.NH_AMT_1+A.NH_AMT_2),(A.SELF_AMT_1+A.SELF_AMT_2)
                                   --Y:申報 N:轉下月申報 X:不申報
FROM   onh{0}{1}.acntopd a
       LEFT JOIN onh{0}{1}.ptopd d
              ON a.chart_no = d.chart_no
                 AND a.clinic_date = d.clinic_date
                 AND a.duplicate_no = d.duplicate_no
       LEFT JOIN mast.div b
              ON a.div_no = b.div_no
       LEFT JOIN mast.doctor c
              ON D.doctor_no = c.doctor_no
       LEFT JOIN mast.pttype F
              ON D.pt_type = F.pt_type
       LEFT JOIN mast.acntmap G
              ON A.acnt_no_1 = G.acnt_no
       LEFT JOIN mast.acntname H
              ON G.map_receipt_acnt_no = H.acnt_no
WHERE  D.div_no != '60'
       AND A.nh_apply_flag NOT IN ( 'N', 'X' ) --不含轉下期及不申報
       AND D.merge_flag != 'A' --不含合併療程後資料
       AND D.nh_apply_flag IN ( 'Y', 'U' )
       AND h.acnt_type = '1' --類別為收據科目
       AND d.korder_flag IN ( 'Y', 'P' ) --已看診及批價
       AND d.clinic_flag IN ( 'O', 'V', 'E' ) --門診與約診
       AND D.apply_case_type IN ( 'A3', 'B1', 'B6', 'B7',
                                  'B8', 'B9', 'C4', 'D1',
                                  'D2', 'NH', 'BA', 'DF',
                                  'E2', 'E3', 'C5' ) --代辦之案件分類
       AND A.self_amt_1 = '0' --自費等於0
       AND a.clinic_date LIKE '{2}{1}%'
GROUP  BY a.clinic_date,
          d.clinic_apn,
          d.clinic_flag,
          d.clinic_no,
          a.chart_no,
          a.code,
          F.pt_type_full_name,
          a.doctor_no,
          c.doctor_name,
          b.div_full_name,
          d.nh_clinic_seq,
          A.total_qty,
          A.nh_amt_1,
          --主鍵對應 就醫日+午別+診間號+診別
          A.nh_apply_flag,
          G.map_receipt_acnt_no,
          H.acnt_full_name,
          D.apply_case_type,
          D.building_no 
'''





# In[8]:


#6-3.YR-BC肝(R)
sql3='''
SELECT                                                                         --BC肝
a.clinic_date,                                                                 --AS "就醫日",
a.chart_no,                                                                    --AS "病歷號",
rawtohex(utl_raw.cast_to_raw(a.pt_name)),                                      --AS "病人姓名",
f.seq,                                                                         --AS "序號",
f.disease_code,                                                                --AS "診斷碼",
rawtohex(utl_raw.cast_to_raw(c.doctor_name)),                                  --AS "醫師姓名",
rawtohex(utl_raw.cast_to_raw(g.div_full_name)),                                --AS" 科別全名",
b.code,                                                                        --AS "院內碼",
b.nh_code,                                                                     --AS "健保碼",
d.map_onh_acnt_no,                                                             --AS "健保科目",
rawtohex(utl_raw.cast_to_raw(e.acnt_full_name)),                               --AS "科目名稱",
Sum(b.nh_amt_1 + b.nh_amt_2),                                                  --AS "健保金額",
a.building_no                                                                  --AS "院區"
FROM   onh{0}{1}.acntopd b
       LEFT JOIN onh{0}{1}.ptopd a
              ON a.chart_no = b.chart_no
       LEFT JOIN mast.doctor c
              ON a.doctor_no = c.doctor_no
       LEFT JOIN mast.acntmap d
              ON b.acnt_no_1 = d.acnt_no
       LEFT JOIN mast.acntname e
              ON d.map_onh_acnt_no = e.acnt_no
       LEFT JOIN onh{0}{1}.ordaopd10 f
              ON b.chart_no = f.chart_no
       LEFT JOIN mast.div g
              ON c.div_no = g.div_no
WHERE  b.nh_code IN ( 'A043302100', 'A044650100', 'A048027100', 'A048152100',
                      'AC44650100', 'AC48027100', 'AC57322100', 'B016536209',
                      'B016536216', 'B016536221', 'B016536229', 'B016536299',
                      'B016955216', 'B016955220', 'B016955227', 'B016955237',
                      'B022178216', 'B022178220', 'B022178223', 'B022178227',
                      'B022178237', 'B022490216', 'B022490237', 'B022491221',
                      'B022491229', 'B023208100', 'B023920100', 'B024468100',
                      'B024469100', 'B024662100', 'B024690100', 'B025219166',
                      'B025232100', 'BC23208100', 'K000589237', 'K000591243',
                      'K000667255', 'K000669248', 'K000674253', 'K000675257',
                      'K000696266', 'K000700216', 'K000700220', 'K000700223',
                      'K000700227', 'K000752221', 'K000753216', 'K000754209',
                      'K000765209', 'K000766209', 'K000788277', 'K000789277',
                      'K000815248', 'K000816255', 'K000817257', 'K000818253',
                      'A049918100', 'AB48027100', 'AB57268100', 'AC57268100',
                      'AC57786100', 'AC57856100', 'AC57923100', 'B025605100',
                      'AC58039100', 'BC25232100', 'X000130216', 'AC49918100',
                      'AC58063100', 'AC58091100', 'K000929277', 'K000930277',
                      'AB57923100', 'BC26310100', 'AC43302100', 'AC58334100',
                      'AC58335100', 'BA24468100', 'BA24469100', 'BC23920100',
                      'BC24662100', 'BC24690100', 'BC25219166', 'KC00589237',
                      'KC00667255', 'KC00669248', 'KC00674253', 'KC00675257',
                      'KC00788277', 'KC00789277', 'KC00815248', 'KC00816255',
                      'KC00817257', 'KC00818253', 'KC00929277', 'AC58541100',
                      'AC58542100', 'AB57786100', 'AC48152100', 'BC25605100',
                      'KC00700216', 'KC00871277', 'KC00872277', 'AA57786100',
                      'AB58091100', 'KC00930277', 'AA57322100', 'AC58778100',
                      'KC00974277', 'KC00975277', 'AA58091100', 'AA58542100',
                      'AC58369100', 'BC26716100', 'BC26717100', 'AC59253100',
                      'AC59378100', 'BC26944100', 'BC26945100', 'AC59621100',
                      'BC27042100', 'AC59671100', 'AC59704100', 'AC59723100',
                      'BC27086100', 'AC58600100', 'AC59750100', 'AC59872100',
                      'AC59895100', 'AC60109100', 'AC60178100', 'AC60206100',
                      'AC60219100', 'BC27375100', 'BC27700100', 'BC27701100',
                      'BC27829100', 'AC60515100' )
       AND f.seq IN ( '00' )
       AND a.div_no != '60'                                                    --排除中醫
       AND a.clinic_date = b.clinic_date
       AND f.clinic_date = b.clinic_date
       AND a.duplicate_no = b.duplicate_no
       AND a.chart_no = f.chart_no
       AND a.duplicate_no = f.duplicate_no
       AND a.nh_apply_flag IN ( 'Y', 'U' )
                                                                               -- nh_apply_flag申報否 U:尚未審核 Y:確定申報 N:暫不申報 X:永不申報
       AND b.nh_apply_flag = 'Y'
                                                                               -- nh_apply_flag申報否  Y:申報 N:轉下月申報 X:不申報
       AND a.nh_apply_type = '1'
                                                                               --申報類別 1:送核 2:補報 3:福山補報 9:專案補報 A:整筆就醫補報(A) a:部份醫令補報(a) B:整筆就醫補報(B) b:部份醫令補報(b) ....
       AND a.merge_flag NOT IN ( 'D', 'd' )
                                                                               --同療程合併註記 A:合併後新資料 D:合併後刪除資料
       AND b.pricing_flag NOT IN ( 'S', 's', 'Y', 'y',
                                   'h', 'X', 'x', 'Z' )
                                                                               --計價註記 Y/N/S,s/H,h/X,x/Z/V  Y:自費計價 N:健保計價申報 H:健保不計價申報(user KEYIN) h:健保不計價申報(電腦依price,joint,C1,05案件設定) S:任何身份皆自費(user KEYIN) s:任何身份皆自費(電腦依joint設定) X:不計價不申報(user KEYIN) x:不計價不申報(電腦依price,joint設定) Z:自費病人自費,健保病人不申報不計價 V:虛醫令,交付調劑之藥品空針
       AND a.div_no NOT IN ( 'EA', '2B' )
                                                                               --科別 NOT IN 居家照護(EA)及洗腎(2B)
       AND b.apply_order_type IN ( '1' )
       AND e.acnt_type = '2'
                                                                               --費用科目類別 0明細科目 1收據科目 2健保門診 3健保住院 P程式科目 R報表科目 T報表科目2
       AND b.ref_no = ' '                                                      --單據號碼 : 檢驗申請單,檢查申請單etc
       AND d.map_onh_acnt_no NOT IN ( '010' )
                                                                               --對應門診健保科目 不等於血液透析費(010)
GROUP  BY a.clinic_date,
          a.chart_no,
          a.pt_name,
          c.doctor_name,
          g.div_full_name,
          b.nh_code,
          d.map_onh_acnt_no,
          e.acnt_full_name,
          f.seq,
          f.disease_code,
          b.code,
          a.building_no 
'''


# In[9]:


#6-4.YR-C肝口服(R)
sql4 = '''
SELECT                                                                         -- C肝
A.clinic_date,                                                                 --AS "就醫日",
A.chart_no,                                                                    --AS "病歷號",
rawtohex(utl_raw.cast_to_raw(A.pt_name)),                                      --AS "病人姓名",
F.seq,                                                                         --AS "序號",
F.disease_code,                                                                --AS "診斷碼",
rawtohex(utl_raw.cast_to_raw(C.doctor_name)),                                  --AS "醫師姓名",
rawtohex(utl_raw.cast_to_raw(G.div_full_name)),                                --AS "科別全名",
B.code,                                                                        --AS "院內碼",
B.nh_code,                                                                     --AS "健保碼",
D.map_onh_acnt_no,                                                             --AS "健保科目",
rawtohex(utl_raw.cast_to_raw(E.acnt_full_name)),                               --AS "科目名稱",
Sum(B.nh_amt_1 + B.nh_amt_2) ,                                                 --AS "健保金額",
A.building_no                                                                  --AS "院區"
FROM   onh{0}{1}.acntopd B
       LEFT JOIN onh{0}{1}.ptopd A
              ON A.chart_no = B.chart_no
       LEFT JOIN mast.doctor C
              ON A.doctor_no = C.doctor_no
       LEFT JOIN mast.acntmap D
              ON B.acnt_no_1 = D.acnt_no
       LEFT JOIN mast.acntname E
              ON D.map_onh_acnt_no = E.acnt_no
       LEFT JOIN onh{0}{1}.ordaopd10 F
              ON B.chart_no = F.chart_no
       LEFT JOIN mast.div G
              ON C.div_no = G.div_no
WHERE  B.nh_code IN ( 'HCVDAA0001', 'HCVDAA0002', 'HCVDAA0003', 'HCVDAA0004',
                      'HCVDAA0005', 'HCVDAA0006', 'HCVDAA0007', 'HCVDAA0008',
                      'HCVDAA0009', 'HCVDAA0010', 'HCVDAA0011', 'HCVDAA0012',
                      'HCVDAA0013', 'HCVDAA0014', 'HCVDAA0015', 'HCVDAA0016',
                             'HCVDAA0017' )
       AND F.seq IN ( '00' )
       AND a.div_no != '60'                                                    --排除中醫
       AND A.clinic_date = B.clinic_date
       AND F.clinic_date = B.clinic_date
       AND A.duplicate_no = B.duplicate_no
       AND A.chart_no = F.chart_no
       AND A.duplicate_no = F.duplicate_no
       AND a.nh_apply_flag IN ( 'Y', 'U' )
                                                                               -- nh_apply_flag申報否 U:尚未審核 Y:確定申報 N:暫不申報 X:永不申報
       AND B.nh_apply_flag = 'Y'
                                                                               -- nh_apply_flag申報否  Y:申報 N:轉下月申報 X:不申報
       AND a.nh_apply_type = '1'
                                                                               --申報類別 1:送核 2:補報 3:福山補報 9:專案補報 A:整筆就醫補報(A) a:部份醫令補報(a) B:整筆就醫補報(B) b:部份醫令補報(b) ....
       AND A.merge_flag NOT IN ( 'D', 'd' )
                                                                               --同療程合併註記 A:合併後新資料 D:合併後刪除資料
       AND B.pricing_flag NOT IN ( 'S', 's', 'Y', 'y',
                                   'h', 'X', 'x', 'Z' )
                                                                               --計價註記 Y/N/S,s/H,h/X,x/Z/V  Y:自費計價 N:健保計價申報 H:健保不計價申報(user KEYIN) h:健保不計價申報(電腦依price,joint,C1,05案件設定) S:任何身份皆自費(user KEYIN) s:任何身份皆自費(電腦依joint設定) X:不計價不申報(user KEYIN) x:不計價不申報(電腦依price,joint設定) Z:自費病人自費,健保病人不申報不計價 V:虛醫令,交付調劑之藥品空針
       AND A.div_no NOT IN ( 'EA', '2B' )
                                                                               --科別 NOT IN 居家照護(EA)及洗腎(2B)
       AND B.apply_order_type IN ( '1' )
       AND E.acnt_type = '2'
                                                                               --費用科目類別 0明細科目 1收據科目 2健保門診 3健保住院 P程式科目 R報表科目 T報表科目2
       AND b.ref_no = ' '                                                      --單據號碼 : 檢驗申請單,檢查申請單etc
       AND D.map_onh_acnt_no NOT IN ( '010' )
                                                                               --對應門診健保科目 不等於血液透析費(010)
GROUP  BY A.clinic_date,
          A.chart_no,
          A.pt_name,
          C.doctor_name,
          G.div_full_name,
          B.nh_code,
          D.map_onh_acnt_no,
          E.acnt_full_name,
          F.seq,
          F.disease_code,
          B.code,
          A.building_no 
'''


# In[10]:


#6-5.YR門診非藥費(急轉住)
sql5 = '''
SELECT                                                                         --非藥費 1101228更新
A.clinic_date,                                                                 --AS "就醫日",
A.chart_no,                                                                    --AS "病歷號",
rawtohex(utl_raw.cast_to_raw(A.pt_name)),                                      --AS "病人姓名",
                                                                               --F.SEQ AS " 序號",
                                                                               --F.DISEASE_CODE AS "診斷碼",
rawtohex(utl_raw.cast_to_raw(C.doctor_name)),                                  --AS "醫師姓名",
rawtohex(utl_raw.cast_to_raw(F.div_full_name)),                                --AS "科別全名",
B.code,                                                                        --AS "院內碼",
B.nh_code,                                                                     --AS "健保碼",
D.map_onh_acnt_no,                                                             --AS "健保科目",
rawtohex(utl_raw.cast_to_raw(E.acnt_full_name)),                               --AS "科目名稱",
( B.nh_amt_1 + B.nh_amt_2 ),                                                   --AS "健保點數",
a.clinic_flag,
CASE g.er_acnt_close_flag
  WHEN 'T0' THEN 'C2E0A6EDB07C'
  WHEN 'T1' THEN 'C2E0A6EDB07CA440AFEBAF66A9D0'
  WHEN 'T2' THEN 'C2E0A6EDB07CA55BC540AF66A9D0'
  ELSE 'AAC5A5D5'
END                         er_acnt_close_flag,                                --AS "轉住院註記",
A.building_no                                                                  --AS "院區"
                                                                                /*轉住院註記的轉換16位元進制
                                                                                SELECT Rawtohex(utl_raw.Cast_to_raw('轉住院'))         AS  "轉住院_轉16位元進制",
                                                                                       Rawtohex(utl_raw.Cast_to_raw('轉住院一般病房')) AS "轉住院一般病房_轉16位元進制)",
                                                                                       Rawtohex(utl_raw.Cast_to_raw('轉住院加護病房')) AS "轉住院加護病房_轉16位元進制",
                                                                                       Rawtohex(utl_raw.Cast_to_raw(' '))              AS " _轉16位元進制"
                                                                                FROM   dual 
                                                                                */
FROM   onh{0}{1}.acntopd B
       left join onh{0}{1}.ptopd A
              ON A.chart_no = B.chart_no
                 AND B.clinic_date = a.clinic_date
                 AND A.duplicate_no = B.duplicate_no
       left join onh{0}{1}.pter g
              ON A.chart_no = g.chart_no
                 AND A.clinic_date = g.clinic_date
                 AND A.duplicate_no = g.duplicate_no
       left join mast.doctor C
              ON A.doctor_no = C.doctor_no
       left join mast.acntmap D
              ON B.acnt_no_1 = D.acnt_no
       left join mast.acntname E
              ON D.map_onh_acnt_no = E.acnt_no
       left join mast.div F
              ON C.div_no = F.div_no
WHERE  E.acnt_type = '2'                                                       --費用科目類別 = 2健保門診
       AND a.div_no != '60'                                                    --排除中醫
       AND B.apply_order_type NOT IN ( '1', '4' )
                                                                               --醫令類別排除藥品'1'及內含項目'4' 1-藥品,2-其他,3-材料,4-內含項,ex:論病例計酬,洗腎,復健,中醫藥品內含項目
                                                                               --一般性排除
       AND A.merge_flag NOT IN 'D'
                                                                               --排除同療程合併註記->D:合併後刪除資料
                                                                               --AND A.NH_APPLY_TYPE NOT IN ('2')   --申報類別排除補報'2'
       AND ( B.nh_amt_1 != 0.00
              OR B.nh_amt_2 != 0.00 )
       AND b.nh_apply_flag IN( 'Y' )                                           --Y:申報 N:轉下月申報 X:不申報
       AND A.apply_case_type NOT IN ( 'A3', 'D2', 'C1' )
                                                                               --排除B類案件(代辦業務)
                                                                               --B1_1 案件類別
       AND A.apply_case_type NOT IN ( 'B1', 'B6', 'B7', 'B8',
                                      'B9', 'BA', 'DF', 'C5' )
                                                                               --B1_2
                                                                               --愛滋病--無
                                                                               --B1_3
       AND ( B.apply_order_type NOT IN '2'
              OR ( B.apply_order_type IN '2'
                   AND B.nh_code NOT IN ( '11', '12', '13', '15',
                                          '16', '17', '19', '20',
                                          '71', '72', '73', '75',
                                          '76', '77', '79', '31',
                                          '33', '35', '37', '91',
                                          '93', '21', '22', '23',
                                          '24', '25', '26', '27',
                                          '28', '85', '95', '97',
                                          'L1001C', '41', '50', '61',
                                          '63', '64', '65', '66',
                                          '68', '69', '4A', '4B',
                                          '4C', '4D', '4E', '6A',
                                          '6B', '6C', '6D', 'P3501C',
                                          'P3502C', 'P3503C', '21', '22',
                                          '23', '24', '25', '26',
                                          '27', '28', '21+L1001C', '25+L1001C',
                                          '27+L1001C', '98', '99' ) ) )
                                                                               --B1_4後半
       AND ( B.apply_order_type NOT IN '2'
              OR ( B.apply_order_type IN '2'
                   AND B.nh_code NOT IN ( 'E4003C', 'E4004C', 'E4005C' ) ) )
                                                                               --B1_4
                                                                               /*      AND (A.APPLY_CASE_TYPE NOT IN ('D2')
                                                                               
                                                                                          OR ( A.APPLY_CASE_TYPE IN ('D2')
                                                                               
                                                                                         AND B.NH_CODE NOT IN ('A2001C','KH1N2C','A2051C','A2052C','A3001C','E5002C','E5003C','E5004C','E5005C')       
                                                                                             )       
                                                                                           )
                                                                               */
                                                                               --B1_5
       AND A.apply_case_type NOT IN ( 'A1', 'A2', 'A5', 'A6', 'A7' )
       AND ( A.apply_case_type NOT IN 'E1'
              OR ( A.apply_case_type IN 'E1'
                   AND Trim(A.special_exam_item_1) != 'EC'
                   AND Trim(A.special_exam_item_2) != 'EC'
                   AND Trim(A.special_exam_item_3) != 'EC'
                   AND Trim(A.special_exam_item_4) != 'EC' ) )
                                                                               --B1_8
       AND A.apply_case_type NOT IN ( 'C4' )
                                                                               --C1_5前半
       AND ( B.apply_order_type NOT IN '2'
              OR ( B.apply_order_type IN '2'
                   AND B.nh_code NOT IN ( 'P1407C', 'P1408C', 'P1409C', 'P1410C'
                                          ,
                                          'P1411C', 'P1612C', 'P1613C', 'P1614B'
                                          ,
                                          'P1615C', 'P1801C', 'P1802C', 'P1803C'
                                          ,
                                          'P3402C', 'P3403C', 'P3404C', 'P3405C'
                                          ,
                                          'P3406C', 'P3407C', 'P3408C', 'P3409C'
                                          ,
                                          'P3903C', 'P3904C', 'P3905C', 'P4201C'
                                          ,
                                          'P4202C', 'P4203C', 'P4204C', 'P4205C'
                                          ,
                                          'P4301C', 'P4302C', 'P4303C', 'P5113B'
                                          ,
                                          'P5114B', 'P5115B', 'P5116B', 'P5117B'
                                          ,
                                          'P5118B', 'P5123B', 'P5124B', 'P5125B'
                                          ,
                                          'P5501B', 'P5502B', 'P5503B', 'P5504B'
                                          ,
                                          'P5505B', 'P5510B', 'P5511B', 'P5512B'
                                          ,
                                          'P5513B', 'P5514B', 'P5515B', 'P5516B'
                                          ,
                                          'P5517B', 'P5301C', 'P3410C', 'P3411C'
                                          ,
                                          'P6011C', 'P6012C', 'P6013C', 'P6014C'
                                          ,
                                          'P6015C', 'P5506B', 'P5507B', 'P5508B'
                                          ,
                                          'P5509B', 'P5126B', 'P5127B', 'P5128B'
                                          ,
                                          'P5132C', 'P5135B' ) ) )
                                                                               --C1_5後半
       AND ( B.apply_order_type NOT IN ( '0' )
              OR ( B.apply_order_type IN ( '0' )
                   AND B.nh_code NOT IN ( 'P5201C', 'P5202C', 'P5203C', 'P5204C'
                                        ) ) )
                                                                               --C1_7
       AND ( B.apply_order_type NOT IN ( '2', 'X', 'Z', 'K' )
              OR ( B.apply_order_type IN ( '2', 'X', 'Z', 'K' )
                   AND B.nh_code NOT IN ( 'P4601B', 'P4602B', 'P4603B', 'P4604B'
                                          ,
                                          'P4605B', 'P4606B', 'P4607B', 'P4608B'
                                          ,
                                          'P4609B', 'P4610B', 'P4611B', 'P4612B'
                                          ,
                                          'P4613B', 'P4614B', 'P4615B', 'P4616B'
                                          ,
                                          'P4617B', 'P4618B' ) ) )
                                                                               --C1_8
       AND ( A.special_exam_item_1 NOT IN ( 'JA', 'JB' )
              OR A.special_exam_item_2 NOT IN ( 'JA', 'JB' )
              OR A.special_exam_item_3 NOT IN ( 'JA', 'JB' )
              OR A.special_exam_item_4 NOT IN ( 'JA', 'JB' ) )
                                                                               --D1_1
       AND ( A.apply_case_type NOT IN 'E1'
              OR ( A.apply_case_type NOT IN 'E1'
                   AND Trim(A.special_exam_item_1) != 'G8' ) )
                                                                               --D1_2
       AND A.apply_case_type NOT IN '03'
                                                                               --D1_4
       AND B.nh_code NOT IN ( '05221A' )
                                                                               --排除門診血液透析費
       AND ( A.clinic_flag NOT IN ( 'O', 'V', 'H', 'F', 'P' )
              OR ( A.clinic_flag IN ( 'O', 'V', 'H', 'F', 'P' )
                   AND D.map_onh_acnt_no NOT IN ( '010' ) ) )
       AND ( B.ref_no = ' '                                                    --單據號碼是一個空格
              OR ( B.ref_no <> ' '
                   AND B.acnt_no_1 IN (SELECT acnt_no
                                       FROM   mast.acntmap
                                       WHERE  map_prog_acnt_no BETWEEN
                                              '040' AND '048')
                                                                               --明細科目的對應程式科目 040-048
                   AND B.ref_no IN (SELECT apply_no
                                    FROM   xray.ptexamindex
                                    WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                  )
              OR ( B.acnt_no_1 IN (SELECT acnt_no
                                   FROM   mast.acntmap
                                   WHERE  map_prog_acnt_no BETWEEN
                                          '120' AND '128')
                                                                               --明細科目的對應程式科目 120-128
                   AND ( B.ref_no ) IN (SELECT apply_no
                                        FROM   exam.ptexamindex
                                        WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                  )
              OR ( B.ref_no <> ' '
                   AND B.acnt_no_1 IN (SELECT acnt_no
                                       FROM   mast.acntmap
                                       WHERE  map_prog_acnt_no IN (
                                              '050', '056', '058'
                                                                  ))
                                                                               --明細科目的對應程式科目 050, 056, 058
                   AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                   FROM   lab.labx1
                                                   WHERE
                       report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"細菌初步報告表頭"檢驗單號 確認有檢查
                  )
              OR ( B.ref_no <> ' '
                   AND B.acnt_no_1 IN (SELECT acnt_no
                                       FROM   mast.acntmap
                                       WHERE  map_prog_acnt_no IN ( '055' ))
                                                                               --明細科目的對應程式科目 055
                   AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                   FROM   lam.labx1
                                                   WHERE
                       report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"???LAM?"檢驗單號 確認有檢查
                  ) ) 
'''


# In[11]:


#6-6.YR門診藥費(急轉住)
sql6 = '''
SELECT                                                                         --藥費 1091123 更新
A.clinic_date,                                                                 --AS "就醫日",
A.chart_no,                                                                    --AS "病歷號",
rawtohex(utl_raw.cast_to_raw(A.pt_name)),                                      --AS "病人姓名",
F.seq,                                                                         --AS "序號",
F.disease_code,                                                                --AS "診斷碼",
rawtohex(utl_raw.cast_to_raw(C.doctor_name)),                                  --AS "醫師姓名",
rawtohex(utl_raw.cast_to_raw(G.div_full_name)),                                --AS "科別全名",
B.code,                                                                        --AS "院內碼",
B.nh_code,                                                                     --AS "健保碼",
D.map_onh_acnt_no,                                                             --AS "健保科目",
rawtohex(utl_raw.cast_to_raw(E.acnt_full_name)),                               --AS "科目名稱",
SUM(B.nh_amt_1 + B.nh_amt_2),                                                  --AS "健保金額",
a.clinic_flag,                                                                 --AS "診別",
CASE h.er_acnt_close_flag
  WHEN 'T0' THEN 'C2E0A6EDB07C'
  WHEN 'T1' THEN 'C2E0A6EDB07CA440AFEBAF66A9D0'
  WHEN 'T2' THEN 'C2E0A6EDB07CA55BC540AF66A9D0'
  ELSE           'AAC5A5D5'
END                          er_acnt_close_flag,                               --AS "轉住院註記",
a.building_no                                                                  --AS "院區"
                                                                                /*離院關帳否的轉換16位元進制
                                                                                SELECT Rawtohex(utl_raw.Cast_to_raw('轉住院'))         AS "轉住院_轉16位元進制",
                                                                                       Rawtohex(utl_raw.Cast_to_raw('轉住院一般病房')) AS "轉住院一般病房_轉16位元進制)",
                                                                                       Rawtohex(utl_raw.Cast_to_raw('轉住院加護病房')) AS "轉住院加護病房_轉16位元進制",
                                                                                       Rawtohex(utl_raw.Cast_to_raw('空白'))           AS "空白_轉16位元進制"
                                                                                FROM   dual 
                                                                                */
FROM   onh{0}{1}.acntopd B
       left join onh{0}{1}.ptopd A
              ON A.chart_no = B.chart_no
       left join onh{0}{1}.pter h
              ON A.chart_no = h.chart_no
                 AND A.clinic_date = h.clinic_date
                 AND A.duplicate_no = h.duplicate_no
       left join mast.doctor C
              ON A.doctor_no = C.doctor_no
       left join mast.acntmap D
              ON B.acnt_no_1 = D.acnt_no
       left join mast.acntname E
              ON D.map_onh_acnt_no = E.acnt_no
       left join onh{0}{1}.ordaopd10 F
              ON B.chart_no = F.chart_no
       left join mast.div G
              ON C.div_no = G.div_no
                                                                               -- PTOPD門急診病患檔,(A), ACNTOPD門診病患醫令明細檔(B)  DOCTOR 醫師代碼檔(C) ACNTMAP科目對照檔(D), ACNTNAME科目名稱檔,(E) ORDAOPD10門診病患診斷檔10(F)
WHERE  a.pt_type IN (SELECT pt_type
                     FROM   mast.pttype
                     WHERE  insurance_type = '2'
                            AND supple_flag <> 'Y')
                                                                               -- a.pt_type身分別 in (pttype身份代碼檔 INSURANCE保險區分=健保 AND SUPPLE_FLAG待補卡/單註記 不等於 代補卡)
       AND a.div_no != '60'                                                    --排除中醫
       AND A.clinic_date = B.clinic_date
       AND F.clinic_date = B.clinic_date
       AND A.duplicate_no = B.duplicate_no
       AND A.chart_no = F.chart_no
       AND A.duplicate_no = F.duplicate_no
       AND a.nh_apply_flag IN ( 'Y', 'U' )
                                                                               -- nh_apply_flag申報否 U:尚未審核 Y:確定申報 N:暫不申報 X:永不申報
       AND B.nh_apply_flag = 'Y'
                                                                               -- nh_apply_flag申報否  Y:申報 N:轉下月申報 X:不申報
       AND a.nh_apply_type = '1'
                                                                               --申報類別 1:送核 2:補報 3:福山補報 9:專案補報 A:整筆就醫補報(A) a:部份醫令補報(a) B:整筆就醫補報(B) b:部份醫令補報(b) ....
       AND A.merge_flag NOT IN ( 'D', 'd' )
                                                                               --同療程合併註記 A:合併後新資料 D:合併後刪除資料
       AND B.pricing_flag NOT IN ( 'S', 's', 'Y', 'y',
                                   'h', 'X', 'x', 'Z' )
                                                                               --計價註記 Y/N/S,s/H,h/X,x/Z/V  Y:自費計價 N:健保計價申報 H:健保不計價申報(user KEYIN) h:健保不計價申報(電腦依price,joint,C1,05案件設定) S:任何身份皆自費(user KEYIN) s:任何身份皆自費(電腦依joint設定) X:不計價不申報(user KEYIN) x:不計價不申報(電腦依price,joint設定) Z:自費病人自費,健保病人不申報不計價 V:虛醫令,交付調劑之藥品空針
       AND A.div_no NOT IN ( 'EA', '2B' )
                                                                               --科別 NOT IN 居家照護(EA)及洗腎(2B)
       AND E.acnt_type = '2'
                                                                               --費用科目類別 0明細科目 1收據科目 2健保門診 3健保住院 P程式科目 R報表科目 T報表科目2
       AND b.ref_no = ' '                                                      --單據號碼 : 檢驗申請單,檢查申請單etc
       AND D.map_onh_acnt_no NOT IN ( '010' )
                                                                               --對應門診健保科目 不等於血液透析費(010)
                                                                               --AND B.APPLY_ORDER_TYPE IN ('1') --申報醫令類別 1-藥品,2-其他,3-材料,4-內含項,ex:論病例計酬,洗腎,復健,中醫藥品內含項目
                                                                               --//AND A.APPLY_CASE_TYPE NOT IN ('B1','B6','A1','A2','AZ','B7','B9','DF','A3','D2','C4','E1','03')
       AND B.apply_order_type IN ( '1' )
                                                                               --B1_1
       AND A.apply_case_type NOT IN ( 'B1', 'B6', 'B7', 'B8',
                                      'B9', 'BA', 'DF', 'C5' )
                                                                               --B1_5
       AND A.apply_case_type NOT IN ( 'A1', 'A2', 'A5', 'A6', 'A7' )
       AND ( A.apply_case_type NOT IN 'E1'
              OR ( A.apply_case_type IN 'E1'
                   AND Trim(A.special_exam_item_1) != 'EC'
                   AND Trim(A.special_exam_item_2) != 'EC'
                   AND Trim(A.special_exam_item_3) != 'EC'
                   AND Trim(A.special_exam_item_4) != 'EC' ) )
                                                                               --B1_8
       AND A.apply_case_type NOT IN ( 'C4' )
                                                                               --C1_8
       AND ( A.special_exam_item_1 NOT IN ( 'JA', 'JB' )
              OR A.special_exam_item_2 NOT IN ( 'JA', 'JB' )
              OR A.special_exam_item_3 NOT IN ( 'JA', 'JB' )
              OR A.special_exam_item_4 NOT IN ( 'JA', 'JB' ) )
                                                                               --D1_1
       AND Trim(A.special_exam_item_1) != 'G8'
                                                                               --D1_2
       AND A.apply_case_type NOT IN '03'
                                                                               --C_1 >C1_1
                                                                               --AND (a.SERIOUS_DISEASE_FLAG NOT IN 'Y'
                                                                               --              OR (A.SERIOUS_DISEASE_FLAG  IN 'Y'
                                                                               --                 AND  TRIM(A.NH_BURDEN_CODE) != '001'
                                                                               --                )
                                                                               --     )
                                                                               --C_3 > C1_3
                                                                               --排除BC肝
       AND B.nh_code NOT IN ( 'A043302100', 'A044650100', 'A048027100',
                              'A048152100',
                              'AC44650100', 'AC48027100', 'AC57322100',
                              'B016536209',
                              'B016536216', 'B016536221', 'B016536229',
                              'B016536299',
                              'B016955216', 'B016955220', 'B016955227',
                              'B016955237',
                              'B022178216', 'B022178220', 'B022178223',
                              'B022178227',
                              'B022178237', 'B022490216', 'B022490237',
                              'B022491221',
                              'B022491229', 'B023208100', 'B023920100',
                              'B024468100',
                              'B024469100', 'B024662100', 'B024690100',
                              'B025219166',
                              'B025232100', 'BC23208100', 'K000589237',
                              'K000591243',
                              'K000667255', 'K000669248', 'K000674253',
                              'K000675257',
                              'K000696266', 'K000700216', 'K000700220',
                              'K000700223',
                              'K000700227', 'K000752221', 'K000753216',
                              'K000754209',
                              'K000765209', 'K000766209', 'K000788277',
                              'K000789277',
                              'K000815248', 'K000816255', 'K000817257',
                              'K000818253',
                              'A049918100', 'AB48027100', 'AB57268100',
                              'AC57268100',
                              'AC57786100', 'AC57856100', 'AC57923100',
                              'B025605100',
                              'AC58039100', 'BC25232100', 'X000130216',
                              'AC49918100',
                              'AC58063100', 'AC58091100', 'K000929277',
                              'K000930277',
                              'AB57923100', 'BC26310100', 'AC43302100',
                              'AC58334100',
                              'AC58335100', 'BA24468100', 'BA24469100',
                              'BC23920100',
                              'BC24662100', 'BC24690100', 'BC25219166',
                              'KC00589237',
                              'KC00667255', 'KC00669248', 'KC00674253',
                              'KC00675257',
                              'KC00788277', 'KC00789277', 'KC00815248',
                              'KC00816255',
                              'KC00817257', 'KC00818253', 'KC00929277',
                              'AC58541100',
                              'AC58542100', 'AB57786100', 'AC48152100',
                              'BC25605100',
                              'KC00700216', 'KC00871277', 'KC00872277',
                              'AA57786100',
                              'AB58091100', 'KC00930277', 'AA57322100',
                              'AC58778100',
                              'KC00974277', 'KC00975277', 'AA58091100',
                              'AA58542100',
                              'AC58369100', 'BC26716100', 'BC26717100',
                              'AC59253100',
                              'AC59378100', 'BC26944100', 'BC26945100',
                              'AC59621100',
                              'BC27042100', 'AC59671100', 'AC59704100',
                              'AC59723100',
                              'BC27086100', 'AC58600100', 'AC59750100',
                              'AC59872100',
                              'AC59895100', 'AC60109100', 'AC60178100',
                              'AC60206100',
                              'AC60219100', 'BC27375100', 'BC27700100',
                              'BC27701100',
                                  'BC27829100' )
                                                                               --排除c肝
       AND B.nh_code NOT IN ( 'HCVDAA0001', 'HCVDAA0002', 'HCVDAA0003',
                              'HCVDAA0004',
                              'HCVDAA0005', 'HCVDAA0006', 'HCVDAA0007',
                              'HCVDAA0008',
                              'HCVDAA0009', 'HCVDAA0010', 'HCVDAA0011',
                              'HCVDAA0012',
                              'HCVDAA0013', 'HCVDAA0014', 'HCVDAA0015',
                              'HCVDAA0016',
                                  'HCVDAA0017' )
                                                                               --申報案件分類 B1:代辦性病患者全面篩檢愛滋病毒計、B6：職災案件、A1:居家照護、A2:精神疾病社區復健、AZ：職業傷病住院膳食費、B7：代辦門診戒菸、B9：代辦孕婦全面篩檢愛滋計畫、DF代辦登革熱 NS1 抗原快速診斷試劑、A3:預防保健、D2:代辦65歲以上老人流行性感冒疫苗接種、C4:代辦無健保結核病患就醫案件、E1：支付制度試辦計畫、03:西醫門診手術
       AND F.seq IN ( '00' )                                                   --序號 00為主診斷
       AND Substr (F.disease_code, 1, 1) NOT IN ( 'C' )
                                                                               --疾病代碼（診斷碼） 取開頭第一個字，但是不等於C開頭的
GROUP  BY A.clinic_date,
          A.chart_no,
          A.pt_name,
          C.doctor_name,
          G.div_full_name,
          B.nh_code,
          D.map_onh_acnt_no,
          E.acnt_full_name,
          F.seq,
          F.disease_code,
          B.code,
          a.clinic_flag,
          h.er_acnt_close_flag,
          a.building_no 
'''


# In[12]:


#6-7.YR門診非藥費人頭數
#確認SQL
if sqlmonth1 in [1,4,7,10]:
    sql7_1 = '''
    SELECT                                                                     --非藥費 1091123更新
    A.chart_no,                                                                --AS "病歷號",
    A.building_no                                                              --AS "院區"
    FROM   onh{0}{1}.acntopd B
           left join onh{0}{1}.ptopd A
                  ON A.chart_no = B.chart_no
                     AND B.clinic_date = a.clinic_date
                     AND A.duplicate_no = B.duplicate_no
           left join onh{0}{1}.pter g
                  ON A.chart_no = g.chart_no
                     AND A.clinic_date = g.clinic_date
                     AND A.duplicate_no = g.duplicate_no
           left join mast.doctor C
                  ON A.doctor_no = C.doctor_no
           left join mast.acntmap D
                  ON B.acnt_no_1 = D.acnt_no
           left join mast.acntname E
                  ON D.map_onh_acnt_no = E.acnt_no
           left join mast.div F
                  ON C.div_no = F.div_no
    WHERE  E.acnt_type = '2'                                                   --費用科目類別 = 2健保門診
           AND a.div_no != '60'                                                --排除中醫
           AND B.apply_order_type NOT IN ( '1', '4' )
                                                                               --醫令類別排除藥品'1'及內含項目'4' 1-藥品,2-其他,3-材料,4-內含項,ex:論病例計酬,洗腎,復健,中醫藥品內含項目
                                                                               --一般性排除
           AND A.merge_flag NOT IN 'D'
                                                                               --排除同療程合併註記->D:合併後刪除資料
                                                                               --AND A.NH_APPLY_TYPE NOT IN ('2')                                  --申報類別排除補報'2'
           AND ( B.nh_amt_1 != 0.00
                  OR B.nh_amt_2 != 0.00 )
           AND b.nh_apply_flag IN( 'Y' )                                       --Y:申報 N:轉下月申報 X:不申報
           AND A.apply_case_type NOT IN ( 'A3', 'D2', 'C1' )
                                                                               --排除B類案件(代辦業務)
                                                                               --B1_1 案件類別
           AND A.apply_case_type NOT IN ( 'B1', 'B6', 'B7', 'B8',
                                          'B9', 'BA', 'DF','C5' )
                                                                               --B1_2
                                                                               --愛滋病--無
                                                                               --B1_3
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( '11', '12', '13', '15',
                                              '16', '17', '19', '20',
                                              '71', '72', '73', '75',
                                              '76', '77', '79', '31',
                                              '33', '35', '37', '91',
                                              '93', '21', '22', '23',
                                              '24', '25', '26', '27',
                                              '28', '85', '95', '97', 'L1001C' ) ) )
                                                                               --B1_4後半
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( 'E4003C', 'E4004C', 'E4005C' ) ) )
                                                                               --B1_4
                                                                               --AND (A.APPLY_CASE_TYPE NOT IN ('D2')
                                                                               --     OR ( A.APPLY_CASE_TYPE IN ('D2')
                                                                               --    AND B.NH_CODE NOT IN ('A2001C','KH1N2C','A2051C','A2052C')
                                                                               --        )
                                                                               --    )
                                                                               --B1_5
           AND A.apply_case_type NOT IN ( 'A1', 'A2', 'A5', 'A6', 'A7' )
           AND ( A.apply_case_type NOT IN 'E1'
                  OR ( A.apply_case_type IN 'E1'
                       AND Trim(A.special_exam_item_1) != 'EC'
                       AND Trim(A.special_exam_item_2) != 'EC'
                       AND Trim(A.special_exam_item_3) != 'EC'
                       AND Trim(A.special_exam_item_4) != 'EC' ) )
                                                                               --B1_8
           AND A.apply_case_type NOT IN ( 'C4' )
                                                                               --C1_5前半
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( 'P1407C', 'P1408C', 'P1409C', 'P1410C'
                                              ,
                                              'P1411C', 'P1612C', 'P1613C', 'P1614B'
                                              ,
                                              'P1615C', 'P1801C', 'P1802C', 'P1803C'
                                              ,
                                              'P3402C', 'P3403C', 'P3404C', 'P3405C'
                                              ,
                                              'P3406C', 'P3407C', 'P3408C', 'P3409C'
                                              ,
                                              'P3903C', 'P3904C', 'P3905C', 'P4201C'
                                              ,
                                              'P4202C', 'P4203C', 'P4204C', 'P4205C'
                                              ,
                                              'P4301C', 'P4302C', 'P4303C', 'P5113B'
                                              ,
                                              'P5114B', 'P5115B', 'P5116B', 'P5117B'
                                              ,
                                              'P5118B', 'P5123B', 'P5124B', 'P5125B'
                                              ,
                                              'P5501B', 'P5502B', 'P5503B', 'P5504B'
                                              ,
                                              'P5505B', 'P5510B', 'P5511B', 'P5512B'
                                              ,
                                              'P5513B', 'P5514B', 'P5515B', 'P5516B'
                                              ,
                                              'P5517B', 'P5301C', 'P3410C', 'P3411C'
                                              ,
                                              'P6011C', 'P6012C', 'P6013C', 'P6014C'
                                              ,
                                              'P6015C', 'P5506B', 'P5507B', 'P5508B'
                                              ,
                                              'P5509B', 'P5126B', 'P5127B', 'P5128B'
                                              ,
                                              'P5132C', 'P5135B' ) ) )
                                                                               --C1_5後半
           AND ( B.apply_order_type NOT IN ( '0' )
                  OR ( B.apply_order_type IN ( '0' )
                       AND B.nh_code NOT IN ( 'P5201C', 'P5202C', 'P5203C', 'P5204C'
                                            ) ) )
                                                                               --C1_7
           AND ( B.apply_order_type NOT IN ( '2', 'X', 'Z', 'K' )
                  OR ( B.apply_order_type IN ( '2', 'X', 'Z', 'K' )
                       AND B.nh_code NOT IN ( 'P4601B', 'P4602B', 'P4603B', 'P4604B'
                                              ,
                                              'P4605B', 'P4606B', 'P4607B', 'P4608B'
                                              ,
                                              'P4609B', 'P4610B', 'P4611B', 'P4612B'
                                              ,
                                              'P4613B', 'P4614B', 'P4615B', 'P4616B'
                                              ,
                                              'P4617B', 'P4618B' ) ) )
                                                                               --C1_8
           AND ( A.special_exam_item_1 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_2 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_3 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_4 NOT IN ( 'JA', 'JB' ) )
                                                                               --D1_1
           AND ( A.apply_case_type NOT IN 'E1'
                  OR ( A.apply_case_type NOT IN 'E1'
                       AND Trim(A.special_exam_item_1) != 'G8' ) )
                                                                               --D1_2
           AND A.apply_case_type NOT IN '03'
                                                                               --D1_4
           AND B.nh_code NOT IN ( '05221A' )
                                                                               --排除門診血液透析費
           AND ( A.clinic_flag NOT IN ( 'O', 'V', 'H', 'F', 'P' )
                  OR ( A.clinic_flag IN ( 'O', 'V', 'H', 'F', 'P' )
                       AND D.map_onh_acnt_no NOT IN ( '010' ) ) )
           AND ( B.ref_no = ' '                                                --單據號碼是一個空格
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no BETWEEN
                                                  '040' AND '048')
                                                                               --明細科目的對應程式科目 040-048
                       AND B.ref_no IN (SELECT apply_no
                                        FROM   xray.ptexamindex
                                        WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                      )
                  OR ( B.acnt_no_1 IN (SELECT acnt_no
                                       FROM   mast.acntmap
                                       WHERE  map_prog_acnt_no BETWEEN
                                              '120' AND '128')
                                                                               --明細科目的對應程式科目 120-128
                       AND ( B.ref_no ) IN (SELECT apply_no
                                            FROM   exam.ptexamindex
                                            WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                      )
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no IN (
                                                  '050', '056', '058'
                                                                      ))
                                                                               --明細科目的對應程式科目 050, 056, 058
                       AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                       FROM   lab.labx1
                                                       WHERE
                           report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"細菌初步報告表頭"檢驗單號 確認有檢查
                      )
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no IN ( '055' ))
                                                                               --明細科目的對應程式科目 055
                       AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                       FROM   lam.labx1
                                                       WHERE
                           report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"???LAM?"檢驗單號 確認有檢查
                      ) )
    
    GROUP  BY A.chart_no,
              A.building_no

    '''
    sql7_1 = sql7_1.format(sqlyear, sqlmonth7_1)   
    sql7 = sql7_1
elif sqlmonth1 in [2,5,8,11]:
    sql7_2 = '''
    SELECT                                                                     --非藥費 1091123更新
    A.chart_no,                                                                --AS "病歷號",
    A.building_no                                                              --AS "院區"
    FROM   onh{0}{1}.acntopd B
           left join onh{0}{1}.ptopd A
                  ON A.chart_no = B.chart_no
                     AND B.clinic_date = a.clinic_date
                     AND A.duplicate_no = B.duplicate_no
           left join onh{0}{1}.pter g
                  ON A.chart_no = g.chart_no
                     AND A.clinic_date = g.clinic_date
                     AND A.duplicate_no = g.duplicate_no
           left join mast.doctor C
                  ON A.doctor_no = C.doctor_no
           left join mast.acntmap D
                  ON B.acnt_no_1 = D.acnt_no
           left join mast.acntname E
                  ON D.map_onh_acnt_no = E.acnt_no
           left join mast.div F
                  ON C.div_no = F.div_no
    WHERE  E.acnt_type = '2'                                                   --費用科目類別 = 2健保門診
           AND a.div_no != '60'                                                --排除中醫
           AND B.apply_order_type NOT IN ( '1', '4' )
                                                                               --醫令類別排除藥品'1'及內含項目'4' 1-藥品,2-其他,3-材料,4-內含項,ex:論病例計酬,洗腎,復健,中醫藥品內含項目
                                                                               --一般性排除
           AND A.merge_flag NOT IN 'D'
                                                                               --排除同療程合併註記->D:合併後刪除資料
                                                                               --AND A.NH_APPLY_TYPE NOT IN ('2')   --申報類別排除補報'2'
           AND ( B.nh_amt_1 != 0.00
                  OR B.nh_amt_2 != 0.00 )
           AND b.nh_apply_flag IN( 'Y' )                                       --Y:申報 N:轉下月申報 X:不申報
           AND A.apply_case_type NOT IN ( 'A3', 'D2', 'C1' )
                                                                               --排除B類案件(代辦業務)
                                                                               --B1_1 案件類別
           AND A.apply_case_type NOT IN ( 'B1', 'B6', 'B7', 'B8',
                                          'B9', 'BA', 'DF','C5' )
                                                                               --B1_2
                                                                               --愛滋病--無
                                                                               --B1_3
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( '11', '12', '13', '15',
                                              '16', '17', '19', '20',
                                              '71', '72', '73', '75',
                                              '76', '77', '79', '31',
                                              '33', '35', '37', '91',
                                              '93', '21', '22', '23',
                                              '24', '25', '26', '27',
                                              '28', '85', '95', '97', 'L1001C' ) ) )
                                                                               --B1_4後半
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( 'E4003C', 'E4004C', 'E4005C' ) ) )
                                                                               --B1_4
                                                                               --AND (A.APPLY_CASE_TYPE NOT IN ('D2')
                                                                               --     OR ( A.APPLY_CASE_TYPE IN ('D2')
                                                                               --    AND B.NH_CODE NOT IN ('A2001C','KH1N2C','A2051C','A2052C')
                                                                               --        )
                                                                               --    )
                                                                               --B1_5
           AND A.apply_case_type NOT IN ( 'A1', 'A2', 'A5', 'A6', 'A7' )
           AND ( A.apply_case_type NOT IN 'E1'
                  OR ( A.apply_case_type IN 'E1'
                       AND Trim(A.special_exam_item_1) != 'EC'
                       AND Trim(A.special_exam_item_2) != 'EC'
                       AND Trim(A.special_exam_item_3) != 'EC'
                       AND Trim(A.special_exam_item_4) != 'EC' ) )
                                                                               --B1_8
           AND A.apply_case_type NOT IN ( 'C4' )
                                                                               --C1_5前半
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( 'P1407C', 'P1408C', 'P1409C', 'P1410C'
                                              ,
                                              'P1411C', 'P1612C', 'P1613C', 'P1614B'
                                              ,
                                              'P1615C', 'P1801C', 'P1802C', 'P1803C'
                                              ,
                                              'P3402C', 'P3403C', 'P3404C', 'P3405C'
                                              ,
                                              'P3406C', 'P3407C', 'P3408C', 'P3409C'
                                              ,
                                              'P3903C', 'P3904C', 'P3905C', 'P4201C'
                                              ,
                                              'P4202C', 'P4203C', 'P4204C', 'P4205C'
                                              ,
                                              'P4301C', 'P4302C', 'P4303C', 'P5113B'
                                              ,
                                              'P5114B', 'P5115B', 'P5116B', 'P5117B'
                                              ,
                                              'P5118B', 'P5123B', 'P5124B', 'P5125B'
                                              ,
                                              'P5501B', 'P5502B', 'P5503B', 'P5504B'
                                              ,
                                              'P5505B', 'P5510B', 'P5511B', 'P5512B'
                                              ,
                                              'P5513B', 'P5514B', 'P5515B', 'P5516B'
                                              ,
                                              'P5517B', 'P5301C', 'P3410C', 'P3411C'
                                              ,
                                              'P6011C', 'P6012C', 'P6013C', 'P6014C'
                                              ,
                                              'P6015C', 'P5506B', 'P5507B', 'P5508B'
                                              ,
                                              'P5509B', 'P5126B', 'P5127B', 'P5128B'
                                              ,
                                              'P5132C', 'P5135B' ) ) )
                                                                               --C1_5後半
           AND ( B.apply_order_type NOT IN ( '0' )
                  OR ( B.apply_order_type IN ( '0' )
                       AND B.nh_code NOT IN ( 'P5201C', 'P5202C', 'P5203C', 'P5204C'
                                            ) ) )
                                                                               --C1_7
           AND ( B.apply_order_type NOT IN ( '2', 'X', 'Z', 'K' )
                  OR ( B.apply_order_type IN ( '2', 'X', 'Z', 'K' )
                       AND B.nh_code NOT IN ( 'P4601B', 'P4602B', 'P4603B', 'P4604B'
                                              ,
                                              'P4605B', 'P4606B', 'P4607B', 'P4608B'
                                              ,
                                              'P4609B', 'P4610B', 'P4611B', 'P4612B'
                                              ,
                                              'P4613B', 'P4614B', 'P4615B', 'P4616B'
                                              ,
                                              'P4617B', 'P4618B' ) ) )
                                                                               --C1_8
           AND ( A.special_exam_item_1 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_2 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_3 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_4 NOT IN ( 'JA', 'JB' ) )
                                                                               --D1_1
           AND ( A.apply_case_type NOT IN 'E1'
                  OR ( A.apply_case_type NOT IN 'E1'
                       AND Trim(A.special_exam_item_1) != 'G8' ) )
                                                                               --D1_2
           AND A.apply_case_type NOT IN '03'
                                                                               --D1_4
           AND B.nh_code NOT IN ( '05221A' )
                                                                               --排除門診血液透析費
           AND ( A.clinic_flag NOT IN ( 'O', 'V', 'H', 'F', 'P' )
                  OR ( A.clinic_flag IN ( 'O', 'V', 'H', 'F', 'P' )
                       AND D.map_onh_acnt_no NOT IN ( '010' ) ) )
           AND ( B.ref_no = ' '                                                --單據號碼是一個空格
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no BETWEEN
                                                  '040' AND '048')
                                                                               --明細科目的對應程式科目 040-048
                       AND B.ref_no IN (SELECT apply_no
                                        FROM   xray.ptexamindex
                                        WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                      )
                  OR ( B.acnt_no_1 IN (SELECT acnt_no
                                       FROM   mast.acntmap
                                       WHERE  map_prog_acnt_no BETWEEN
                                              '120' AND '128')
                                                                               --明細科目的對應程式科目 120-128
                       AND ( B.ref_no ) IN (SELECT apply_no
                                            FROM   exam.ptexamindex
                                            WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                      )
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no IN (
                                                  '050', '056', '058'
                                                                      ))
                                                                               --明細科目的對應程式科目 050, 056, 058
                       AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                       FROM   lab.labx1
                                                       WHERE
                           report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"細菌初步報告表頭"檢驗單號 確認有檢查
                      )
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no IN ( '055' ))
                                                                               --明細科目的對應程式科目 055
                       AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                       FROM   lam.labx1
                                                       WHERE
                           report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"???LAM?"檢驗單號 確認有檢查
                      ) )
    
    GROUP  BY A.chart_no,
              A.building_no
    
              UNION ALL
    
    SELECT                                                                     --非藥費 1091123更新
    A.chart_no,
    A.building_no                                                              --AS "院區"
    FROM   onh{0}{2}.acntopd B
           left join onh{0}{2}.ptopd A
                  ON A.chart_no = B.chart_no
                     AND B.clinic_date = a.clinic_date
                     AND A.duplicate_no = B.duplicate_no
           left join onh{0}{2}.pter g
                  ON A.chart_no = g.chart_no
                     AND A.clinic_date = g.clinic_date
                     AND A.duplicate_no = g.duplicate_no
           left join mast.doctor C
                  ON A.doctor_no = C.doctor_no
           left join mast.acntmap D
                  ON B.acnt_no_1 = D.acnt_no
           left join mast.acntname E
                  ON D.map_onh_acnt_no = E.acnt_no
           left join mast.div F
                  ON C.div_no = F.div_no
    WHERE  E.acnt_type = '2'                                                   --費用科目類別 = 2健保門診
           AND a.div_no != '60'                                                --排除中醫
           AND B.apply_order_type NOT IN ( '1', '4' )
                                                                               --醫令類別排除藥品'1'及內含項目'4' 1-藥品,2-其他,3-材料,4-內含項,ex:論病例計酬,洗腎,復健,中醫藥品內含項目
                                                                               --一般性排除
           AND A.merge_flag NOT IN 'D'
                                                                               --排除同療程合併註記->D:合併後刪除資料
                                                                               --AND A.NH_APPLY_TYPE NOT IN ('2')   --申報類別排除補報'2'
           AND ( B.nh_amt_1 != 0.00
                  OR B.nh_amt_2 != 0.00 )
           AND b.nh_apply_flag IN( 'Y' )                                       --Y:申報 N:轉下月申報 X:不申報
           AND A.apply_case_type NOT IN ( 'A3', 'D2', 'C1' )
                                                                               --排除B類案件(代辦業務)
                                                                               --B1_1 案件類別
           AND A.apply_case_type NOT IN ( 'B1', 'B6', 'B7', 'B8',
                                          'B9', 'BA', 'DF','C5'  )
                                                                               --B1_2
                                                                               --愛滋病--無
                                                                               --B1_3
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                    AND B.nh_code NOT IN ( '11', '12', '13', '15',
                                              '16', '17', '19', '20',
                                              '71', '72', '73', '75',
                                              '76', '77', '79', '31',
                                              '33', '35', '37', '91',
                                              '93', '21', '22', '23',
                                              '24', '25', '26', '27',
                                              '28', '85', '95', '97', 'L1001C',
                                              '41','50','61','63','64','65','66','68','69'
                                              ,'4A','4B','4C','4D','4E','6A','6B','6C','6D','P3501C','P3502C','P3503C'
                                              ,'21','22','23','24','25','26','27','28','21+L1001C','25+L1001C','27+L1001C','98','99' ) ) )
                                                                               --B1_4後半
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( 'E4003C', 'E4004C', 'E4005C' ) ) )
                                                                               --B1_4
                                                                               --      AND (A.APPLY_CASE_TYPE NOT IN ('D2')
                                                                               --          OR ( A.APPLY_CASE_TYPE IN ('D2')
                                                                               --         AND B.NH_CODE NOT IN ('A2001C','KH1N2C','A2051C','A2052C','A3001C','E5002C','E5003C','E5004C','E5005C')
                                                                               --             )
                                                                               --         )
                                                                               --B1_5
           AND A.apply_case_type NOT IN ( 'A1', 'A2', 'A5', 'A6', 'A7' )
           AND ( A.apply_case_type NOT IN 'E1'
                  OR ( A.apply_case_type IN 'E1'
                       AND Trim(A.special_exam_item_1) != 'EC'
                       AND Trim(A.special_exam_item_2) != 'EC'
                       AND Trim(A.special_exam_item_3) != 'EC'
                       AND Trim(A.special_exam_item_4) != 'EC' ) )
                                                                               --B1_8
           AND A.apply_case_type NOT IN ( 'C4' )
                                                                               --C1_5前半
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( 'P1407C', 'P1408C', 'P1409C', 'P1410C'
                                              ,
                                              'P1411C', 'P1612C', 'P1613C', 'P1614B'
                                              ,
                                              'P1615C', 'P1801C', 'P1802C', 'P1803C'
                                              ,
                                              'P3402C', 'P3403C', 'P3404C', 'P3405C'
                                              ,
                                              'P3406C', 'P3407C', 'P3408C', 'P3409C'
                                              ,
                                              'P3903C', 'P3904C', 'P3905C', 'P4201C'
                                              ,
                                              'P4202C', 'P4203C', 'P4204C', 'P4205C'
                                              ,
                                              'P4301C', 'P4302C', 'P4303C', 'P5113B'
                                              ,
                                              'P5114B', 'P5115B', 'P5116B', 'P5117B'
                                              ,
                                              'P5118B', 'P5123B', 'P5124B', 'P5125B'
                                              ,
                                              'P5501B', 'P5502B', 'P5503B', 'P5504B'
                                              ,
                                              'P5505B', 'P5510B', 'P5511B', 'P5512B'
                                              ,
                                              'P5513B', 'P5514B', 'P5515B', 'P5516B'
                                              ,
                                              'P5517B', 'P5301C', 'P3410C', 'P3411C'
                                              ,
                                              'P6011C', 'P6012C', 'P6013C', 'P6014C'
                                              ,
                                              'P6015C', 'P5506B', 'P5507B', 'P5508B'
                                              ,
                                              'P5509B', 'P5126B', 'P5127B', 'P5128B'
                                              ,
                                              'P5132C', 'P5135B' ) ) )
                                                                               --C1_5後半
           AND ( B.apply_order_type NOT IN ( '0' )
                  OR ( B.apply_order_type IN ( '0' )
                       AND B.nh_code NOT IN ( 'P5201C', 'P5202C', 'P5203C', 'P5204C'
                                            ) ) )
                                                                               --C1_7
           AND ( B.apply_order_type NOT IN ( '2', 'X', 'Z', 'K' )
                  OR ( B.apply_order_type IN ( '2', 'X', 'Z', 'K' )
                       AND B.nh_code NOT IN ( 'P4601B', 'P4602B', 'P4603B', 'P4604B'
                                              ,
                                              'P4605B', 'P4606B', 'P4607B', 'P4608B'
                                              ,
                                              'P4609B', 'P4610B', 'P4611B', 'P4612B'
                                              ,
                                              'P4613B', 'P4614B', 'P4615B', 'P4616B'
                                              ,
                                              'P4617B', 'P4618B' ) ) )
                                                                               --C1_8
           AND ( A.special_exam_item_1 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_2 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_3 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_4 NOT IN ( 'JA', 'JB' ) )
                                                                               --D1_1
           AND ( A.apply_case_type NOT IN 'E1'
                  OR ( A.apply_case_type NOT IN 'E1'
                       AND Trim(A.special_exam_item_1) != 'G8' ) )
                                                                               --D1_2
           AND A.apply_case_type NOT IN '03'
                                                                               --D1_4
           AND B.nh_code NOT IN ( '05221A' )
                                                                               --排除門診血液透析費
           AND ( A.clinic_flag NOT IN ( 'O', 'V', 'H', 'F', 'P' )
                  OR ( A.clinic_flag IN ( 'O', 'V', 'H', 'F', 'P' )
                       AND D.map_onh_acnt_no NOT IN ( '010' ) ) )
           AND ( B.ref_no = ' '                                                --單據號碼是一個空格
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no BETWEEN
                                                  '040' AND '048')
                                                                               --明細科目的對應程式科目 040-048
                       AND B.ref_no IN (SELECT apply_no
                                        FROM   xray.ptexamindex
                                        WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                      )
                  OR ( B.acnt_no_1 IN (SELECT acnt_no
                                       FROM   mast.acntmap
                                       WHERE  map_prog_acnt_no BETWEEN
                                              '120' AND '128')
                                                                               --明細科目的對應程式科目 120-128
                       AND ( B.ref_no ) IN (SELECT apply_no
                                            FROM   exam.ptexamindex
                                            WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                      )
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no IN (
                                                  '050', '056', '058'
                                                                      ))
                                                                               --明細科目的對應程式科目 050, 056, 058
                       AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                       FROM   lab.labx1
                                                       WHERE
                           report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"細菌初步報告表頭"檢驗單號 確認有檢查
                      )
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no IN ( '055' ))
                                                                               --明細科目的對應程式科目 055
                       AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                       FROM   lam.labx1
                                                       WHERE
                           report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"???LAM?"檢驗單號 確認有檢查
                      ) )
    
    GROUP  BY A.chart_no,
              A.building_no
    '''
    sql7_2 = sql7_2.format(sqlyear, sqlmonth7_1, sqlmonth7_2)
    sql7 = sql7_2
elif sqlmonth1 in [3,6,9,12]:
    sql7_3 = '''    
    SELECT                                                                     --非藥費 1091123更新
    A.chart_no,                                                                --AS "病歷號",
    A.building_no                                                              --AS "院區"
    FROM   onh{0}{1}.acntopd B
           left join onh{0}{1}.ptopd A
                  ON A.chart_no = B.chart_no
                     AND B.clinic_date = a.clinic_date
                     AND A.duplicate_no = B.duplicate_no
           left join onh{0}{1}.pter g
                  ON A.chart_no = g.chart_no
                     AND A.clinic_date = g.clinic_date
                     AND A.duplicate_no = g.duplicate_no
           left join mast.doctor C
                  ON A.doctor_no = C.doctor_no
           left join mast.acntmap D
                  ON B.acnt_no_1 = D.acnt_no
           left join mast.acntname E
                  ON D.map_onh_acnt_no = E.acnt_no
           left join mast.div F
                  ON C.div_no = F.div_no
    WHERE  E.acnt_type = '2'                                                   --費用科目類別 = 2健保門診
           AND a.div_no != '60'                                                --排除中醫
           AND B.apply_order_type NOT IN ( '1', '4' )
                                                                               --醫令類別排除藥品'1'及內含項目'4' 1-藥品,2-其他,3-材料,4-內含項,ex:論病例計酬,洗腎,復健,中醫藥品內含項目
                                                                               --一般性排除
           AND A.merge_flag NOT IN 'D'
                                                                               --排除同療程合併註記->D:合併後刪除資料
                                                                               --AND A.NH_APPLY_TYPE NOT IN ('2')   --申報類別排除補報'2'
           AND ( B.nh_amt_1 != 0.00
                  OR B.nh_amt_2 != 0.00 )
           AND b.nh_apply_flag IN( 'Y' )                                       --Y:申報 N:轉下月申報 X:不申報
           AND A.apply_case_type NOT IN ( 'A3', 'D2', 'C1' )
                                                                               --排除B類案件(代辦業務)
                                                                               --B1_1 案件類別
           AND A.apply_case_type NOT IN ( 'B1', 'B6', 'B7', 'B8',
                                          'B9', 'BA', 'DF','C5' )
                                                                               --B1_2
                                                                               --愛滋病--無
                                                                               --B1_3
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( '11', '12', '13', '15',
                                              '16', '17', '19', '20',
                                              '71', '72', '73', '75',
                                              '76', '77', '79', '31',
                                              '33', '35', '37', '91',
                                              '93', '21', '22', '23',
                                              '24', '25', '26', '27',
                                              '28', '85', '95', '97', 'L1001C' ) ) )
                                                                               --B1_4後半
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( 'E4003C', 'E4004C', 'E4005C' ) ) )
                                                                               --B1_4
                                                                               --AND (A.APPLY_CASE_TYPE NOT IN ('D2')
                                                                               --     OR ( A.APPLY_CASE_TYPE IN ('D2')
                                                                               --    AND B.NH_CODE NOT IN ('A2001C','KH1N2C','A2051C','A2052C')
                                                                               --        )
                                                                               --    )
                                                                               --B1_5
           AND A.apply_case_type NOT IN ( 'A1', 'A2', 'A5', 'A6', 'A7' )
           AND ( A.apply_case_type NOT IN 'E1'
                  OR ( A.apply_case_type IN 'E1'
                       AND Trim(A.special_exam_item_1) != 'EC'
                       AND Trim(A.special_exam_item_2) != 'EC'
                       AND Trim(A.special_exam_item_3) != 'EC'
                       AND Trim(A.special_exam_item_4) != 'EC' ) )
                                                                               --B1_8
           AND A.apply_case_type NOT IN ( 'C4' )
                                                                               --C1_5前半
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( 'P1407C', 'P1408C', 'P1409C', 'P1410C'
                                              ,
                                              'P1411C', 'P1612C', 'P1613C', 'P1614B'
                                              ,
                                              'P1615C', 'P1801C', 'P1802C', 'P1803C'
                                              ,
                                              'P3402C', 'P3403C', 'P3404C', 'P3405C'
                                              ,
                                              'P3406C', 'P3407C', 'P3408C', 'P3409C'
                                              ,
                                              'P3903C', 'P3904C', 'P3905C', 'P4201C'
                                              ,
                                              'P4202C', 'P4203C', 'P4204C', 'P4205C'
                                              ,
                                              'P4301C', 'P4302C', 'P4303C', 'P5113B'
                                              ,
                                              'P5114B', 'P5115B', 'P5116B', 'P5117B'
                                              ,
                                              'P5118B', 'P5123B', 'P5124B', 'P5125B'
                                              ,
                                              'P5501B', 'P5502B', 'P5503B', 'P5504B'
                                              ,
                                              'P5505B', 'P5510B', 'P5511B', 'P5512B'
                                              ,
                                              'P5513B', 'P5514B', 'P5515B', 'P5516B'
                                              ,
                                              'P5517B', 'P5301C', 'P3410C', 'P3411C'
                                              ,
                                              'P6011C', 'P6012C', 'P6013C', 'P6014C'
                                              ,
                                              'P6015C', 'P5506B', 'P5507B', 'P5508B'
                                              ,
                                              'P5509B', 'P5126B', 'P5127B', 'P5128B'
                                              ,
                                              'P5132C', 'P5135B' ) ) )
                                                                               --C1_5後半
           AND ( B.apply_order_type NOT IN ( '0' )
                  OR ( B.apply_order_type IN ( '0' )
                       AND B.nh_code NOT IN ( 'P5201C', 'P5202C', 'P5203C', 'P5204C'
                                            ) ) )
                                                                               --C1_7
           AND ( B.apply_order_type NOT IN ( '2', 'X', 'Z', 'K' )
                  OR ( B.apply_order_type IN ( '2', 'X', 'Z', 'K' )
                       AND B.nh_code NOT IN ( 'P4601B', 'P4602B', 'P4603B', 'P4604B'
                                              ,
                                              'P4605B', 'P4606B', 'P4607B', 'P4608B'
                                              ,
                                              'P4609B', 'P4610B', 'P4611B', 'P4612B'
                                              ,
                                              'P4613B', 'P4614B', 'P4615B', 'P4616B'
                                              ,
                                              'P4617B', 'P4618B' ) ) )
                                                                               --C1_8
           AND ( A.special_exam_item_1 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_2 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_3 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_4 NOT IN ( 'JA', 'JB' ) )
                                                                               --D1_1
           AND ( A.apply_case_type NOT IN 'E1'
                  OR ( A.apply_case_type NOT IN 'E1'
                       AND Trim(A.special_exam_item_1) != 'G8' ) )
                                                                               --D1_2
           AND A.apply_case_type NOT IN '03'
                                                                               --D1_4
           AND B.nh_code NOT IN ( '05221A' )
                                                                               --排除門診血液透析費
           AND ( A.clinic_flag NOT IN ( 'O', 'V', 'H', 'F', 'P' )
                  OR ( A.clinic_flag IN ( 'O', 'V', 'H', 'F', 'P' )
                       AND D.map_onh_acnt_no NOT IN ( '010' ) ) )
           AND ( B.ref_no = ' '                                                --單據號碼是一個空格
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no BETWEEN
                                                  '040' AND '048')
                                                                               --明細科目的對應程式科目 040-048
                       AND B.ref_no IN (SELECT apply_no
                                        FROM   xray.ptexamindex
                                        WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                      )
                  OR ( B.acnt_no_1 IN (SELECT acnt_no
                                       FROM   mast.acntmap
                                       WHERE  map_prog_acnt_no BETWEEN
                                              '120' AND '128')
                                                                               --明細科目的對應程式科目 120-128
                       AND ( B.ref_no ) IN (SELECT apply_no
                                            FROM   exam.ptexamindex
                                            WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                      )
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no IN (
                                                  '050', '056', '058'
                                                                      ))
                                                                               --明細科目的對應程式科目 050, 056, 058
                       AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                       FROM   lab.labx1
                                                       WHERE
                           report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"細菌初步報告表頭"檢驗單號 確認有檢查
                      )
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no IN ( '055' ))
                                                                               --明細科目的對應程式科目 055
                       AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                       FROM   lam.labx1
                                                       WHERE
                           report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"???LAM?"檢驗單號 確認有檢查
                      ) )
    
    GROUP  BY A.chart_no,
              A.building_no
    
              UNION ALL
    
    SELECT                                                                     --非藥費 1091123更新
    A.chart_no,
    A.building_no AS "院區"
    FROM   onh{0}{2}.acntopd B
           left join onh{0}{2}.ptopd A
                  ON A.chart_no = B.chart_no
                     AND B.clinic_date = a.clinic_date
                     AND A.duplicate_no = B.duplicate_no
           left join onh{0}{2}.pter g
                  ON A.chart_no = g.chart_no
                     AND A.clinic_date = g.clinic_date
                     AND A.duplicate_no = g.duplicate_no
           left join mast.doctor C
                  ON A.doctor_no = C.doctor_no
           left join mast.acntmap D
                  ON B.acnt_no_1 = D.acnt_no
           left join mast.acntname E
                  ON D.map_onh_acnt_no = E.acnt_no
           left join mast.div F
                  ON C.div_no = F.div_no
    WHERE  E.acnt_type = '2'                                                   --費用科目類別 = 2健保門診
           AND a.div_no != '60' --排除中醫
           AND B.apply_order_type NOT IN ( '1', '4' )
                                                                               --醫令類別排除藥品'1'及內含項目'4' 1-藥品,2-其他,3-材料,4-內含項,ex:論病例計酬,洗腎,復健,中醫藥品內含項目
                                                                               --一般性排除
           AND A.merge_flag NOT IN 'D'
                                                                               --排除同療程合併註記->D:合併後刪除資料
                                                                               --AND A.NH_APPLY_TYPE NOT IN ('2')   --申報類別排除補報'2'
           AND ( B.nh_amt_1 != 0.00
                  OR B.nh_amt_2 != 0.00 )
           AND b.nh_apply_flag IN( 'Y' )                                       --Y:申報 N:轉下月申報 X:不申報
           AND A.apply_case_type NOT IN ( 'A3', 'D2', 'C1' )
                                                                               --排除B類案件(代辦業務)
                                                                               --B1_1 案件類別
           AND A.apply_case_type NOT IN ( 'B1', 'B6', 'B7', 'B8',
                                          'B9', 'BA', 'DF','C5'  )
                                                                               --B1_2
                                                                               --愛滋病--無
                                                                               --B1_3
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                    AND B.nh_code NOT IN ( '11', '12', '13', '15',
                                              '16', '17', '19', '20',
                                              '71', '72', '73', '75',
                                              '76', '77', '79', '31',
                                              '33', '35', '37', '91',
                                              '93', '21', '22', '23',
                                              '24', '25', '26', '27',
                                              '28', '85', '95', '97', 'L1001C',
                                              '41','50','61','63','64','65','66','68','69'
                                              ,'4A','4B','4C','4D','4E','6A','6B','6C','6D','P3501C','P3502C','P3503C'
                                              ,'21','22','23','24','25','26','27','28','21+L1001C','25+L1001C','27+L1001C','98','99' ) ) )
                                                                               --B1_4後半
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( 'E4003C', 'E4004C', 'E4005C' ) ) )
                                                                               --B1_4
                                                                               --      AND (A.APPLY_CASE_TYPE NOT IN ('D2')
                                                                               --          OR ( A.APPLY_CASE_TYPE IN ('D2')
                                                                               --         AND B.NH_CODE NOT IN ('A2001C','KH1N2C','A2051C','A2052C','A3001C','E5002C','E5003C','E5004C','E5005C')
                                                                               --             )
                                                                               --         )
                                                                               --B1_5
           AND A.apply_case_type NOT IN ( 'A1', 'A2', 'A5', 'A6', 'A7' )
           AND ( A.apply_case_type NOT IN 'E1'
                  OR ( A.apply_case_type IN 'E1'
                       AND Trim(A.special_exam_item_1) != 'EC'
                       AND Trim(A.special_exam_item_2) != 'EC'
                       AND Trim(A.special_exam_item_3) != 'EC'
                       AND Trim(A.special_exam_item_4) != 'EC' ) )
                                                                               --B1_8
           AND A.apply_case_type NOT IN ( 'C4' )
                                                                               --C1_5前半
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( 'P1407C', 'P1408C', 'P1409C', 'P1410C'
                                              ,
                                              'P1411C', 'P1612C', 'P1613C', 'P1614B'
                                              ,
                                              'P1615C', 'P1801C', 'P1802C', 'P1803C'
                                              ,
                                              'P3402C', 'P3403C', 'P3404C', 'P3405C'
                                              ,
                                              'P3406C', 'P3407C', 'P3408C', 'P3409C'
                                              ,
                                              'P3903C', 'P3904C', 'P3905C', 'P4201C'
                                              ,
                                              'P4202C', 'P4203C', 'P4204C', 'P4205C'
                                              ,
                                              'P4301C', 'P4302C', 'P4303C', 'P5113B'
                                              ,
                                              'P5114B', 'P5115B', 'P5116B', 'P5117B'
                                              ,
                                              'P5118B', 'P5123B', 'P5124B', 'P5125B'
                                              ,
                                              'P5501B', 'P5502B', 'P5503B', 'P5504B'
                                              ,
                                              'P5505B', 'P5510B', 'P5511B', 'P5512B'
                                              ,
                                              'P5513B', 'P5514B', 'P5515B', 'P5516B'
                                              ,
                                              'P5517B', 'P5301C', 'P3410C', 'P3411C'
                                              ,
                                              'P6011C', 'P6012C', 'P6013C', 'P6014C'
                                              ,
                                              'P6015C', 'P5506B', 'P5507B', 'P5508B'
                                              ,
                                              'P5509B', 'P5126B', 'P5127B', 'P5128B'
                                              ,
                                              'P5132C', 'P5135B' ) ) )
                                                                               --C1_5後半
           AND ( B.apply_order_type NOT IN ( '0' )
                  OR ( B.apply_order_type IN ( '0' )
                       AND B.nh_code NOT IN ( 'P5201C', 'P5202C', 'P5203C', 'P5204C'
                                            ) ) )
                                                                               --C1_7
           AND ( B.apply_order_type NOT IN ( '2', 'X', 'Z', 'K' )
                  OR ( B.apply_order_type IN ( '2', 'X', 'Z', 'K' )
                       AND B.nh_code NOT IN ( 'P4601B', 'P4602B', 'P4603B', 'P4604B'
                                              ,
                                              'P4605B', 'P4606B', 'P4607B', 'P4608B'
                                              ,
                                              'P4609B', 'P4610B', 'P4611B', 'P4612B'
                                              ,
                                              'P4613B', 'P4614B', 'P4615B', 'P4616B'
                                              ,
                                              'P4617B', 'P4618B' ) ) )
                                                                               --C1_8
           AND ( A.special_exam_item_1 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_2 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_3 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_4 NOT IN ( 'JA', 'JB' ) )
                                                                               --D1_1
           AND ( A.apply_case_type NOT IN 'E1'
                  OR ( A.apply_case_type NOT IN 'E1'
                       AND Trim(A.special_exam_item_1) != 'G8' ) )
                                                                               --D1_2
           AND A.apply_case_type NOT IN '03'
                                                                               --D1_4
           AND B.nh_code NOT IN ( '05221A' )
                                                                               --排除門診血液透析費
           AND ( A.clinic_flag NOT IN ( 'O', 'V', 'H', 'F', 'P' )
                  OR ( A.clinic_flag IN ( 'O', 'V', 'H', 'F', 'P' )
                       AND D.map_onh_acnt_no NOT IN ( '010' ) ) )
           AND ( B.ref_no = ' '                                                --單據號碼是一個空格
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no BETWEEN
                                                  '040' AND '048')
                                                                               --明細科目的對應程式科目 040-048
                       AND B.ref_no IN (SELECT apply_no
                                        FROM   xray.ptexamindex
                                        WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                      )
                  OR ( B.acnt_no_1 IN (SELECT acnt_no
                                       FROM   mast.acntmap
                                       WHERE  map_prog_acnt_no BETWEEN
                                              '120' AND '128')
                                                                               --明細科目的對應程式科目 120-128
                       AND ( B.ref_no ) IN (SELECT apply_no
                                            FROM   exam.ptexamindex
                                            WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                      )
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no IN (
                                                  '050', '056', '058'
                                                                      ))
                                                                               --明細科目的對應程式科目 050, 056, 058
                       AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                       FROM   lab.labx1
                                                       WHERE
                           report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"細菌初步報告表頭"檢驗單號 確認有檢查
                      )
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no IN ( '055' ))
                                                                               --明細科目的對應程式科目 055
                       AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                       FROM   lam.labx1
                                                       WHERE
                           report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"???LAM?"檢驗單號 確認有檢查
                      ) )
    
    GROUP  BY A.chart_no,
              A.building_no
    
              UNION ALL
    
    SELECT                                                                     --非藥費 1091123更新
    A.chart_no,
    A.building_no                                                              -- AS "院區"
    FROM   onh{0}{3}.acntopd B
           left join onh{0}{3}.ptopd A
                  ON A.chart_no = B.chart_no
                     AND B.clinic_date = a.clinic_date
                     AND A.duplicate_no = B.duplicate_no
           left join onh{0}{3}.pter g
                  ON A.chart_no = g.chart_no
                     AND A.clinic_date = g.clinic_date
                     AND A.duplicate_no = g.duplicate_no
           left join mast.doctor C
                  ON A.doctor_no = C.doctor_no
           left join mast.acntmap D
                  ON B.acnt_no_1 = D.acnt_no
           left join mast.acntname E
                  ON D.map_onh_acnt_no = E.acnt_no
           left join mast.div F
                  ON C.div_no = F.div_no
    WHERE  E.acnt_type = '2'                                                   --費用科目類別 = 2健保門診
           AND a.div_no != '60'                                                --排除中醫
           AND B.apply_order_type NOT IN ( '1', '4' )
                                                                               --醫令類別排除藥品'1'及內含項目'4' 1-藥品,2-其他,3-材料,4-內含項,ex:論病例計酬,洗腎,復健,中醫藥品內含項目
                                                                               --一般性排除
           AND A.merge_flag NOT IN 'D'
                                                                               --排除同療程合併註記->D:合併後刪除資料
                                                                               --AND A.NH_APPLY_TYPE NOT IN ('2')   --申報類別排除補報'2'
           AND ( B.nh_amt_1 != 0.00
                  OR B.nh_amt_2 != 0.00 )
           AND b.nh_apply_flag IN( 'Y' )                                       --Y:申報 N:轉下月申報 X:不申報
           AND A.apply_case_type NOT IN ( 'A3', 'D2', 'C1' )
                                                                               --排除B類案件(代辦業務)
                                                                               --B1_1 案件類別
           AND A.apply_case_type NOT IN ( 'B1', 'B6', 'B7', 'B8',
                                          'B9', 'BA', 'DF','C5'  )
                                                                               --B1_2
                                                                               --愛滋病--無
                                                                               --B1_3
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( '11', '12', '13', '15',
                                              '16', '17', '19', '20',
                                              '71', '72', '73', '75',
                                              '76', '77', '79', '31',
                                              '33', '35', '37', '91',
                                              '93', '21', '22', '23',
                                              '24', '25', '26', '27',
                                              '28', '85', '95', '97', 'L1001C' ) ) )
                                                                               --B1_4後半
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( 'E4003C', 'E4004C', 'E4005C' ) ) )
                                                                               --B1_4
                                                                               --AND (A.APPLY_CASE_TYPE NOT IN ('D2')
                                                                               --     OR ( A.APPLY_CASE_TYPE IN ('D2')
                                                                               --    AND B.NH_CODE NOT IN ('A2001C','KH1N2C','A2051C','A2052C')
                                                                               --        )
                                                                               --    )
                                                                               --B1_5
           AND A.apply_case_type NOT IN ( 'A1', 'A2', 'A5', 'A6', 'A7' )
           AND ( A.apply_case_type NOT IN 'E1'
                  OR ( A.apply_case_type IN 'E1'
                       AND Trim(A.special_exam_item_1) != 'EC'
                       AND Trim(A.special_exam_item_2) != 'EC'
                       AND Trim(A.special_exam_item_3) != 'EC'
                       AND Trim(A.special_exam_item_4) != 'EC' ) )
                                                                               --B1_8
           AND A.apply_case_type NOT IN ( 'C4' )
                                                                               --C1_5前半
           AND ( B.apply_order_type NOT IN '2'
                  OR ( B.apply_order_type IN '2'
                       AND B.nh_code NOT IN ( 'P1407C', 'P1408C', 'P1409C', 'P1410C'
                                              ,
                                              'P1411C', 'P1612C', 'P1613C', 'P1614B'
                                              ,
                                              'P1615C', 'P1801C', 'P1802C', 'P1803C'
                                              ,
                                              'P3402C', 'P3403C', 'P3404C', 'P3405C'
                                              ,
                                              'P3406C', 'P3407C', 'P3408C', 'P3409C'
                                              ,
                                              'P3903C', 'P3904C', 'P3905C', 'P4201C'
                                              ,
                                              'P4202C', 'P4203C', 'P4204C', 'P4205C'
                                              ,
                                              'P4301C', 'P4302C', 'P4303C', 'P5113B'
                                              ,
                                              'P5114B', 'P5115B', 'P5116B', 'P5117B'
                                              ,
                                              'P5118B', 'P5123B', 'P5124B', 'P5125B'
                                              ,
                                              'P5501B', 'P5502B', 'P5503B', 'P5504B'
                                              ,
                                              'P5505B', 'P5510B', 'P5511B', 'P5512B'
                                              ,
                                              'P5513B', 'P5514B', 'P5515B', 'P5516B'
                                              ,
                                              'P5517B', 'P5301C', 'P3410C', 'P3411C'
                                              ,
                                              'P6011C', 'P6012C', 'P6013C', 'P6014C'
                                              ,
                                              'P6015C', 'P5506B', 'P5507B', 'P5508B'
                                              ,
                                              'P5509B', 'P5126B', 'P5127B', 'P5128B'
                                              ,
                                              'P5132C', 'P5135B' ) ) )
                                                                               --C1_5後半
           AND ( B.apply_order_type NOT IN ( '0' )
                  OR ( B.apply_order_type IN ( '0' )
                       AND B.nh_code NOT IN ( 'P5201C', 'P5202C', 'P5203C', 'P5204C'
                                            ) ) )
                                                                               --C1_7
           AND ( B.apply_order_type NOT IN ( '2', 'X', 'Z', 'K' )
                  OR ( B.apply_order_type IN ( '2', 'X', 'Z', 'K' )
                       AND B.nh_code NOT IN ( 'P4601B', 'P4602B', 'P4603B', 'P4604B'
                                              ,
                                              'P4605B', 'P4606B', 'P4607B', 'P4608B'
                                              ,
                                              'P4609B', 'P4610B', 'P4611B', 'P4612B'
                                              ,
                                              'P4613B', 'P4614B', 'P4615B', 'P4616B'
                                              ,
                                              'P4617B', 'P4618B' ) ) )
                                                                               --C1_8
           AND ( A.special_exam_item_1 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_2 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_3 NOT IN ( 'JA', 'JB' )
                  OR A.special_exam_item_4 NOT IN ( 'JA', 'JB' ) )
                                                                               --D1_1
           AND ( A.apply_case_type NOT IN 'E1'
                  OR ( A.apply_case_type NOT IN 'E1'
                       AND Trim(A.special_exam_item_1) != 'G8' ) )
                                                                               --D1_2
           AND A.apply_case_type NOT IN '03'
                                                                               --D1_4
           AND B.nh_code NOT IN ( '05221A' )
                                                                               --排除門診血液透析費
           AND ( A.clinic_flag NOT IN ( 'O', 'V', 'H', 'F', 'P' )
                  OR ( A.clinic_flag IN ( 'O', 'V', 'H', 'F', 'P' )
                       AND D.map_onh_acnt_no NOT IN ( '010' ) ) )
           AND ( B.ref_no = ' '                                                --單據號碼是一個空格
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no BETWEEN
                                                  '040' AND '048')
                                                                               --明細科目的對應程式科目 040-048
                       AND B.ref_no IN (SELECT apply_no
                                        FROM   xray.ptexamindex
                                        WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                      )
                  OR ( B.acnt_no_1 IN (SELECT acnt_no
                                       FROM   mast.acntmap
                                       WHERE  map_prog_acnt_no BETWEEN
                                              '120' AND '128')
                                                                               --明細科目的對應程式科目 120-128
                       AND ( B.ref_no ) IN (SELECT apply_no
                                            FROM   exam.ptexamindex
                                            WHERE  status > '2')
                                                                               --由單據號碼 串到 申請單索引檔 確認有檢查
                      )
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no IN (
                                                  '050', '056', '058'
                                                                      ))
                                                                               --明細科目的對應程式科目 050, 056, 058
                       AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                       FROM   lab.labx1
                                                       WHERE
                           report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"細菌初步報告表頭"檢驗單號 確認有檢查
                      )
                  OR ( B.ref_no <> ' '
                       AND B.acnt_no_1 IN (SELECT acnt_no
                                           FROM   mast.acntmap
                                           WHERE  map_prog_acnt_no IN ( '055' ))
                                                                               --明細科目的對應程式科目 055
                       AND Substr(B.ref_no, 6, 10) IN (SELECT lab_no
                                                       FROM   lam.labx1
                                                       WHERE
                           report_date <> '0000000')
                                                                               --由單據號碼後幾碼串到"???LAM?"檢驗單號 確認有檢查
                      ) )
    
    GROUP  BY A.chart_no,
              A.building_no
    '''
    sql7_3 = sql7_3.format(sqlyear, sqlmonth7_1, sqlmonth7_2, sqlmonth7_3)
    sql7 = sql7_3


# In[13]:


#6-8.YR門診藥費人頭數
if sqlmonth1 in [1,4,7,10]:
    sql8_1 = '''
    SELECT
    A.chart_no,
    A.building_no
                                                                               --select a.nh_case_type,a.nh_seq_no,B.NH_CASE_NAME,A.NH_APPLY_FLAG
    FROM   onh{0}{1}.ptopd a
           LEFT JOIN onh.onhcasetype B
                  ON A.nh_case_type = B.nh_case_type
    WHERE  a.nh_apply_type = '1'
                                                                               --and a.NH_SEQ_NO = '1'
           AND a.nh_case_type IN ( '02', '04', 'A3', '06',
                                   '08', 'C1', '09', 'E1' )
           AND A.med_no NOT IN ( ' ', '0' )
    GROUP  BY A.chart_no, A.building_no
    '''
    sql8_1 = sql8_1.format(sqlyear, sqlmonth8_1)   
    sql8 = sql8_1
elif sqlmonth1 in [2,5,8,11]:
    sql8_2 = '''
    SELECT
    A.chart_no,
    A.building_no
                                                                               --select a.nh_case_type,a.nh_seq_no,B.NH_CASE_NAME,A.NH_APPLY_FLAG
    FROM   onh{0}{1}.ptopd a
           LEFT JOIN onh.onhcasetype B
                  ON A.nh_case_type = B.nh_case_type
    WHERE  a.nh_apply_type = '1'
                                                                               --and a.NH_SEQ_NO = '1'
           AND a.nh_case_type IN ( '02', '04', 'A3', '06',
                                   '08', 'C1', '09', 'E1' )
           AND A.med_no NOT IN ( ' ', '0' )
    GROUP  BY A.chart_no, A.building_no
    
    UNION ALL
    
    SELECT
    A.chart_no,
    A.building_no
                                                                               --select a.nh_case_type,a.nh_seq_no,B.NH_CASE_NAME,A.NH_APPLY_FLAG
    FROM   onh{0}{2}.ptopd a
           LEFT JOIN onh.onhcasetype B
                  ON A.nh_case_type = B.nh_case_type
    WHERE  a.nh_apply_type = '1'
                                                                               --and a.NH_SEQ_NO = '1'
           AND a.nh_case_type IN ( '02', '04', 'A3', '06',
                                   '08', 'C1', '09', 'E1' )
           AND A.med_no NOT IN ( ' ', '0' )
    GROUP  BY A.chart_no, A.building_no
    '''
    sql8_2 = sql8_2.format(sqlyear, sqlmonth8_1, sqlmonth8_2)
    sql8 = sql8_2
elif sqlmonth1 in [3,6,9,12]:
    sql8_3 = '''    
    SELECT
    A.chart_no,
    A.building_no
                                                                               --select a.nh_case_type,a.nh_seq_no,B.NH_CASE_NAME,A.NH_APPLY_FLAG
    FROM   onh{0}{1}.ptopd a
           LEFT JOIN onh.onhcasetype B
                  ON A.nh_case_type = B.nh_case_type
    WHERE  a.nh_apply_type = '1'
                                                                               --and a.NH_SEQ_NO = '1'
           AND a.nh_case_type IN ( '02', '04', 'A3', '06',
                                   '08', 'C1', '09', 'E1' )
           AND A.med_no NOT IN ( ' ', '0' )
    GROUP  BY A.chart_no, A.building_no
    
    UNION ALL
    
    SELECT
    A.chart_no,
    A.building_no
                                                                               --select a.nh_case_type,a.nh_seq_no,B.NH_CASE_NAME,A.NH_APPLY_FLAG
    FROM   onh{0}{2}.ptopd a
           LEFT JOIN onh.onhcasetype B
                  ON A.nh_case_type = B.nh_case_type
    WHERE  a.nh_apply_type = '1'
                                                                               --and a.NH_SEQ_NO = '1'
           AND a.nh_case_type IN ( '02', '04', 'A3', '06',
                                   '08', 'C1', '09', 'E1' )
           AND A.med_no NOT IN ( ' ', '0' )
    GROUP  BY A.chart_no, A.building_no
    
    UNION ALL
    
    SELECT
    A.chart_no,
    A.building_no
                                                                               --select a.nh_case_type,a.nh_seq_no,B.NH_CASE_NAME,A.NH_APPLY_FLAG
    FROM   onh{0}{3}.ptopd a
           LEFT JOIN onh.onhcasetype B
                  ON A.nh_case_type = B.nh_case_type
    WHERE  a.nh_apply_type = '1'
                                                                               --and a.NH_SEQ_NO = '1'
           AND a.nh_case_type IN ( '02', '04', 'A3', '06',
                                   '08', 'C1', '09', 'E1' )
           AND A.med_no NOT IN ( ' ', '0' )
    GROUP  BY A.chart_no, A.building_no
    '''
    sql8_3 = sql8_3.format(sqlyear, sqlmonth8_1, sqlmonth8_2, sqlmonth8_3)
    sql8 = sql8_3


# In[14]:


#6-9.YR住院非藥費人次數
sql9 = '''
                                                                               --住院非藥費
        SELECT nh_apply_yyymm,                                                 --AS "申報年月",
               id_no                 ,                                         --AS "申報的身份証號",
               ( total_amt - amt_13 ),                                         --AS "非藥費",
               CASE Substr(bed_no, 0, 1)
                 WHEN 'B' THEN 'B'
                 ELSE 'A'
               END                   bed_no                                    --AS "院區"
        FROM   inh{2}.dtlfb @hi
        WHERE  nh_apply_yyymm = '{1}{2}'
               AND nh_case_type NOT IN( 'AZ', 'DZ','C5' )
               AND tw_drg_unused_flag NOT IN ( '0', '1', '2', '3',
                                               '4', '5', '6', '9',
                                               'B', 'F', 'G', 'J',
                                               'K', 'L' )
               AND emg_bed_days <= 60
        
        UNION ALL
        SELECT nh_apply_yyymm,                                                 --AS "申報年月",
               id_no,                                                          --AS "申報的身份証號",
               ( total_amt - amt_13 ),                                         --AS "非藥費",
               CASE Substr(bed_no, 0, 1)
                 WHEN 'B' THEN 'B'
                 ELSE 'A'
               END                   bed_no                                    --AS "院區"
        FROM   inh{3}.dtlfb @hi
        WHERE  nh_apply_yyymm = '{1}{3}'
               AND nh_case_type NOT IN( 'AZ', 'DZ','C5' )
               AND tw_drg_unused_flag NOT IN ( '0', '1', '2', '3',
                                               '4', '5', '6', '9',
                                               'B', 'F', 'G', 'J',
                                               'K', 'L' )
               AND emg_bed_days <= 60
        
        
        UNION ALL
        SELECT nh_apply_yyymm,                                                 --AS "申報年月",
               id_no,                                                          --AS "申報的身份証號",
               ( total_amt - amt_13 ),                                         --AS "非藥費",
               CASE Substr(bed_no, 0, 1)
                 WHEN 'B' THEN 'B'
                 ELSE 'A'
               END                   bed_no                                    --AS "院區"
        FROM   inh{4}.dtlfb @hi
        WHERE  nh_apply_yyymm = '{1}{4}'
               AND nh_case_type NOT IN( 'AZ', 'DZ','C5' )
               AND tw_drg_unused_flag NOT IN ( '0', '1', '2', '3',
                                               '4', '5', '6', '9',
                                               'B', 'F', 'G', 'J',
                                               'K', 'L' )
               AND emg_bed_days <= 60
'''


# In[15]:


#6-10.YR住院藥費人頭
sql10 = '''
                                                                               --住院藥費
        SELECT a.nh_apply_yyymm,                                               --AS "申報年月",
               a.id_no,                                                        --AS "申報的身份証號",
               a.amt_13,                                                       --AS "藥費",
               CASE Substr(bed_no, 0, 1)
                 WHEN 'B' THEN 'B'
                 ELSE 'A'
               END             bed_no                                          --AS "院區"
        FROM   inh{2}.dtlfb @hi a
        WHERE  a.nh_apply_yyymm = '{1}{2}'
               AND A.nh_case_type NOT IN ( 'A1', 'A2', 'A3', 'A4',
                                           'C1', '7', 'AZ', 'DZ',
                                           '5', 'C2', 'C3', 'C4' ,'C5')
               AND ( A.nh_paid_type != '9' )
               AND a.amt_13 != '0'
        UNION ALL
        SELECT a.nh_apply_yyymm,                                               --AS "申報年月",
               a.id_no,                                                        --AS "申報的身份証號",
               a.amt_13,                                                       --AS "藥費",
               CASE Substr(bed_no, 0, 1)
                 WHEN 'B' THEN 'B'
                 ELSE 'A'
               END             bed_no                                          --AS "院區"
        FROM   inh{2}.dtlfb @hi a
        WHERE  a.nh_apply_yyymm = '{1}{2}'
               AND A.nh_case_type NOT IN ( 'A1', 'A2', 'A3', 'A4',
                                           'C1', '7', 'AZ', 'DZ',
                                           '5', 'C2', 'C3', 'C4' ,'C5' )
               AND ( A.nh_paid_type = '9'
                     AND nh_case_type = '1' )
               AND a.amt_13 != '0'
        
        UNION ALL
        
        SELECT a.nh_apply_yyymm,                                               --AS "申報年月",
               a.id_no,                                                        --AS "申報的身份証號",
               a.amt_13,                                                       --AS "藥費",
               CASE Substr(bed_no, 0, 1)
                 WHEN 'B' THEN 'B'
                 ELSE 'A'
               END             bed_no                                          --AS "院區"
        FROM   inh{3}.dtlfb @hi a
        WHERE  a.nh_apply_yyymm = '{1}{3}'
               AND A.nh_case_type NOT IN ( 'A1', 'A2', 'A3', 'A4',
                                           'C1', '7', 'AZ', 'DZ',
                                           '5', 'C2', 'C3', 'C4' ,'C5' )
               AND ( A.nh_paid_type != '9' )
               AND a.amt_13 != '0'
        UNION ALL
        SELECT a.nh_apply_yyymm,                                               --AS "申報年月",
               a.id_no,                                                        --AS "申報的身份証號",
               a.amt_13,                                                       --AS "藥費",
               CASE Substr(bed_no, 0, 1)
                 WHEN 'B' THEN 'B'
                 ELSE 'A'
               END             bed_no                                          --AS "院區"
        FROM   inh{3}.dtlfb @hi a
        WHERE  a.nh_apply_yyymm = '{1}{3}'
               AND A.nh_case_type NOT IN ( 'A1', 'A2', 'A3', 'A4',
                                           'C1', '7', 'AZ', 'DZ',
                                           '5', 'C2', 'C3', 'C4' ,'C5' )
               AND ( A.nh_paid_type = '9'
                     AND nh_case_type = '1' )
               AND a.amt_13 != '0'
        
        UNION ALL
        SELECT a.nh_apply_yyymm,                                               --AS "申報年月",
               a.id_no,                                                        --AS "申報的身份証號",
               a.amt_13,                                                       --AS "藥費",
               CASE Substr(bed_no, 0, 1)
                 WHEN 'B' THEN 'B'
                 ELSE 'A'
               END             bed_no                                          --AS "院區"
        FROM   inh{4}.dtlfb @hi a
        WHERE  a.nh_apply_yyymm = '{1}{4}'
               AND A.nh_case_type NOT IN ( 'A1', 'A2', 'A3', 'A4',
                                           'C1', '7', 'AZ', 'DZ',
                                           '5', 'C2', 'C3', 'C4' ,'C5' )
               AND ( A.nh_paid_type != '9' )
               AND a.amt_13 != '0'
        UNION ALL
        SELECT a.nh_apply_yyymm,                                               --AS "申報年月",
               a.id_no,                                                        --AS "申報的身份証號",
               a.amt_13,                                                       --AS "藥費",
               CASE Substr(bed_no, 0, 1)
                 WHEN 'B' THEN 'B'
                 ELSE 'A'
               END             bed_no                                          --AS "院區"
        FROM   inh{4}.dtlfb @hi a
        WHERE  a.nh_apply_yyymm = '{1}{4}'
               AND A.nh_case_type NOT IN ( 'A1', 'A2', 'A3', 'A4',
                                           'C1', '7', 'AZ', 'DZ',
                                           '5', 'C2', 'C3', 'C4' ,'C5' )
               AND ( A.nh_paid_type = '9'
                     AND nh_case_type = '1' )
               AND a.amt_13 != '0'
'''


# In[16]:


#6-11.住院BC肝及C肝口服
sql11 = '''
SELECT A.nh_apply_yyymm,                                                       --AS "申報年月"
       A.admit_no,                                                             --AS "住院編號"
       A.code,                                                                 --AS "院內碼"
       A.nh_code,                                                              --AS "健保醫令代碼"
       A.nh_amt,                                                               --AS "健保金額"
       CASE Substr(nh_bed_no, 0, 1)
         WHEN 'B' THEN 'B'
         ELSE 'A'
       END      nh_bed_no                                                      --AS "院區"
FROM   inh{2}.ordfb@hi A
WHERE  A.code IN ( 'OHCVD1', 'OHCVD2', 'OHCVD3', 'OHCVD4',
                   'OHCVD5', 'OHCVD6', 'OHCVD7', 'OHCVD8',
                   'OHCVD9', 'OHCVD10', 'OHCVD11', 'OHCVD12',
                   'OHCVD13', 'OHCVD14', 'OHCVD15', 'OHCVD16','OHCVD17',
                   'ORIBAVI', 'ORIBA', 'OBARA1', 'OENTE00',
                   'OTELB00', 'OTENOF', 'IPEGAS' )
       AND nh_apply_yyymm = '{1}{2}'

UNION ALL
SELECT A.nh_apply_yyymm,                                                       --AS "申報年月"
       A.admit_no,                                                             --AS "住院編號"
       A.code,                                                                 --AS "院內碼"
       A.nh_code,                                                              --AS "健保醫令代碼"
       A.nh_amt,                                                               --AS "健保金額"
       CASE Substr(nh_bed_no, 0, 1)
         WHEN 'B' THEN 'B'
         ELSE 'A'
       END      nh_bed_no                                                      --AS "院區"
FROM   inh{3}.ordfb@hi A
WHERE  A.code IN ( 'OHCVD1', 'OHCVD2', 'OHCVD3', 'OHCVD4',
                   'OHCVD5', 'OHCVD6', 'OHCVD7', 'OHCVD8',
                   'OHCVD9', 'OHCVD10', 'OHCVD11', 'OHCVD12',
                   'OHCVD13', 'OHCVD14', 'OHCVD15', 'OHCVD16','OHCVD17',
                   'ORIBAVI', 'ORIBA', 'OBARA1', 'OENTE00',
                   'OTELB00', 'OTENOF', 'IPEGAS' )
       AND nh_apply_yyymm = '{1}{3}'

UNION ALL
SELECT A.nh_apply_yyymm,                                                       --AS "申報年月"
       A.admit_no,                                                             --AS "住院編號"
       A.code,                                                                 --AS "院內碼"
       A.nh_code,                                                              --AS "健保醫令代碼"
       A.nh_amt,                                                               --AS "健保金額"
       CASE Substr(nh_bed_no, 0, 1)
         WHEN 'B' THEN 'B'
         ELSE 'A'
       END      nh_bed_no                                                      --AS "院區"
FROM   inh{4}.ordfb@hi A
WHERE  A.code IN ( 'OHCVD1', 'OHCVD2', 'OHCVD3', 'OHCVD4',
                   'OHCVD5', 'OHCVD6', 'OHCVD7', 'OHCVD8',
                   'OHCVD9', 'OHCVD10', 'OHCVD11', 'OHCVD12',
                   'OHCVD13', 'OHCVD14', 'OHCVD15', 'OHCVD16','OHCVD17',
                   'ORIBAVI', 'ORIBA', 'OBARA1', 'OENTE00',
                   'OTELB00', 'OTENOF', 'IPEGAS' )
       AND nh_apply_yyymm = '{1}{4}'
       '''


# In[17]:


#6-12.1110919YR-住院代辦案件預估(含C5)
sql12 = '''
--住院代辦案件

SELECT A.nh_apply_yyymm,        --"申報年月",
       A.id_no,                 --"申報的身份証號",
       A.total_amt,             --"申報金額",
       CASE Substr(A.bed_no, 0, 1)
         WHEN 'B' THEN 'B'
         ELSE 'A'
       END                  bed_no-- AS "院區"
FROM   inh{2}.dtlfb @hi A
WHERE  nh_apply_yyymm = '{1}{2}'
       AND nh_case_type  IN( 'A3','B1','B6','B7','C5','D2' )
       AND emg_bed_days <= 60
       
       
UNION ALL


SELECT nh_apply_yyymm        "申報年月",
       id_no                 "申報的身份証號",
       ( total_amt )          "申報金額",
       CASE Substr(bed_no, 0, 1)
         WHEN 'B' THEN 'B'
         ELSE 'A'
       END                   AS "院區"
FROM   inh{3}.dtlfb @hi
WHERE  nh_apply_yyymm = '{1}{3}'
       AND nh_case_type  IN( 'A3','B1','B6','B7','C5','D2' )
       AND emg_bed_days <= 60
       
       
UNION ALL


SELECT nh_apply_yyymm        "申報年月",
       id_no                 "申報的身份証號",
       ( total_amt )          "申報金額",
       CASE Substr(bed_no, 0, 1)
         WHEN 'B' THEN 'B'
         ELSE 'A'
       END                   AS "院區"
FROM   inh{4}.dtlfb @hi
WHERE  nh_apply_yyymm = '{1}{4}'
       AND nh_case_type  IN( 'A3','B1','B6','B7','C5','D2' )
       AND emg_bed_days <= 60
'''


# In[18]:


#7.sql字串格式化       
sql1 = sql1.format(sqlyear, sqlmonth)
sql2 = sql2.format(sqlyear, sqlmonth, year)
sql3 = sql3.format(sqlyear, sqlmonth)
sql4 = sql4.format(sqlyear, sqlmonth)
sql5 = sql5.format(sqlyear, sqlmonth)
sql6 = sql6.format(sqlyear, sqlmonth)
sql7 = sql7
sql8 = sql8
sql9  = sql9.format(sqlyear, sqlyear9, sqlmonth9_1, sqlmonth9_2, sqlmonth9_3)
sql10 = sql10.format(sqlyear, sqlyear10, sqlmonth10_1, sqlmonth10_2, sqlmonth10_3)
sql11 = sql11.format(sqlyear, sqlyear11, sqlmonth11_1, sqlmonth11_2, sqlmonth11_3)
sql12 = sql12.format(sqlyear, sqlyear12, sqlmonth12_1, sqlmonth12_2, sqlmonth12_3)


# In[19]:


#8.資料庫連線套件初始化
try:
    cx_Oracle.init_oracle_client(lib_dir=r"C:\oracle\instantclient_11_2")
except Exception as e:
    e_type = e.__class__.__name__
    print("例外錯誤型態: " + e_type)
    print("例外訊息: " + str(e))
    pass


# In[20]:


#9.資料庫連線
#註記:需根據連線字串與TNANAMES
connection = cx_Oracle.connect('XXXX/XXXX@xxx.xxx.xxx.xxx:xxxx/xx')
#9.讀取資料 
df1 = pd.read_sql(sql1, con=connection)
df2 = pd.read_sql(sql2, con=connection)
df3 = pd.read_sql(sql3, con=connection)
df4 = pd.read_sql(sql4, con=connection)
df5 = pd.read_sql(sql5, con=connection)
df6 = pd.read_sql(sql6, con=connection)
df7 = pd.read_sql(sql7, con=connection)
df8 = pd.read_sql(sql8, con=connection)
df9 = pd.read_sql(sql9, con=connection)
df10 = pd.read_sql(sql10, con=connection)
df11 = pd.read_sql(sql11, con=connection)
df12 = pd.read_sql(sql12, con=connection)
df12 = pd.read_sql(sql12, con=connection)


# In[21]:


#10.資料處理
#10-1.df1
#資料轉碼
df1copy = df1.copy()
column1 = ''
column2 = ''
column3 = ''
column4 = ''
df1column1 = []
df1column2 = []
df1column3 = []
df1column4 = []

for index, data in df1copy.iterrows():    
    column1 = binascii.unhexlify(data["PT_NAME"]).decode('big5','ignore').rstrip()
    column2 = binascii.unhexlify(data["DOCTOR_NAME"]).decode('big5','ignore').rstrip()
    column3 = binascii.unhexlify(data["DIV_FULL_NAME"]).decode('big5','ignore').rstrip()
    column4 = binascii.unhexlify(data["ACNT_FULL_NAME"]).decode('big5','ignore').rstrip()
    df1column1.append(column1)
    df1column2.append(column2)
    df1column3.append(column3)
    df1column4.append(column4)
df1copy["PT_NAME"] = pd.Series(df1column1)
df1copy["DOCTOR_NAME"] = pd.Series(df1column2)
df1copy["DIV_FULL_NAME"] = pd.Series(df1column3)
df1copy["ACNT_FULL_NAME"] = pd.Series(df1column4)

#欄位轉換中文
df1copy.rename(columns = {'CLINIC_DATE':'就醫日',
                          'CHART_NO':'病歷號',
                          'PT_NAME':'病人姓名',
                          'DOCTOR_NAME':'醫師姓名',
                          'DIV_FULL_NAME':'科別全名',
                          'CODE':'院內碼',
                          'NH_CODE':'健保碼',
                          'MAP_ONH_ACNT_NO':'健保科目',
                          'ACNT_FULL_NAME':'科目名稱',
                          '(B.NH_AMT_1+B.NH_AMT_2)':'健保點數',
                          'BUILDING_NO':'院區'
                          }, inplace = True)
#樞紐分析 
df1copysum = pd.DataFrame({'健保點數':[df1copy['健保點數'].astype("int").sum()]})

#轉出試算表
df1filename = '1.YR' + str(year) + str(month) + str(day) + '申報前門診申報點數' + '.xlsx'
with pd.ExcelWriter(df1filename) as writer:
    df1copy.to_excel(writer, index=False, sheet_name='Result')
    df1copysum.to_excel(writer, index=False, sheet_name='Pivot')


# In[22]:


#10-2.df2
#資料轉碼
df2copy = df2.copy()
column5 = ''
column6 = ''
column7 = ''
column8 = ''
column89 = ''
column9 = ''
df2column5 = []
df2column6 = []
df2column7 = []
df2column8 = []
df2column89 = []
df2column9 = []

for index, data in df2copy.iterrows():    
    column5 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(D.CLINIC_NO))"]).decode('big5','ignore').rstrip()
    column6 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(F.PT_TYPE_FULL_NAME))"]).decode('big5','ignore').rstrip()
    column7 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))"]).decode('big5','ignore').rstrip()
    column8 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(B.DIV_FULL_NAME))"]).decode('big5','ignore').rstrip()
    column89 = binascii.unhexlify(data["NH_APPLY_FLAG"]).decode('big5','ignore').rstrip()
    column9 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(H.ACNT_FULL_NAME))"]).decode('big5','ignore').rstrip()
    df2column5.append(column5)
    df2column6.append(column6)
    df2column7.append(column7)
    df2column8.append(column8)
    df2column89.append(column89)
    df2column9.append(column9)
    
df2copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(D.CLINIC_NO))"] = pd.Series(df2column5)
df2copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(F.PT_TYPE_FULL_NAME))"] = pd.Series(df2column6)
df2copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))"] = pd.Series(df2column7)
df2copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(B.DIV_FULL_NAME))"] = pd.Series(df2column8)
df2copy["NH_APPLY_FLAG"] = pd.Series(df2column89)
df2copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(H.ACNT_FULL_NAME))"] = pd.Series(df2column9)

#欄位轉換中文
df2copy.rename(columns = {'SUBSTR(A.CLINIC_DATE,0,5)':'年月',
                          'CLINIC_DATE':'就醫日期',
                          'CLINIC_APN':'午別',
                          'CLINIC_FLAG':'診別',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(D.CLINIC_NO))':'診間號',
                          'CHART_NO':'病歷號',
                          'CODE':'院內碼',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(F.PT_TYPE_FULL_NAME))':'身份別',
                          'DOCTOR_NO':'醫師代碼',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))':'醫師姓名',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(B.DIV_FULL_NAME))':'科別名稱',
                          'NH_CLINIC_SEQ':'健保就醫序號',
                          'TOTAL_QTY':'開立總量',
                          'NH_AMT_1':'健保金額',
                          'NH_APPLY_FLAG':'當月申報註記',
                          'MAP_RECEIPT_ACNT_NO':'收據科目',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(H.ACNT_FULL_NAME))':'收據科目名稱',
                          'APPLY_CASE_TYPE':'案件分類',
                          'BUILDING_NO':'院區'
                          }, inplace = True)
#樞紐分析 
df2copysum = pd.DataFrame({'健保金額':[df2copy['健保金額'].astype("int").sum()]})

#轉出試算表
df2filename = '2.' + str(year) + str(month) + str(day) + '-YR代辦案件預估' + '.xlsx'
with pd.ExcelWriter(df2filename) as writer:
    df2copy.to_excel(writer, index=False, sheet_name='Result')
    df2copysum.to_excel(writer, index=False, sheet_name='Pivot') 


# In[23]:


#10-3.df3
#資料轉碼
df3copy = df3.copy()
column10 = ''
column11 = ''
column12 = ''
column13 = ''
df1column10 = []
df1column11 = []
df1column12 = []
df1column13 = []

for index, data in df3copy.iterrows():    
    column10 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(A.PT_NAME))"]).decode('big5','ignore').rstrip()
    column11 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))"]).decode('big5','ignore').rstrip()
    column12 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(G.DIV_FULL_NAME))"]).decode('big5','ignore').rstrip()
    column13 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(E.ACNT_FULL_NAME))"]).decode('big5','ignore').rstrip()
    df1column10.append(column10)
    df1column11.append(column11)
    df1column12.append(column12)
    df1column13.append(column13)
    
df3copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(A.PT_NAME))"] = pd.Series(df1column10)
df3copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))"] = pd.Series(df1column11)
df3copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(G.DIV_FULL_NAME))"] = pd.Series(df1column12)
df3copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(E.ACNT_FULL_NAME))"] = pd.Series(df1column13)

#欄位轉換中文
df3copy.rename(columns = {'CLINIC_DATE':'就醫日',
                          'CHART_NO':'病歷號',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(A.PT_NAME))':'病人姓名',
                          'SEQ':'序號',
                          'DISEASE_CODE':'診斷碼',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))':'醫師姓名',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(G.DIV_FULL_NAME))':'科別全名',
                          'CODE':'院內碼',
                          'NH_CODE':'健保碼',                          
                          'MAP_ONH_ACNT_NO':'健保科目',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(E.ACNT_FULL_NAME))':'科目名稱',
                          'SUM(B.NH_AMT_1+B.NH_AMT_2)':'健保金額',
                          'BUILDING_NO':'院區'
                          }, inplace = True)
#樞紐分析 
df3copysum = pd.DataFrame({'健保金額':[df3copy['健保金額'].astype("int").sum()]})

#轉出試算表
df3filename = '3.' + str(year) + str(month) + str(day) + '-YR-BC肝(R)' + '.xlsx'
with pd.ExcelWriter(df3filename) as writer:
    df3copy.to_excel(writer, index=False, sheet_name='Result')
    df3copysum.to_excel(writer, index=False, sheet_name='Pivot')


# In[24]:


#10-4.df4
#資料轉碼
   
df4copy = df4.copy()
column14 = ''
column15 = ''
column16 = ''
column17 = ''
df1column14 = []
df1column15 = []
df1column16 = []
df1column17 = []

for index, data in df4copy.iterrows():    
    column14 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(A.PT_NAME))"]).decode('big5','ignore').rstrip()
    column15 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))"]).decode('big5','ignore').rstrip()
    column16 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(G.DIV_FULL_NAME))"]).decode('big5','ignore').rstrip()
    column17 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(E.ACNT_FULL_NAME))"]).decode('big5','ignore').rstrip()
    df1column14.append(column14)
    df1column15.append(column15)
    df1column16.append(column16)
    df1column17.append(column17)
    
df4copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(A.PT_NAME))"] = pd.Series(df1column14)
df4copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))"] = pd.Series(df1column15)
df4copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(G.DIV_FULL_NAME))"] = pd.Series(df1column16)
df4copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(E.ACNT_FULL_NAME))"] = pd.Series(df1column17)

#欄位轉換中文
df4copy.rename(columns = {'CLINIC_DATE':'就醫日',
                          'CHART_NO':'病歷號',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(A.PT_NAME))':'病人姓名',
                          'SEQ':'序號',
                          'DISEASE_CODE':'診斷碼',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))':'醫師姓名',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(G.DIV_FULL_NAME))':'科別全名',
                          'CODE':'院內碼',
                          'NH_CODE':'健保碼',                          
                          'MAP_ONH_ACNT_NO':'健保科目',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(E.ACNT_FULL_NAME))':'科目名稱',
                          'SUM(B.NH_AMT_1+B.NH_AMT_2)':'健保金額',
                          'BUILDING_NO':'院區'
                          }, inplace = True)
#樞紐分析 
df4copysum = pd.DataFrame({'健保金額':[df4copy['健保金額'].astype("int").sum()]})

#轉出試算表
df4filename = '4.' + str(year) + str(month) + str(day) + '-YR-C肝口服(R)' + '.xlsx'
with pd.ExcelWriter(df4filename) as writer:
    df4copy.to_excel(writer, index=False, sheet_name='Result')
    df4copysum.to_excel(writer, index=False, sheet_name='Pivot') 


# In[25]:


#10-5.df5
#資料轉碼  
df5copy = df5.copy()
column18 = ''
column19 = ''
column20 = ''
column21 = ''
column22 = ''
df1column18 = []
df1column19 = []
df1column20 = []
df1column21 = []
df1column22 = []

for index, data in df5copy.iterrows():    
    column18 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(A.PT_NAME))"]).decode('big5','ignore').rstrip()
    column19 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))"]).decode('big5','ignore').rstrip()
    column20 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(F.DIV_FULL_NAME))"]).decode('big5','ignore').rstrip()
    column21 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(E.ACNT_FULL_NAME))"]).decode('big5','ignore').rstrip()
    column22 = binascii.unhexlify(data["ER_ACNT_CLOSE_FLAG"]).decode('big5','ignore').rstrip()
    df1column18.append(column18)
    df1column19.append(column19)
    df1column20.append(column20)
    df1column21.append(column21)
    df1column22.append(column22)
    
df5copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(A.PT_NAME))"] = pd.Series(df1column18)
df5copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))"] = pd.Series(df1column19)
df5copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(F.DIV_FULL_NAME))"] = pd.Series(df1column20)
df5copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(E.ACNT_FULL_NAME))"] = pd.Series(df1column21)
df5copy["ER_ACNT_CLOSE_FLAG"] = pd.Series(df1column22)

#欄位轉換中文
df5copy.rename(columns = {'CLINIC_DATE':'就醫日',
                          'CHART_NO':'病歷號',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(A.PT_NAME))':'病人姓名',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))':'醫師姓名',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(F.DIV_FULL_NAME))':'科別全名',
                          'CODE':'院內碼',
                          'NH_CODE':'健保碼',
                          'MAP_ONH_ACNT_NO':'健保科目',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(E.ACNT_FULL_NAME))':'科目名稱',
                          '(B.NH_AMT_1+B.NH_AMT_2)':'健保點數',
                          'CLINIC_FLAG':'診別',
                          'ER_ACNT_CLOSE_FLAG':'轉住院註記',
                          'BUILDING_NO':'院區'
                          }, inplace = True)
#樞紐分析 
df5copysum1 = df5copy[["轉住院註記","健保點數"]]
df5copysum2 = df5copysum1.groupby("轉住院註記").sum()
df5copysum3 = df5copysum2.reset_index()
df5copysum4 = pd.DataFrame({"轉住院註記":["總計"],"健保點數":[df5copysum1["健保點數"].sum()]})
df5copysum5 = df5copysum3.append(df5copysum4).reset_index(drop = True)

#轉出試算表
df5filename = '5.YR' + str(year) + str(month) + str(day) + '門診非藥費(急轉住)' + '.xlsx'
with pd.ExcelWriter(df5filename) as writer:
    df5copy.to_excel(writer, index=False, sheet_name='Result')
    df5copysum5.to_excel(writer, index=False, sheet_name='Pivot') 


# In[26]:


#10-6.df6
#資料轉碼  
df6copy = df6.copy()
column23 = ''
column24 = ''
column25 = ''
column26 = ''
column27 = ''
df1column23 = []
df1column24 = []
df1column25 = []
df1column26 = []
df1column27 = []

for index, data in df6copy.iterrows():    
    column23 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(A.PT_NAME))"]).decode('big5','ignore').rstrip()
    column24 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))"]).decode('big5','ignore').rstrip()
    column25 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(G.DIV_FULL_NAME))"]).decode('big5','ignore').rstrip()
    column26 = binascii.unhexlify(data["RAWTOHEX(UTL_RAW.CAST_TO_RAW(E.ACNT_FULL_NAME))"]).decode('big5','ignore').rstrip()
    column27 = binascii.unhexlify(data["ER_ACNT_CLOSE_FLAG"]).decode('big5','ignore').rstrip()
    df1column23.append(column23)
    df1column24.append(column24)
    df1column25.append(column25)
    df1column26.append(column26)
    df1column27.append(column27)
    
df6copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(A.PT_NAME))"] = pd.Series(df1column23)
df6copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))"] = pd.Series(df1column24)
df6copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(G.DIV_FULL_NAME))"] = pd.Series(df1column25)
df6copy["RAWTOHEX(UTL_RAW.CAST_TO_RAW(E.ACNT_FULL_NAME))"] = pd.Series(df1column26)
df6copy["ER_ACNT_CLOSE_FLAG"] = pd.Series(df1column27)

#欄位轉換中文
df6copy.rename(columns = {'CLINIC_DATE':'就醫日',
                          'CHART_NO':'病歷號',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(A.PT_NAME))':'病人姓名',
                          'SEQ':'序號',
                          'DISEASE_CODE':'診斷碼',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(C.DOCTOR_NAME))':'醫師姓名',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(G.DIV_FULL_NAME))':'科別全名',
                          'CODE':'院內碼',
                          'NH_CODE':'健保碼',
                          'MAP_ONH_ACNT_NO':'健保科目',
                          'RAWTOHEX(UTL_RAW.CAST_TO_RAW(E.ACNT_FULL_NAME))':'科目名稱',
                          'SUM(B.NH_AMT_1+B.NH_AMT_2)':'健保金額',
                          'CLINIC_FLAG':'診別',
                          'ER_ACNT_CLOSE_FLAG':'轉住院註記',
                          'BUILDING_NO':'院區'
                          }, inplace = True)
#樞紐分析 
df6copysum1 = df6copy[["轉住院註記","健保金額"]]
df6copysum2 = df6copysum1.groupby("轉住院註記").sum()
df6copysum3 = df6copysum2.reset_index()
df6copysum4 = pd.DataFrame({"轉住院註記":["總計"],"健保金額":[df6copysum1["健保金額"].sum()]})
df6copysum5 = df6copysum3.append(df6copysum4).reset_index(drop = True)

#轉出試算表
df6filename = '6.YR' + str(year) + str(month) + str(day) + '門診藥費(急轉住)' + '.xlsx'
with pd.ExcelWriter(df6filename) as writer:
    df6copy.to_excel(writer, index=False, sheet_name='Result')
    df6copysum5.to_excel(writer, index=False, sheet_name='Pivot')


# In[27]:


#10-7.df7
#資料複製  
df7copy = df7.copy()
#欄位轉換中文
df7copy.rename(columns = {'CHART_NO':'病歷號',
                          'BUILDING_NO':'院區'
                          }, inplace = True)
#樞紐分析 
df7copysum1 = pd.DataFrame(df7copy["病歷號"])
df7copysum2 = df7copysum1.drop_duplicates()
df7copysum3 = df7copysum2.count().reset_index()
df7copysum3.rename(columns = {'index':'病歷號',
                              0:'計數'
                              }, inplace = True)


#轉出試算表
df7filename = '7.YR' + str(year) + quarter + '-' + str(year) + str(month) + str(day) + '門診非藥費人頭數' + '.xlsx'
with pd.ExcelWriter(df7filename) as writer:
    df7copy.to_excel(writer, index=False, sheet_name='Result')
    df7copysum3.to_excel(writer, index=False, sheet_name='Pivot')


# In[28]:


#10-8.df8
#資料複製  
df8copy = df8.copy()
#欄位轉換中文
df8copy.rename(columns = {'CHART_NO':'病歷號',
                          'BUILDING_NO':'院區'
                          }, inplace = True)
#樞紐分析 
df8copysum1 = pd.DataFrame(df8copy["病歷號"])
df8copysum2 = df8copysum1.drop_duplicates()
df8copysum3 = df8copysum2.count().reset_index()
df8copysum3.rename(columns = {'index':'病歷號',
                              0:'計數'
                              }, inplace = True)


#轉出試算表
df8filename = '8.YR' + str(year) + quarter + '-' + str(year) + str(month) + str(day) + '門診藥費人頭數' + '.xlsx'
with pd.ExcelWriter(df8filename) as writer:
    df8copy.to_excel(writer, index=False, sheet_name='Result')
    df8copysum3.to_excel(writer, index=False, sheet_name='Pivot')


# In[29]:


#10-9.df9
#資料複製  
df9copy = df9.copy()
#欄位轉換中文
df9copy.rename(columns = {'NH_APPLY_YYYMM':'申報年月',
                          'ID_NO':'申報的身份証號',
                          '(TOTAL_AMT-AMT_13)':'非藥費',
                          'BED_NO':'院區',
                          }, inplace = True)
#樞紐分析 
df9copysum1 = pd.DataFrame(df9copy[["申報的身份証號","非藥費"]])
df9copysum2 = df9copysum1["申報的身份証號"].count()
df9copysum3 = df9copysum1["非藥費"].sum()
df9copysum4 = pd.DataFrame({"申報的身份証號":[df9copysum2],"非藥費":[df9copysum3]})


#轉出試算表
df9filename = '9.YR' + str(year) + quarter + '-' + str(year) + str(month) + str(day) + '住院非藥費人次數' + '.xlsx'
with pd.ExcelWriter(df9filename) as writer:
    df9copy.to_excel(writer, index=False, sheet_name='Result')
    df9copysum4.to_excel(writer, index=False, sheet_name='Pivot')


# In[30]:


#10-10.df10
#資料複製  
df10copy = df10.copy()
#欄位轉換中文
df10copy.rename(columns = {'NH_APPLY_YYYMM':'申報年月',
                          'ID_NO':'申報的身份証號',
                          'AMT_13':'藥費',
                          'BED_NO':'院區',
                          }, inplace = True)
#樞紐分析 
df10copysum1 = pd.DataFrame(df10copy["申報的身份証號"]).drop_duplicates().count().values[0]
df10copysum2 = df10copy["藥費"].sum()
df10copysum3 = pd.DataFrame({"申報的身份証號":[df10copysum1],"藥費":[df10copysum2]})


#轉出試算表
df10filename = '10.YR' + str(year) + quarter + '-' + str(year) + str(month) + str(day) + '住院藥費人頭' + '.xlsx'
with pd.ExcelWriter(df10filename) as writer:
    df10copy.to_excel(writer, index=False, sheet_name='Result')
    df10copysum3.to_excel(writer, index=False, sheet_name='Pivot')


# In[31]:


#10-11.df11
#資料複製  
df11copy = df11.copy()
#欄位轉換中文
df11copy.rename(columns = {'NH_APPLY_YYYMM':'申報年月',
                          'ADMIT_NO':'住院編號',
                          'CODE':'院內碼',
                          'NH_CODE':'健保醫令代碼',
                          'NH_AMT':'健保金額',
                          'NH_BED_NO':'院區'
                          }, inplace = True)
#樞紐分析
df11copysum1 = df11copy[["申報年月","健保金額"]]    
df11copysum2 = df11copysum1.groupby("申報年月").sum()
df11copysum3 = df11copysum2.reset_index()


#轉出試算表
df11filename = '11.' + str(year) + str(month) + str(day) + 'YR-住院BC肝及C肝口服' + '.xlsx'
with pd.ExcelWriter(df11filename) as writer:
    df11copy.to_excel(writer, index=False, sheet_name='Result')
    df11copysum3.to_excel(writer, index=False, sheet_name='Pivot')


# In[32]:


#10-12.df12
#資料複製  
df12copy = df12.copy()
#欄位轉換中文
df12copy.rename(columns = {'NH_APPLY_YYYMM':'申報年月',
                          'ID_NO':'申報的身份証號',
                          'TOTAL_AMT':'申報金額',
                          'BED_NO':'院區',
                          }, inplace = True)
#樞紐分析 
df12copysum1 = pd.DataFrame(df12copy[["申報金額"]])
df12copysum2 = df12copysum1["申報金額"].sum()
df12copysum3 = pd.DataFrame({"申報金額":[df12copysum2]})


#轉出試算表
df12filename = '12.YR' + str(year) + quarter + str(year) + str(month) + str(day) + '-院代辦案件預估(含C5)' + '.xlsx'
with pd.ExcelWriter(df12filename) as writer:
    df12copy.to_excel(writer, index=False, sheet_name='Result')
    df12copysum3.to_excel(writer, index=False, sheet_name='Pivot') 


# In[33]:


#11.關閉資料庫連線
connection.close()

