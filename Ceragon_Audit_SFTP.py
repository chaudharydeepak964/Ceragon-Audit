print("welcome to Ceragon Audit Report SFTP ")

import pandas as pd
import numpy as np
import paramiko 
import datetime
from datetime import datetime, timedelta
import win32com.client as wincl
from openpyxl import load_workbook
import os
import glob


d57= datetime.now() - timedelta(1)
d56= datetime.now() - timedelta(2)
d55 = datetime.now() - timedelta(3)
d54 = datetime.now() - timedelta(4)
d53 = datetime.now() - timedelta(5)
d52 = datetime.now() - timedelta(6)
d51 = datetime.now() - timedelta(7)

d57 = datetime.strftime(d57, '%d_%m_%Y')
d56 = datetime.strftime(d56, '%d_%m_%Y')
d55 = datetime.strftime(d55, '%d_%m_%Y')
d54 = datetime.strftime(d54, '%d_%m_%Y')
d53 = datetime.strftime(d53, '%d_%m_%Y')
d52 = datetime.strftime(d52, '%d_%m_%Y')
d51 = datetime.strftime(d51, '%d_%m_%Y')


da=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')
da1=(datetime.now() - timedelta(2)).strftime('%d_%m_%Y')
da2=(datetime.now() - timedelta(3)).strftime('%d_%m_%Y')
da3=(datetime.now() - timedelta(4)).strftime('%d_%m_%Y')

do=(datetime.now() - timedelta(1)).strftime('%Y%m%d')
do1=(datetime.now() - timedelta(2)).strftime('%Y%m%d')
do2=(datetime.now() - timedelta(3)).strftime('%Y%m%d')


dm=(datetime.now() - timedelta(1)).strftime('%m-%y')
dn=(datetime.now() - timedelta(1)).strftime('%d-%m-%Y')
dnn=(datetime.now() - timedelta(1)).strftime('%d.%m.%Y')

print(d57)

dmm=(datetime.now() - timedelta(1)).strftime('%m-%y')
print(dmm)

print('Raw files Removing....')



directory=r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON'
os.chdir(directory)
files=glob.glob('*.csv')
for filename in files:
    os.unlink(filename)

directory=r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM'
os.chdir(directory)
files=glob.glob('*.csv')
for filename in files:
    os.unlink(filename)


    
print("Removed....")

ssh3=paramiko.SSHClient()
ssh3.set_missing_host_key_policy(paramiko.AutoAddPolicy())
try:
    ssh3.connect(hostname='10.10.10.10',username='admin',password='admin',port=22)
except:
    pass

try:
    ssh3.connect(hostname='11.11.11.11',username='admin',password='admin',port=22)
except:
    pass

sftp_client1=ssh3.open_sftp()


print("** Downloading 1st Day CM Files ** ")

sftp_client1.chdir('/opt/mycom/cm1/ASM_I/')
try:
    sftp_client1.get('Report_'+da+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ASM_I_Report_'+da+'.csv')
except:
    pass
print("1.Full link ASM_I.csv downloaded")

sftp_client1.chdir('/opt/mycom/cm1/BIH_I')
try:
    sftp_client1.get('Report_'+da+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/BIH_I_Report_'+da+'.csv')
except:
    pass
print("2.Full link BIH_I.csv downloaded")

sftp_client1.chdir('/opt/mycom/cm1/ODI/')
try:
    sftp_client1.get('Report_'+do+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ODI_Report_'+do+'.csv')
except:
    pass
print("3.Full link ODI.csv downloaded")


sftp_client1.chdir('/opt/mycom/cm1/UPE/')
try:
    sftp_client1.get('Report_'+da+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/UPE_Report_'+da+'.csv')
except:
    pass
print("5.Full link UPE.csv downloaded")

sftp_client1.chdir('/opt/mycom/cm1/ROB/')
try:
    sftp_client1.get('Report_'+da+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ROB_Report_'+da+'.csv')
except:
    pass
print("6.Full link ROB.csv downloaded")

sftp_client1.chdir('/opt/mycom/cm1/KAR/')
try:
    sftp_client1.get('Report_'+da+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/KAR_Report_'+da+'.csv')
except:
    pass
print("6.Full link KAR.csv downloaded")

                                      ##### ************************ 2nd Day Downloading CM Files ****************************** ####




print("** Downloading 2nd Day CM Files ** ")

sftp_client1.chdir('/opt/mycom/cm1/ASM_I/')
try:
    sftp_client1.get('Report_'+da1+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ASM_I_Report_'+da1+'.csv')
except:
    pass
print("1.Full link ASM_I.csv downloaded")

sftp_client1.chdir('/opt/mycom/cm1/BIH_I')
try:
    sftp_client1.get('Report_'+da1+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/BIH_I_Report_'+da1+'.csv')
except:
    pass
print("2.Full link BIH_I.csv downloaded")

sftp_client1.chdir('/opt/mycom/cm1/ODI/')
try:
    sftp_client1.get('Report_'+do2+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ODI_Report_'+do2+'.csv')
except:
    pass
print("3.Full link ODI.csv downloaded")

sftp_client1.chdir('/opt/mycom/data/ceragon/microwave/csvascii_nr21/cm1/UPE/')
try:
    sftp_client1.get('Report_'+da1+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/UPE_Report_'+da1+'.csv')
except:
    pass
print("5.Full link UPE.csv downloaded")

sftp_client1.chdir('/opt/mycom/data/ceragon/microwave/csvascii_nr21/cm1/ROB/')
try:
    sftp_client1.get('Report_'+da1+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ROB_Report_'+da1+'.csv')
except:
    pass
print("6.Full link ROB.csv downloaded")

sftp_client1.chdir('/opt/mycom/data/ceragon/microwave/csvascii_nr21/cm1/KAR/')
try:
    sftp_client1.get('Report_'+da1+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/KAR_Report_'+da1+'.csv')
except:
    pass
print("6.Full link KAR.csv downloaded")




                              ##### ************************ 3rd Day Downloading CM Files ****************************** ####




print("** Downloading 3rd Day CM Files ** ")

sftp_client1.chdir('/opt/mycom/cm1/ASM_I/')
try:
    sftp_client1.get('Report_'+da2+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ASM_I_Report_'+da2+'.csv')
except:
    pass
print("1.Full link ASM_I.csv downloaded")

sftp_client1.chdir('/opt/mycom/cm1/BIH_I')
try:
    sftp_client1.get('Report_'+da2+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/BIH_I_Report_'+da2+'.csv')
except:
    pass
print("2.Full link BIH_I.csv downloaded")

sftp_client1.chdir('/opt/mycom/cm1/ODI/')
try:
    sftp_client1.get('Report_'+do1+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ODI_Report_'+do1+'.csv')
except:
    pass
print("3.Full link ODI.csv downloaded")

sftp_client1.chdir('/opt/mycom/data/ceragon/microwave/csvascii_nr21/cm1/UPE/')
try:
    sftp_client1.get('Report_'+da2+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/UPE_Report_'+da2+'.csv')
except:
    pass
print("5.Full link UPE.csv downloaded")

sftp_client1.chdir('/opt/mycom/data/ceragon/microwave/csvascii_nr21/cm1/ROB/')
try:
    sftp_client1.get('Report_'+da2+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ROB_Report_'+da2+'.csv')
except:
    pass
print("6.Full link ROB.csv downloaded")

sftp_client1.chdir('/opt/mycom/data/ceragon/microwave/csvascii_nr21/cm1/KAR/')
try:
    sftp_client1.get('Report_'+da2+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/KAR_Report_'+da2+'.csv')
except:
    pass
print("6.Full link KAR.csv downloaded")


print("** CM FILES DOWNLOADED ** ")



                                        ##### ************************ PM Downloading ****************************** #####


print("** Downloading Latest PM Files ** ")


sftp_client1.chdir('/opt/mycom/pm1/ASM_I/')
try:
    sftp_client1.get('Radio_PM_Report_IP-20_24h_'+da+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\A_Radio_PM_Report_IP-20_24h_'+da+'.csv')
except:
    pass
try:
    sftp_client1.get('Radio_PM_Report_IP-10_24h_'+da+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\A_Radio_PM_Report_IP-10_24h_'+da+'.csv')
except:
    pass
print("1.ASM_I PM downloaded")

sftp_client1.chdir('/opt/mycom/pm1/BIH_I')
try:
    sftp_client1.get('Radio_PM_Report_IP-20_24h_'+da+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\B_Radio_PM_Report_IP-20_24h_'+da+'.csv')
except:
    pass
try:
    sftp_client1.get('Radio_PM_Report_IP-10_24h_'+da+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\B_Radio_PM_Report_IP-10_24h_'+da+'.csv')
except:
    pass
print("2.BIH_I PM downloaded")

sftp_client1.chdir('/opt/mycom/data/ceragon/microwave/csvascii_nr21/pm1/UPE/')
try:
    sftp_client1.get('Radio_PM_Report_IP-20_24h_'+da+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\U_Radio_PM_Report_IP-20_24h_'+da+'.csv')
except:
    pass
print("5.UPE PM downloaded")

sftp_client1.chdir('/opt/mycom/data/ceragon/microwave/csvascii_nr21/pm1/ROB/')
try:
    sftp_client1.get('Radio_PM_Report_IP-20_24h_'+da+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\R_Radio_PM_Report_IP-20_24h_'+da+'.csv')
except:
    pass
print("6.ROB PM downloaded")

sftp_client1.chdir('/opt/mycom/data/ceragon/microwave/csvascii_nr21/pm1/KAR/')
try:
    sftp_client1.get('Radio_PM_Report_IP-20_24h_'+da+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\K_Radio_PM_Report_IP-20_24h_'+da+'.csv')
except:
    pass
print("6.KAR PM downloaded")

print("Done")








