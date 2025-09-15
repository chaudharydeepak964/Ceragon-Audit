print("welcome to Audit Report ")

import pandas as pd
import numpy as np
import paramiko 
import datetime
from datetime import datetime, timedelta
import win32com.client as wincl
from openpyxl import load_workbook
import os
import glob
import re

#############

# Auto date

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

do=(datetime.now() - timedelta(1)).strftime('%Y%m%d')
do1=(datetime.now() - timedelta(2)).strftime('%Y%m%d')
do2=(datetime.now() - timedelta(3)).strftime('%Y%m%d')


dm=(datetime.now() - timedelta(1)).strftime('%m-%y')
dm1=(datetime.now() - timedelta(1)).strftime('%m%Y')
dn=(datetime.now() - timedelta(1)).strftime('%d-%m-%Y')
dnn=(datetime.now() - timedelta(1)).strftime('%d.%m.%Y')

print(d57)

dmm=(datetime.now() - timedelta(1)).strftime('%m-%y')
print(dmm)


print("1st Day Reading Start")

####working start ########

#asm= pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ASM_I_Full_Link_Report_'+da+'.csv',skiprows=5)
#bih = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/BIH_I_Full_Link_Report_'+da+'.csv',skiprows=5)
rob = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ROB_Full_Link_Report_'+da+'.csv',skiprows=5,encoding= 'unicode_escape')
kar= pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/KAR_Full_Link_Report_'+da+'.csv',skiprows=5)
#odi = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ODI_Full_Link_Report_'+do+'.csv',skiprows=5)
upe = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/UPE_Full_Link_Report_'+da+'.csv',skiprows=5,encoding= 'unicode_escape')



print("2nd Day Reading Start")


#asm1= pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ASM_I_Full_Link_Report_'+da1+'.csv',skiprows=5)
#bih1 = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/BIH_I_Full_Link_Report_'+da1+'.csv',skiprows=5)
rob1 = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ROB_Full_Link_Report_'+da1+'.csv',skiprows=5,encoding= 'unicode_escape')
kar1= pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/KAR_Full_Link_Report_'+da1+'.csv',skiprows=5)
#odi1 = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ODI_Full_Link_Report_'+do1+'.csv',skiprows=5)
upe1 = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/UPE_Full_Link_Report_'+da1+'.csv',skiprows=5,encoding= 'unicode_escape')



print("3rd Day Reading Start")


#asm2= pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ASM_I_Full_Link_Report_'+da2+'.csv',skiprows=5)
#bih2 = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/BIH_I_Full_Link_Report_'+da2+'.csv',skiprows=5)
rob2 = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ROB_Full_Link_Report_'+da2+'.csv',skiprows=5,encoding= 'unicode_escape')
kar2= pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/KAR_Full_Link_Report_'+da2+'.csv',skiprows=5)
#odi2 = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/ODI_Full_Link_Report_'+do2+'.csv',skiprows=5)
upe2 = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\RAW\CERAGON/UPE_Full_Link_Report_'+da2+'.csv',skiprows=5,encoding= 'unicode_escape')


print("Reading Done")

# Blank Column Create

#odi['ATPC']=''
#odi1['ATPC']=''
#odi2['ATPC']=''



# Selected Column Pickup from 1st Day Files


#asm=asm[['Site A Name','Site A Physical Port','Site B Name','Site B Physical Port','Link Configuration','Site A IP','Site B IP',
#         'Site A Radio','Site B Radio','Site A Tx Freq [MHz]','Site B Tx Freq [MHz]','Site A Radio Script','Site B Radio Script',
#         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]


#bih=bih[['Site A Name','Site A Physical Port','Site B Name','Site B Physical Port','Link Configuration','Site A IP','Site B IP',
#         'Site A Radio','Site B Radio','Site A Tx Freq [MHz]','Site B Tx Freq [MHz]','Site A Radio Script','Site B Radio Script',
#         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]


#odi=odi[['Site A Name','Site A Physical Port','Site B Name','Site B Physical Port','Link Configuration','Site A Id','Site B Id',
#         'Site A Radio','Site B Radio','Site A Tx Freq [MHz]','Site B Tx Freq [MHz]','Site A Radio Script','Site B Radio Script','ATPC']]


upe=upe[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP',
         'Site A Radio','Site Z Radio','Site A Tx Freq [MHz]','Site Z Tx Freq [MHz]','Site A Radio Script','Site Z Radio Script',
         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]


kar=kar[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP',
         'Site A Radio','Site Z Radio','Site A Tx Freq [MHz]','Site Z Tx Freq [MHz]','Site A Radio Script','Site Z Radio Script',
         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]


rob=rob[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP',
         'Site A Radio','Site Z Radio','Site A Tx Freq [MHz]','Site Z Tx Freq [MHz]','Site A Radio Script','Site Z Radio Script',
         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]



# Selected Column Pickup from 2nd Day Files


#asm1=asm1[['Site A Name','Site A Physical Port','Site B Name','Site B Physical Port','Link Configuration','Site A IP','Site B IP',
#         'Site A Radio','Site B Radio','Site A Tx Freq [MHz]','Site B Tx Freq [MHz]','Site A Radio Script','Site B Radio Script',
#         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]


#bih1=bih1[['Site A Name','Site A Physical Port','Site B Name','Site B Physical Port','Link Configuration','Site A IP','Site B IP',
#         'Site A Radio','Site B Radio','Site A Tx Freq [MHz]','Site B Tx Freq [MHz]','Site A Radio Script','Site B Radio Script',
#         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]


#odi1=odi1[['Site A Name','Site A Physical Port','Site B Name','Site B Physical Port','Link Configuration','Site A Id','Site B Id',
#         'Site A Radio','Site B Radio','Site A Tx Freq [MHz]','Site B Tx Freq [MHz]','Site A Radio Script','Site B Radio Script','ATPC']]


upe1=upe1[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP',
         'Site A Radio','Site Z Radio','Site A Tx Freq [MHz]','Site Z Tx Freq [MHz]','Site A Radio Script','Site Z Radio Script',
         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]


kar1=kar1[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP',
         'Site A Radio','Site Z Radio','Site A Tx Freq [MHz]','Site Z Tx Freq [MHz]','Site A Radio Script','Site Z Radio Script',
         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]


rob1=rob1[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP',
         'Site A Radio','Site Z Radio','Site A Tx Freq [MHz]','Site Z Tx Freq [MHz]','Site A Radio Script','Site Z Radio Script',
         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]



# Selected Column Pickup from 2nd Day Files


#asm2=asm2[['Site A Name','Site A Physical Port','Site B Name','Site B Physical Port','Link Configuration','Site A IP','Site B IP',
#         'Site A Radio','Site B Radio','Site A Tx Freq [MHz]','Site B Tx Freq [MHz]','Site A Radio Script','Site B Radio Script',
#         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]


#bih2=bih2[['Site A Name','Site A Physical Port','Site B Name','Site B Physical Port','Link Configuration','Site A IP','Site B IP',
#         'Site A Radio','Site B Radio','Site A Tx Freq [MHz]','Site B Tx Freq [MHz]','Site A Radio Script','Site B Radio Script',
#         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]


#odi2=odi2[['Site A Name','Site A Physical Port','Site B Name','Site B Physical Port','Link Configuration','Site A Id','Site B Id',
#         'Site A Radio','Site B Radio','Site A Tx Freq [MHz]','Site B Tx Freq [MHz]','Site A Radio Script','Site B Radio Script','ATPC']]


upe2=upe2[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP',
         'Site A Radio','Site Z Radio','Site A Tx Freq [MHz]','Site Z Tx Freq [MHz]','Site A Radio Script','Site Z Radio Script',
         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]


kar2=kar2[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP',
         'Site A Radio','Site Z Radio','Site A Tx Freq [MHz]','Site Z Tx Freq [MHz]','Site A Radio Script','Site Z Radio Script',
         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]


rob2=rob2[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP',
         'Site A Radio','Site Z Radio','Site A Tx Freq [MHz]','Site Z Tx Freq [MHz]','Site A Radio Script','Site Z Radio Script',
         'MRMC Script Operational Mode','MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile','ATPC']]




##Rename Column name 1st Day File 

upe.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site Z Radio':'Site B Radio','Site Z Tx Freq [MHz]':'Site B Tx Freq [MHz]','Site Z Radio Script':'Site B Radio Script'},inplace=True)
kar.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site Z Radio':'Site B Radio','Site Z Tx Freq [MHz]':'Site B Tx Freq [MHz]','Site Z Radio Script':'Site B Radio Script'},inplace=True)
rob.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site Z Radio':'Site B Radio','Site Z Tx Freq [MHz]':'Site B Tx Freq [MHz]','Site Z Radio Script':'Site B Radio Script'},inplace=True)
#odi.rename(columns={'Site A Id':'Site A IP','Site B Id':'Site B IP'},inplace=True)



##Rename Column name 2nd Day File 

upe1.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site Z Radio':'Site B Radio','Site Z Tx Freq [MHz]':'Site B Tx Freq [MHz]','Site Z Radio Script':'Site B Radio Script'},inplace=True)
kar1.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site Z Radio':'Site B Radio','Site Z Tx Freq [MHz]':'Site B Tx Freq [MHz]','Site Z Radio Script':'Site B Radio Script'},inplace=True)
rob1.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site Z Radio':'Site B Radio','Site Z Tx Freq [MHz]':'Site B Tx Freq [MHz]','Site Z Radio Script':'Site B Radio Script'},inplace=True)
#odi1.rename(columns={'Site A Id':'Site A IP','Site B Id':'Site B IP'},inplace=True)



##Rename Column name 3rd Day File

upe2.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site Z Radio':'Site B Radio','Site Z Tx Freq [MHz]':'Site B Tx Freq [MHz]','Site Z Radio Script':'Site B Radio Script'},inplace=True)
kar2.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site Z Radio':'Site B Radio','Site Z Tx Freq [MHz]':'Site B Tx Freq [MHz]','Site Z Radio Script':'Site B Radio Script'},inplace=True)
rob2.rename(columns={'Site Z Name':'Site B Name','Site Z Physical Port':'Site B Physical Port','Site Z IP':'Site B IP','Site Z Radio':'Site B Radio','Site Z Tx Freq [MHz]':'Site B Tx Freq [MHz]','Site Z Radio Script':'Site B Radio Script'},inplace=True)
#odi2.rename(columns={'Site A Id':'Site A IP','Site B Id':'Site B IP'},inplace=True)




## Manually Circle Add all 3 days Files

kar['Circle']='KAR'
rob['Circle']='ROB'

kar1['Circle']='KAR'
rob1['Circle']='ROB'

kar2['Circle']='KAR'
rob2['Circle']='ROB'


## Manually Server Add all 3 days Files

#asm['server']='ASM'
#bih['server']='BIH'
#odi['server']='ODI'
upe['server']='UPE'
kar['server']='KAR'
rob['server']='ROB'


#asm1['server']='ASM'
#bih1['server']='BIH'
#odi1['server']='ODI'
upe1['server']='UPE'
kar1['server']='KAR'
rob1['server']='ROB'


#asm2['server']='ASM'
#bih2['server']='BIH'
#odi2['server']='ODI'
upe2['server']='UPE'
kar2['server']='KAR'
rob2['server']='ROB'


#df=pd.concat([asm,bih,odi,upe,rob,kar,asm1,bih1,odi1,upe1,rob1,kar1,asm2,bih2,odi2,upe2,rob2,kar2])

df=pd.concat([upe,rob,kar,upe1,rob1,kar1,upe2,rob2,kar2])


df['LT']='1+0'

#df.loc[df['Link Configuration'].isin(['Xpic', '2+0']), 'LT'] = 'Xpic'

# **** Update Link type ****

df['LT'] = np.where(df['Link Configuration'].str.contains(r'Xpic|2\+0'), 'Xpic', '1+0')


# **** ACM STATUS UPDATE ****

#df['ACM Status'] = df['MRMC Script Operational Mode'].apply(lambda x: 'Enabled' if 'Adaptive' in x else 'Disabled')

df['ACM Status'] = df['MRMC Script Operational Mode'].str.contains(r'\bAdaptive\b').map({True: 'Enabled', False: 'Disabled'})



# **** Finding Modulation and Channel Spaccing as per Radio Script ****


df['Site A Radio Script'] = df['Site A Radio Script'].fillna(df['Site B Radio Script'])

df['Site A Radio Script'] = df['Site A Radio Script'].astype(str)
df['QAM'] = df['Site A Radio Script'].apply(lambda x: re.search(r'(\d+QAM)', x).group(1) if re.search(r'(\d+QAM)', x) else None)

#df['CHANNEL SPACING'] = df['Site A Radio Script'].apply(lambda x: re.search(r'(\d+MHz)', x).group(1) if re.search(r'(\d+\.\d+MHz)', x) else None)



#(ACM_194Mbps-28MHz-256QAM-Grade-1) Extract 28,  10Mbps-6.5MHz-4QAM-Grade-1 Extract 6.5   **MHz Missing

#df['CHANNEL SPACING'] = df['Site A Radio Script'].apply(lambda x: re.search(r'(\d+(\.\d+)?)MHz', x).group(1) if re.search(r'(\d+(\.\d+)?)MHz', x) else None)



#(ACM_194Mbps-28MHz-256QAM-Grade-1) Extract 28MHz,  10Mbps-6.5MHz-4QAM-Grade-1 Extract 6.5MHz    **MHz Added

df['CHANNEL SPACING'] = df['Site A Radio Script'].apply(lambda x: re.search(r'(\d+(\.\d+)?)MHz', x).group(0) if re.search(r'(\d+(\.\d+)?)MHz', x) else None)


pattern_mapping = {
    '1414': '14MHz',
    '2828': '28MHz',
    '4040': '40MHz',
    '5656': '56MHz',
    '250250': '25MHz',
    '028028': '28MHz',
    '056056': '56MHz'
}

def map_channel_spacing(script):
    for pattern, spacing in pattern_mapping.items():
        if pattern in script:
            return spacing
    return None  # If no pattern matched

# Update missing values in 'CHANNEL SPACING' column
df['CHANNEL SPACING'] = df.apply(
    lambda row: map_channel_spacing(row['Site A Radio Script']) if pd.isna(row['CHANNEL SPACING']) else row['CHANNEL SPACING'],
    axis=1
)



# **** Finding Modulation as per MRMC Script Profile ****


def convert_to_numeric(value):
    try:
        return pd.to_numeric(value)
    except ValueError:
        return value

df['MRMC Script Profile'] = df['MRMC Script Profile'].apply(convert_to_numeric)

df['MRMC Script Maximum Profile'] = df['MRMC Script Maximum Profile'].apply(convert_to_numeric)

df['MRMC Script Minimum Profile'] = df['MRMC Script Minimum Profile'].apply(convert_to_numeric)



# Exclude 50CX and 50E
#dff = df.loc[df['Site A Radio']!='RFU-50CX']
#dff = df.loc[~df['Site A Radio'].str.contains('RFU-50', na=False)]
dff = df.loc[~df['Site A Radio'].isin(['RFU-50CX', 'RFU-50E'])]

# FOR IP-10 and 20
MRMC_order = {0:'QPSK',1:'8QAM',2:'16QAM',3:'32QAM',4:'64QAM',5:'128QAM',6:'256QAM',
		7:'512QAM',8:'1024QAMLight',9:'1024QAM',10:'2048QAM',11:'2048QAM',12:'4096QAM'}

dff['Modulation1']=dff['MRMC Script Profile'].map(MRMC_order)
dff['Modulation2']=dff['MRMC Script Maximum Profile'].map(MRMC_order)
dff['Modulation3']=dff['MRMC Script Minimum Profile'].map(MRMC_order)

try:
    dff = dff.reset_index(drop=True)
except:
    pass

## Find Final modulation on the basic of three columns

dff['Mod'] = dff['Modulation1'].fillna(dff['Modulation2'].combine_first(dff['Modulation3']))

dff['Modulation'] = dff['QAM'].fillna(dff['Mod'])

try:
    dff['Modulation'].fillna('QPSK',inplace=True)
except:
    pass

## --------------------------------------------------------------------------------------------------------

# Only 50CX
#df50=df.loc[df['Site A Radio']=='RFU-50CX']
df50=df.loc[df['Site A Radio'].isin(['RFU-50CX'])]


# FOR IP-50
MRMC_order = {0:'2QAM',1:'4QAM',2:'8QAM',3:'16QAM',4:'32QAM',5:'64QAM',6:'128QAM',7:'256QAM',
		8:'512QAM',9:'1024QAM',10:'1024QAM',11:'2048QAM',12:'4096QAM'}


df50['Modulation1']=df50['MRMC Script Profile'].map(MRMC_order)
df50['Modulation2']=df50['MRMC Script Maximum Profile'].map(MRMC_order)
df50['Modulation3']=df50['MRMC Script Minimum Profile'].map(MRMC_order)

df50 = df50.reset_index(drop=True)

## Find Final modulation on the basic of three columns

df50['Mod'] = df50['Modulation1'].fillna(df50['Modulation2'].combine_first(df50['Modulation3']))

df50['Modulation'] = df50['QAM'].fillna(df50['Mod'])

df50['Modulation'].fillna('QPSK',inplace=True)



## --------------------------------------------------------------------------------------------------------

# Only 50CX
#dff50=df.loc[df['Site A Radio']=='RFU-50CX']
dff50=df.loc[df['Site A Radio'].isin(['RFU-50E'])]


# FOR IP-50
MRMC_order = {0:'BPSK9',1:'BPSK10',2:'BPSK',3:'QPSK',4:'8QAM',5:'16QAM',6:'32QAM',7:'64QAM',
		8:'128QAM',9:'256QAM',10:'512QAM'}


dff50['Modulation1']=dff50['MRMC Script Profile'].map(MRMC_order)
dff50['Modulation2']=dff50['MRMC Script Maximum Profile'].map(MRMC_order)
dff50['Modulation3']=dff50['MRMC Script Minimum Profile'].map(MRMC_order)

dff50 = dff50.reset_index(drop=True)

## Find Final modulation on the basic of three columns

dff50['Mod'] = dff50['Modulation1'].fillna(dff50['Modulation2'].combine_first(dff50['Modulation3']))

dff50['Modulation'] = dff50['QAM'].fillna(dff50['Mod'])

dff50['Modulation'].fillna('QPSK',inplace=True)



## --------------------------------------------------------------------------------------------------------

Comb = pd.concat([dff,df50,dff50])

df = Comb.copy()

#***************************#


df['Source Running status']=''
df['Sink Running Status']=''
df['Source Highest ACM Profile']=''
df['Sink Highest ACM Profile']=''


df['Site A Name']=df['Site A Name'].str.strip()
df['Site B Name']=df['Site B Name'].str.strip()

# unique link*******************************
df['uniq link']=np.where((df['Site A Name']<df['Site B Name']),(df['Site A Name']+'-'+df['Site B Name']),(df['Site B Name']+'-'+ df['Site A Name']))
df['uniq link']=df['uniq link'].str.strip()



# unique link*******************************
df['uniq IP']=np.where((df['Site A IP']<df['Site B IP']),(df['Site A IP']+'-'+df['Site B IP']),(df['Site B IP']+'-'+ df['Site A IP']))
df['uniq IP']=df['uniq IP'].str.strip()

    

# circle ******************


df.loc[df['uniq link'].str.contains('IDDL|INDL',na=False),'Circle']='DEL'
df.loc[df['uniq link'].str.contains('IDUW|INUW',na=False),'Circle']='UPW'
df.loc[df['uniq link'].str.contains('IDOD|INOD',na=False),'Circle']='ODI'
df.loc[df['uniq link'].str.contains('IDKL|INKL',na=False),'Circle']='KEL'
df.loc[df['uniq link'].str.contains('IDAS|IDNE|INAS|INNE', na=False), 'Circle'] = 'ASM'
df.loc[df['uniq link'].str.contains('IDUE|INUE|AZMG',na=False),'Circle']='UPE'
df.loc[df['uniq link'].str.contains('IDKA|INKA|MYS0|MYS9',na=False),'Circle']='KAR'
df.loc[df['uniq link'].str.contains('IDWB|IINW|INEW|INWB',na=False),'Circle']='ROB'
df.loc[df['uniq link'].str.contains(' INB|BBSN|BCHN|BDAR|BLXM|BMRU|bnir|BPIA|BR10|BSAS|BTOD|BUGR|IDB0|IDBR|INBR|JBKU|KOLA|BN2083|BPPK',na=False),'Circle']='BIH'
df.loc[df['uniq link'].str.contains('ARJ0|Bhaw|CHK0|CNR0|DMP0|HWY1|IDJK|INJK|JMU0|jmu1|JMU2|NAG0|RAJ0|SRN0|SRR1|VIJ0',na=False),'Circle']='JNK'

#dff=df.copy()

df['Circle'] = df['Circle'].replace(['ASM'], 'ASM_I')
df['Circle'] = df['Circle'].replace(['BIH'], 'BIH_I')
df['Circle'] = df['Circle'].replace(['DEL'], 'DEL_I')

df = df.drop_duplicates(subset=['uniq link'])

# FOR MRMC SCRIPT PROFILE

MRMC = df.reindex(columns=['server','Circle','LT','uniq link','Site A Name','Site A IP','Site B Name','Site B IP',
                        'MRMC Script Profile','MRMC Script Maximum Profile','MRMC Script Minimum Profile'])



df = df.reindex(columns=['server','Circle','LT','uniq link','uniq IP','Site A Name','Site A IP','Site A Physical Port',
                        'Site B Name','Site B IP','Site B Physical Port','Link Configuration','Site A Radio',
                        'Site B Radio','Site A Tx Freq [MHz]','Site B Tx Freq [MHz]','ATPC','CHANNEL SPACING',
                        'Modulation','ACM Status'])



print("Export")


                    #### ****************************************** PM FILE WORKING *********************************** #####


#AA= pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\A_Radio_PM_Report_IP-10_24h_'+da+'.csv',usecols=['IP','Highest ACM Profile'])
#A= pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\A_Radio_PM_Report_IP-20_24h_'+da+'.csv',usecols=['IP','Highest ACM Profile'])
#BB = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\B_Radio_PM_Report_IP-10_24h_'+da+'.csv',usecols=['IP','Highest ACM Profile'])
#B = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\B_Radio_PM_Report_IP-20_24h_'+da+'.csv',usecols=['IP','Highest ACM Profile'])
R = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\R_Radio_PM_Report_IP-20_24h_'+da+'.csv',encoding= 'unicode_escape',usecols=['System Name','IP','Highest ACM Profile'])
K = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\K_Radio_PM_Report_IP-20_24h_'+da+'.csv',usecols=['System Name','IP','Highest ACM Profile'])
U = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\PM\U_Radio_PM_Report_IP-20_24h_'+da+'.csv',encoding= 'unicode_escape',usecols=['System Name','IP','Highest ACM Profile'])

#pm=pd.concat([AA,A,BB,B,K,R,U])

R['server']='ROB'
K['server']='KAR'
U['server']='UPE'

pm=pd.concat([K,R,U])

pm.loc[pm['System Name'].str.contains('IDDL|INDL',na=False),'Circle']='DEL'
pm.loc[pm['System Name'].str.contains('IDUW|INUW',na=False),'Circle']='UPW'
pm.loc[pm['System Name'].str.contains('IDOD|INOD',na=False),'Circle']='ODI'
pm.loc[pm['System Name'].str.contains('IDKL|INKL',na=False),'Circle']='KEL'
pm.loc[pm['System Name'].str.contains('IDAS|IDNE|INAS|INNE', na=False), 'Circle'] = 'ASM'
pm.loc[pm['System Name'].str.contains('IDUE|INUE|AZMG',na=False),'Circle']='UPE'
pm.loc[pm['System Name'].str.contains('IDKA|INKA|MYS0|MYS9',na=False),'Circle']='KAR'
pm.loc[pm['System Name'].str.contains('IDWB|IINW|INEW|INWB|INKO|IDKO',na=False),'Circle']='ROB'
pm.loc[pm['System Name'].str.contains(' INB|BBSN|BCHN|BDAR|BLXM|BMRU|bnir|BPIA|BR10|BSAS|BTOD|BUGR|IDB0|IDBR|INBR|JBKU|KOLA|BN2083|BPPK',na=False),'Circle']='BIH'
pm.loc[pm['System Name'].str.contains('ARJ0|Bhaw|CHK0|CNR0|DMP0|HWY1|IDJK|INJK|JMU0|jmu1|JMU2|NAG0|RAJ0|SRN0|SRR1|VIJ0',na=False),'Circle']='JNK'


                                        ## **** Find Modulation as per Highest ACM Profile **** ##




#pm['Highest ACM Profile'] = df['Highest ACM Profile'].apply(convert_to_numeric)  

pm['Running_Modulation']=pm['Highest ACM Profile'].map(MRMC_order)  ## MRMC_order dictionary prepared for mapping modulation at line no 202

pm = pm.drop_duplicates(subset=['IP'])

#Create Duplicate Columns

pm['Site A IP'] = pm['IP']
pm['Site B IP'] = pm['IP']

pm['Source Running status'] = pm['Running_Modulation']
pm['Sink Running Status'] = pm['Running_Modulation']

## Merging

So_pm=pm[['Site A IP', 'Source Running status']]
df=pd.merge(df,So_pm,on='Site A IP',how='left')


Si_pm=pm[['Site B IP', 'Sink Running Status']]
df=pd.merge(df,Si_pm,on='Site B IP',how='left')

print("writing")

writer = pd.ExcelWriter(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\Output\TXN_Cera_PAN-INDIA_Ceragon_MW_Audit_Report_'+dm1+'_MONTHLY.xlsx'
                        ,engine='openpyxl')

df.to_excel(writer, sheet_name= 'Details',index=False)
MRMC.to_excel(writer, sheet_name= 'MRMC Script',index=False)
pm.to_excel(writer, sheet_name= 'PM_Data',index=False)

writer.close()

print("done")


# Upload code COBRA --- 

print("Login Cobra")

ssh3=paramiko.SSHClient()
ssh3.set_missing_host_key_policy(paramiko.AutoAddPolicy())
try:
    ssh3.connect(hostname='10.115.1.57',username='Cobra',password='Cobra@123',port=22)
except:
    pass
try:
    ssh3.connect(hostname='10.19.62.229',username='Cobra',password='Cobra@123',port=22)
except:
    pass
sftp_client1=ssh3.open_sftp()

sftp_client1.chdir('/opt/MyLog/TX/Modulation_report')

sftp_client1.put(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Audit Report\Output\TXN_Cera_PAN-INDIA_Ceragon_MW_Audit_Report_'+dm1+'_MONTHLY.xlsx', 'TXN_Cera_PAN-INDIA_Ceragon_MW_Audit_Report_'+dm1+'_MONTHLY.xlsx')


sftp_client1.close
ssh3.close

print("Upload done")
