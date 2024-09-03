import pandas as pd 
import numpy as np
import shutil
from tkinter import *
from openpyxl import load_workbook
from tkinter import filedialog as fd
from tkinter.scrolledtext import ScrolledText

root = Tk()

root.title('AMI - SMOC SPOT')
root.geometry('580x480')

global md, excl, path, exp01, exp09, exp11, exp51, bcdf, newmpath

def gettemp():
    global template
    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    Label(root, text = filenames[0]).place(x = 200, y  = 52)
    
    template = filenames[0]
    return template

def getmd():
    global md, path
    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    Label(root, text = filenames[0]).place(x = 200, y  = 92)
    
    path = filenames[0]
    
    md = pd.read_excel(path, sheet_name = 'Sheet1')
    md = md.drop(columns = ['Advanced Metering System','Advanced Meter Capability Grp (AMCG)','Reg.'])
    md = md.rename(columns = {'Equipment' : 'Meter ID'})
    md['MR Unit'] = md['MR Unit'].astype(str)
    md['Installat.'] = md['Installat.'].astype(str)
    
    return md, path

def getexcl():
    global excl
    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    Label(root, text = filenames[0]).place(x = 200, y  = 132)
    
    exclpath = filenames[0]
    excl = pd.read_excel(exclpath)
    excl['Remarks'] = 'Check'
    excl = excl[['Meter ID','Remarks']]
    
    return excl

def getseg():
    global bcdf
    filenames = fd.askopenfilenames(filetypes=[("Text files","*.xlsx")])
    Label(root, text = filenames[0]).place(x = 200, y  = 172)
    
    bcrm = filenames[0]
    bcdf = pd.read_excel(bcrm)
    return bcdf

# what if total installation exceeds limit? - check again what the limit is - add a filtering parameter? 
# meterdata file needs to be named as the correct spot
# save as binary file in the end
# get rid of decimal
# MR unit based on segregated data 
# change CA to number


def expfile():

    global exp01, exp09, exp11, exp51
    name = fd.askdirectory()
    
    path01 = name + '/export01.csv'
    path09 = name + '/export09.csv'
    path11 = name + '/export11.csv'
    path51 = name + '/export51.csv'
    
    #read export files
    exp01 = pd.read_csv(path01)
    exp09 = pd.read_csv(path09)
    exp11 = pd.read_csv(path11)
    exp51 = pd.read_csv(path51)

    exp01 = exp01.rename(columns = {'METERID' : 'Meter ID', 'FINALREAD' : 'Reg 01 (kWh)','FINALREADSTATUS' : 'Reading Status', 'AVERAGECONSUMPTION_NAME':'Avg. Consumption' })
    exp01 = exp01.loc[exp01['Reading Status'] == 'VAL', ['Meter ID','Reg 01 (kWh)','Reading Status','Avg. Consumption']]
    exp01['Reg 01 (kWh)'] = exp01['Reg 01 (kWh)'].apply(np.floor).astype('Int64')

    # merge exports with meterdata dataframe
    exp09 = formatexp(exp09, 'Reg 09 (kWh)')
    exp11 = formatexp(exp11, 'Reg 11 (kWh)')
    exp51 = formatexp(exp51, 'Reg 51 (kWh)')
    
    Label(root, text = name).place(x = 200, y  = 212)
    
def formatexp(df, reg): #for exp9,11,51
    
    global md
    
    df = df.rename(columns = {'METERID' : 'Meter ID', 'FINALREAD' : reg})
    # df = df.drop(df.columns.difference(['Meter ID',reg]),1)
    df = df.loc[df['FINALREADSTATUS'] == 'VAL',['Meter ID',reg]]
    df[reg] = df[reg].apply(np.floor).astype(int)
    
    # md = pd.merge(left = md, right = df, how = 'left')
    return df
    
def genmd():
    global newmpath, md
    
    folder = '\\'.join(path.split('/')[:-1])
    spot = path.split('/')[-1].split('.')[0] 

    newmpath = folder + '\\' + spot + '_MD.xlsx'
    writer = pd.ExcelWriter(newmpath, engine = 'xlsxwriter')
    
    # merge exports with meterdata dataframe
    md = pd.merge(left = md, right = exp01, how = 'left')
    md = pd.merge(left = md, right = exp09, how = 'left')
    md = pd.merge(left = md, right = exp11, how = 'left')
    md = pd.merge(left = md, right = exp51, how = 'left')
    
    #md[['Reg 01 (kWh)','Reg 09 (kWh)','Reg 11 (kWh)','Reg 51 (kWh)']] = md[['Reg 01 (kWh)','Reg 09 (kWh)','Reg 11 (kWh)','Reg 51 (kWh)']].astype(str)
    # for i in range(len(md)):
        # if pd.isna(md.loc[i,'Reg 51 (kWh)']) is True:
            # md.loc[i,'Reading Status'] = '#N/A'

    #checks for incomplete register
    md = md[['MR Unit','Portion','Installat.','Meter ID','Reg 01 (kWh)','Reg 09 (kWh)','Reg 11 (kWh)','Reg 51 (kWh)','Reading Status','Avg. Consumption']]    
    
    md.loc[(pd.isnull(md['Reg 01 (kWh)'])) | (pd.isnull(md['Reg 09 (kWh)'])) | (pd.isnull(md['Reg 11 (kWh)'])) | (pd.isnull(md['Reg 51 (kWh)'])), 'Reading Status'] = '#N/A'
    md.loc[(md['Reading Status'] != 'VAL'), ['Reg 01 (kWh)','Reg 09 (kWh)','Reg 11 (kWh)','Reg 51 (kWh)']] = np.nan

    md.to_excel(writer, sheet_name = 'Meter Data', index = False)

    # make another dataframe with reg 51 > 2 & not kwh forwarded
    reg51 = md.loc[(md['Reg 51 (kWh)'] >= 2) & (md['Avg. Consumption'] == 'kWh received'), ['Meter ID','Reg 51 (kWh)','Reg 01 (kWh)','Avg. Consumption']]
    check = reg51.loc[(reg51['Meter ID'].str.startswith('HOL')) | (reg51['Meter ID'].str.startswith('EDM')) | (reg51['Meter ID'].str.startswith('SAG100')) | (reg51['Meter ID'].str.startswith('SAG101'))]
    check = pd.merge(left = check, right = excl, how = 'left')
    check.to_excel(writer, sheet_name = 'Check', index = False)
    
    # chprint = check.loc[check['Meter ID'] == 'VAL', 'Meter ID']

    writer.close()
    
    textbox.delete('1.0', END)
    # if chprint:   
        # textbox.insert(END, 'Please check meter(s) on BCRM')
        # textbox.insert(END, '\n') 
        # textbox.insert(END, chprint) 

def genspot():
    global bcdf, newmpath, path
    
    units = []
    
    newmd = pd.read_excel(newmpath, sheet_name = 'Meter Data')
    newmd['MR Unit'] = newmd['MR Unit'].astype(str)
    newmd['Installat.'] = newmd['Installat.'].astype(str)
    
    folder = '\\'.join(path.split('/')[:-1])
    spot = path.split('/')[-1].split('.')[0] 
    
    for i in newmd['MR Unit']:
        if i[:3] not in units:
            units.append(i[:3])

    bcdf['MR Unit'] = bcdf['MR Unit'].astype(str)
    bcdf['Installat.'] = bcdf['Installat.'].astype(str)
    bcdf = bcdf.drop(columns = ['Meter Reading Date','Valid fr.','Valid to','IS'])
    
    textbox.delete('1.0', END)
    
    for i in units:
        md2 = newmd[newmd['MR Unit'].str.startswith(i)] 
        bcdf2 = bcdf[bcdf['MR Unit'].str.startswith(i)]

        # get meter ID
        val = pd.merge(left = bcdf2, right = md2, how = 'left', on = 'Installat.')
        val = val.drop(columns =['MR Unit_y','Portion_y'])
        val = val.rename(columns = {'MR Unit_x' : 'MR Unit','Portion_x' : 'Portion'})
        
        #reorder columns
        val['Remarks'] = ''
        val = val[['MR Unit','Contract Account','Sequence Number','Portion','Installat.','Meter ID','Address','Reg 01 (kWh)','Reg 09 (kWh)','Reg 11 (kWh)','Reg 51 (kWh)','Remarks','Reading Status']]
        val = val.drop_duplicates()
        
        val[['Reg 01 (kWh)','Reg 09 (kWh)','Reg 11 (kWh)','Reg 51 (kWh)']] = val[['Reg 01 (kWh)','Reg 09 (kWh)','Reg 11 (kWh)','Reg 51 (kWh)']].astype(str)
        # val.loc[(val['Reg 51 (kWh)'].isnumeric() == False), 'Reading Status'] = '#N/A' #| (md['Reg 09 (kWh)'] == None) | (md['Reg 11 (kWh)'] )  | (md['Reg 51 (kWh)'].isnumeric() == False) , 'Reading Status'] = '#N/A'

        #clean up non-val rows
        val.loc[(val['Reading Status'] != 'VAL'), ['Reg 01 (kWh)','Reg 09 (kWh)','Reg 11 (kWh)','Reg 51 (kWh)']] = '#N/A'
        val.loc[(val['Reading Status'] != 'VAL'), 'Remarks'] = 'Unsuccessful Reading from System'
        val.loc[(val['Reading Status'] != 'VAL'), 'Reading Status'] = '#N/A'
        
        # will go into same directory as spot folder
        dest = folder + '\\' + spot + ' 6' + i + ' MAR.xlsx' #change spot
        shutil.copy(template, dest)

        writer = pd.ExcelWriter(dest, mode = 'a', engine = 'openpyxl')
        writer.book = load_workbook(dest)
        
        writer.book.remove(writer.book['VAL'])
        writer.book.remove(writer.book['METER ID'])
        
        val.to_excel(writer, sheet_name = 'VAL', index = False)
        
        writer.close()
        
        valid = val.loc[val['Reading Status'] == 'VAL']
        noval = val.loc[val['Reading Status'] != 'VAL']
        
        textbox.insert(END, 'Station: ' + str(i))
        textbox.insert(END, '\n')        
        textbox.insert(END, 'Total Accounts: ' + str(len(val)))
        textbox.insert(END, '\n')
        textbox.insert(END, 'Reading from MDMS: ' + str(len(valid)))
        textbox.insert(END, '\n')
        textbox.insert(END, 'No Reading: ' + str(len(noval)))
        textbox.insert(END, '\n')
        textbox.insert(END, '\n')

#uploads
Button(root, text = 'Template SPOT', command = gettemp).place(x = 59, y = 50)
Button(root, text = 'Meter Data BCRM', command = getmd).place(x = 45, y = 90)
Button(root, text = 'Exclusion File', command = getexcl).place(x = 63, y = 130)
Button(root, text = 'Segregated Data', command = getseg).place(x = 52, y = 170)
Button(root, text = 'MyCloud Folder', command = expfile).place(x = 52, y = 210)

Label(root, text = 'Path: ').place(x = 160, y  = 52)
Label(root, text = 'Path: ').place(x = 160, y  = 92)
Label(root, text = 'Path: ').place(x = 160, y  = 132)
Label(root, text = 'Path: ').place(x = 160, y  = 172)
Label(root, text = 'Path: ').place(x = 160, y  = 212)

textvar = StringVar()
textbox = ScrolledText(root, height = 8, width = 45)
textbox.place(x = 110, y = 300)

Button(root, text = 'Generate Meter Data', command = genmd).place(x = 170, y = 260)
Button(root, text = 'Generate SPOT Files', command = genspot).place(x = 300, y = 260)


root.mainloop()