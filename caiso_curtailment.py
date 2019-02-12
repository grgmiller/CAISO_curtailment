import csv
import os
import pandas as pd
from pathlib import Path
import tabula
from datetime import datetime, timedelta

curtailURL = "http://www.caiso.com/informed/Pages/ManagingOversupply.aspx#dailyCurtailment"
pdfFile = 'http://www.caiso.com/Documents/Wind_SolarReal-TimeDispatchCurtailmentReport'
extraPageDate = datetime.strptime('04/13/2017', '%m/%d/%Y')
latestDateFile = Path.cwd() / 'latestdate.txt'

dataFile = Path.cwd() / 'curtail_report.csv'
if not dataFile.exists(): #create the csv file if it doesnt yet exist
    with open(dataFile, 'w', newline="") as f:
        w = csv.writer(f)
        w.writerow(['month','day','year','hour','curtailment_type','reason','curtail_category','fuel_type','curtailed_MWh','curtailed_MW'])
    print('  Created blank csv file')


#create list of dates to retrieve
def latest(): #return the date to start retreiving data
    if not Path(latestDateFile).exists():
        with open(latestDateFile,'w+') as f:
            latestDate = 'Jun30_2016'  #earliest available is June 30, 2016
            f.write(latestDate)
            return latestDate
    else:
        with open(latestDateFile, 'r') as f:
            latestDate = f.readline()
            latestDate_dt = datetime.strptime(latestDate, '%b%d_%Y') + timedelta(days=1)
            latestDate = datetime.strftime(latestDate_dt, '%b%d_%Y')
            return latestDate

latestDate = latest()
datelist = [] #create empty list
datelist.append(latestDate) #add the first date to the list

filedate_dt = datetime.strptime(latestDate, '%b%d_%Y')
yesterday = datetime.now() - timedelta(days=1)

while filedate_dt.date() < yesterday.date(): #create a list of dates to retrieve, starting with the latest downloaded, and ending with yesterday
    filedate_dt = filedate_dt + timedelta(days=1)
    date = datetime.strftime(filedate_dt, '%b%d_%Y')
    datelist.append(date)

#update latest date file with last value from datelist
with open(latestDateFile, 'w') as f:
    f.seek(0)
    f.write(datelist[-1])

for date in datelist:
    date_dt = datetime.strptime(date, '%b%d_%Y')
    # need to add logic for files before April 13, 2017
    if date_dt < extraPageDate: #if the date is before they started adding a third summary page before the chart
        try: #see if the file has data on pages 3 and 4, otherwise only extract from page 4
            df = tabula.read_pdf(pdfFile + date + '.pdf', pages='3-4', lattice=True, java_options=['-Dsun.java2d.cmm=sun.java2d.cmm.kcms.KcmsServiceProvider'])
        except: 
            try:
                df = tabula.read_pdf(pdfFile + date + '.pdf', pages='3', lattice=True, java_options=['-Dsun.java2d.cmm=sun.java2d.cmm.kcms.KcmsServiceProvider'])
            except:
                df = tabula.read_pdf(pdfFile + date + '.pdf', pages='2', lattice=True, java_options=['-Dsun.java2d.cmm=sun.java2d.cmm.kcms.KcmsServiceProvider'])
    else: 
        try: #see if the file has data on pages 4 and 5, otherwise only extract from page 4
            df = tabula.read_pdf(pdfFile + date + '.pdf', pages='4-5', lattice=True, java_options=['-Dsun.java2d.cmm=sun.java2d.cmm.kcms.KcmsServiceProvider'])
        except: 
            df = tabula.read_pdf(pdfFile + date + '.pdf', pages='4', lattice=True, java_options=['-Dsun.java2d.cmm=sun.java2d.cmm.kcms.KcmsServiceProvider'])

    # need to add column for year (or just replace date column with strptime(date))
    df.columns = (['date','hour','curtailment_type','reason','fuel_type','curtailed_MWh','curtailed_MW'])
    df['month'] = date_dt.month #add a column with the month number
    df['day'] = date_dt.day #add a column with the day number
    df['year'] = date_dt.year #add a column with the year number
    df['curtail_category'] = df['curtailment_type'] + ' - ' + df['reason']
    df.drop(['date'], axis=1, inplace=True) #drop the date column
    df = df[['month','day','year','hour','curtailment_type','reason','curtail_category','fuel_type','curtailed_MWh','curtailed_MW']] #re-order columns

    with open(dataFile,'a') as f:    #append dataframe to dataFile CSV
        df.to_csv(f, header=False, index=False)
    print('  {} appended to csv'.format(date))





#download all past 5-min curtailment data
def download_wait(f): #wait for files to finish downloading before continuing
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < 20:
        time.sleep(1) #check every sec
        dl_wait = False
        for fname in os.listdir(Path.cwd() / f):
            if fname.endswith('.crdownload'): #incomplete chrome downloads end in .crdownload
                dl_wait = True
            seconds += 1
    time.sleep(1) #allow 1 sec after downloading

def downloadCurtailment(browser, user_initialized): #download curtailment data (updated monthly)
    print('  Checking for new curtailment data...')
    browser.get(curtailURL) #open webdriver
    time.sleep(1) #wait for page to load
    soup = BeautifulSoup(browser.page_source, 'lxml') #use beautifulsoup to parse html
    postDate = soup.find_all('span', class_='postDate')[0].get_text() #get current postDate from site
    with shelve.open(str(shelf)) as s:
        prevPostDate = s['caiso']['postDate']
    if postDate==prevPostDate: #compare current and previous postdate
        print('  Latest curtailment data already downloaded.') #do nothing if they match; we already have the most current file
    else: #download new curtailment file if more recent data is available
        tmpDelete('downloads') #clear downloads folder
        tmpDelete('curtailments') #delete existing file in curtailments folder
        browser.find_elements_by_partial_link_text('Production and Curtailments Data')[0].click() #download file
        if user_initialized==0: #only notify of new curtailment download if not initiatied by the user
            print('  New curtailment data available!')
        print('  Downloading curtailment Excel file...')
        download_wait('downloads')         #wait for download to finish
        with shelve.open(str(shelf), writeback=True) as s:
            s['caiso']['postDate'] = postDate
        curtailFile = os.listdir(downloads)[0]
        os.rename(downloads / curtailFile, curtailments / curtailFile)  #move file to curtailments directory
        
        #convert the excel downloads to csv
        
        print('  Converting Excel file to CSV. This may take several minutes...')
        wb = openpyxl.load_workbook('curtailments/'+curtailFile) #this step takes a couple minutes to fully load
        sh = wb['Curtailments'] 
        with open(curtailments / 'curtailment_data.csv', 'w', newline="") as f:  #convert xlsx to csv file for faster reading in future
            c = csv.writer(f)
            for r in sh.rows:
                if r[0].value is not None:
                    c.writerow([cell.value for cell in r])
                else: 
                    continue
        time.sleep(1) #pause 1 sec after csv file created
        os.remove(curtailments / curtailFile) #once the new csv file is created, delete the xlsx file
        curtail_read = pd.read_csv(curtailments / 'curtailment_data.csv', dtype=ct_dtypes) #load the csv into a dataframe
        curtail_read.columns = (['date','hour', 'interval','wind_curtail_MW','solar_curtail_MW']) #rename columns

            






#merge reason data with 5-min data

