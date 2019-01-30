import xlrd, datetime, sys

from datetime import timedelta
file_loc=input('Enter the Tracker location: ')

loc = (file_loc)
startdatelist = []
distsysCsid, distuatCsid, distoatCsid, distliveCsid, setDate = set(), set(), set(), set(), set()
listsysCsid, listDate, listDate2 = [], [], []
setsysCsid, setuatCsid, setoatCsid, setliveCsid = set(), set(), set(), set()
sysCount, uatCount, oatCount, liveCount = 0, 0, 0, 0
finalsyscountcsid, finaluatcountcsid, finaloatcountcsid, finallivecountcsid = 0, 0, 0, 0
livecsid, syscsid, uatcsid, oatcsid = set(), set(), set(), set()
syscomponent, uatcomponent, oatcomponent, livecomponent = [], [], [], []

# Pending components count
pendSyscsid, pendUatcsid, pendOatcsid, pendLivecsid = set(), set(), set(), set()
pendSyscomponent, pendUatcomponent, pendOatcomponent, pendLivecomponent = [], [], [], []

lastrow = int(input('Enter the last row number for the excel : '))

st1 = str(input("Enter the start date in format dd/mm/yy: "))
start_date=datetime.datetime.strptime(st1, '%d/%m/%y')
for i in range(0, 5):
    modified_date = datetime.datetime.strftime(start_date, "%d/%m/%y")
    startdatelist.append(modified_date)
    start_date = start_date + timedelta(days=1)
    #print(datetime.datetime.strftime(modified_date, "%d/%m/%y"))

book = xlrd.open_workbook(loc)
sheet = book.sheet_by_index(0)
print(sorted(startdatelist))

for date in sorted(startdatelist):
    for rowx in range(0, lastrow):
        datecell = sheet.cell_value(rowx, colx=0)
        try:
            # print(datecell)
            datecell_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(datecell, book.datemode))
            # print(str(datecell_as_datetime))
            # print(datetime.date(startdate))

            if (str(date) == str(datecell_as_datetime.strftime('%d/%m/%y'))):
                # print('datetime: %s' % datecell_as_datetime.strftime('%d/%m/%y'))
                # print(sheet.cell_value(rowx, colx=3))
                if 'Deployed in SYS' in sheet.cell_value(rowx, 4):
                    syscsid.add(sheet.cell_value(rowx, 2))
                    syscomponent.append(sheet.cell_value(rowx, 3))
                    # print('SYS is ', (syscsid))

                if 'Deployed in UAT' in sheet.cell_value(rowx, 4):
                    uatcsid.add(sheet.cell_value(rowx, 2))
                    uatcomponent.append(sheet.cell_value(rowx, 3))
                    # print('UAT is ', uatcsid)

                if 'Deployed in OAT' in sheet.cell_value(rowx, 4):
                    oatcsid.add(sheet.cell_value(rowx, 2))
                    oatcomponent.append(sheet.cell_value(rowx, 3))
                    # print('OAT is ', oatcsid)

                if 'Deployed in LIVE' in sheet.cell_value(rowx, 4):
                    livecsid.add(sheet.cell_value(rowx, 2))
                    livecomponent.append(sheet.cell_value(rowx, 3))
                    # print('Live is ', livecsid)

                if sheet.cell_value(rowx, 4) in ['Deployed in SYS, OAT' , 'Deployed in SYS,OAT'] :
                    syscsid.add(sheet.cell_value(rowx, 2))
                    syscomponent.append(sheet.cell_value(rowx, 3))
                    oatcsid.add(sheet.cell_value(rowx, 2))
                    oatcomponent.append(sheet.cell_value(rowx, 3))
                    # print('Live is ', livecsid)


                if sheet.cell_value(rowx, 4) in ['Deployed in SYS, UAT' ,'Deployed in SYS,UAT']:
                    syscsid.add(sheet.cell_value(rowx, 2))
                    syscomponent.append(sheet.cell_value(rowx, 3))
                    uatcsid.add(sheet.cell_value(rowx, 2))
                    uatcomponent.append(sheet.cell_value(rowx, 3))
                     # print('Live is ', livecsid)



                if sheet.cell_value(rowx, 4) in ['Deployed in UAT, OAT', 'Deployed in UAT,OAT']:
                    uatcsid.add(sheet.cell_value(rowx, 2))
                    uatcomponent.append(sheet.cell_value(rowx, 3))
                    oatcsid.add(sheet.cell_value(rowx, 2))
                    oatcomponent.append(sheet.cell_value(rowx, 3))
                     # print('Live is ', livecsid)


                if sheet.cell_value(rowx, 4) in ['Deployed in SYS,UAT,OAT', 'Deployed in SYS, UAT, OAT']:
                    syscsid.add(sheet.cell_value(rowx, 2))
                    syscomponent.append(sheet.cell_value(rowx, 3))
                    uatcsid.add(sheet.cell_value(rowx, 2))
                    uatcomponent.append(sheet.cell_value(rowx, 3))
                    oatcsid.add(sheet.cell_value(rowx, 2))
                    oatcomponent.append(sheet.cell_value(rowx, 3))
                     # print('Live is ', livecsid)


                if sheet.cell_value(rowx, 4) in ['Pending in SYS', 'Error in SYS']:
                    pendSyscsid.add(sheet.cell_value(rowx, 2))
                    pendSyscomponent.append(sheet.cell_value(rowx, 3))
                    # print('SYS is ', syscsid)

                if sheet.cell_value(rowx, 4) in ['Pending in UAT', 'Error in UAT']:
                    pendUatcsid.add(sheet.cell_value(rowx, 2))
                    pendUatcomponent.append(sheet.cell_value(rowx, 3))
                    # print('UAT is ', uatcsid)

                if sheet.cell_value(rowx, 4) in ['Pending in OAT', 'Error in OAT']:
                    pendOatcsid.add(sheet.cell_value(rowx, 2))
                    pendOatcomponent.append(sheet.cell_value(rowx, 3))
                    print('OAT is ', oatcsid)

                if sheet.cell_value(rowx, 4) in ['Pending in LIVE', 'Error in LIVE']:
                    pendLivecsid.add(sheet.cell_value(rowx, 2))
                    pendLivecomponent.append(sheet.cell_value(rowx, 3))



        # startdate = startdate + datetime.timedelta(days=1)
        except:
            print('', end='')
    finalsyscountcsid = finalsyscountcsid + len(syscsid)
    finaluatcountcsid = finaluatcountcsid + len(uatcsid)
    finaloatcountcsid = finaloatcountcsid + len(oatcsid)
    finallivecountcsid = finallivecountcsid + len(livecsid)
    # print('SYS new is ', finalsyscountcsid)
    # syscsid=[]
    syscsid.clear()
    uatcsid.clear()
    oatcsid.clear()
    livecsid.clear()

# print(len(livecsid),len(syscsid),len(uatcsid),len(oatcsid))
# print(syscsid)
print("Deployed Components", end='\n\n')
print("SYS      UAT         OAT         LIVE")
print(finalsyscountcsid, '\t\t', finaluatcountcsid, '\t\t\t', finaloatcountcsid, '\t\t', finallivecountcsid)
# print(len(syscsid),'\t\t', len(uatcsid),'\t\t\t',len(oatcsid), '\t\t\t', len(livecsid))
print(len(syscomponent), '\t\t', len(uatcomponent), '\t\t\t', len(oatcomponent), '\t\t', len(livecomponent))

print("--------------------------------------------------------")
print("Pending Components", end='\n\n')
print("SYS      UAT         OAT         LIVE")
# print(len(pendSyscsid),'\t\t', len(pendUatcsid),'\t\t\t',len(pendOatcsid), '\t\t\t', len(pendLivecsid))
print(len(pendSyscomponent), '\t\t', len(pendUatcomponent), '\t\t\t', len(pendOatcomponent), '\t\t',
      len(pendLivecomponent))
