
from openpyxl import load_workbook as load, Workbook
from functions import findRowNumber, insertRow

if __name__ == '__main__':

    #stores list of files to be read in
    fileList = []
    #stores the row in which each station is in. used to find the row in rowlist
    master = {}

    file = 'MasterSpread/CompInventory.xlsx'
    #loads master file workbook
    masterFile = load(filename= file)
    masterSheet = masterFile.active

    #appends files to file list with structure (path, dept, loc, building)
    with open('toBeMappedFiles.txt') as f:
        for line in f:
            fileList.append(line)

    print(fileList)
    print('Loading Master File')
    #loads info into rowlist and master from master file. info in master file
    #starts in row 6
    rowSkip = 1
    for row in masterSheet:
        if rowSkip < 6:
            rowSkip = rowSkip + 1
        else:
            #rowlist.append(row)
            master[row[6].value] = row


    print('File loaded')
    print('Starting Batch...')
    for line in fileList:


        #stores all devices that need to be updated/added in/to the master file.
        devices = []
        #stores the dept name for each station
        dept = {}
        #stores the location name for each station
        loc = {}
        #stores building for each station
        build = {}
        #stores the map key for each station
        mapKey = {}
        #stores the printers associated with each station
        printer = {}

        file = line.split(',')
        print(file)
        fileTarget = file[0]
        deptName = file[1]
        location = file[2]
        building = file[3].split('\n')[0]
        print(building)

        print('Loading File: ', file[1])
        # loads tech's file
        map = load(filename= fileTarget)
        #sets active sheet for both workbooks
        mapSheet = map.active

        print('Finished Loading')
        #variables to map values in map file
        key = 0
        station = ''
        prnt = ''
        print('Updating variables')
        #map file should have file format of (key, station, printer)
        for row in mapSheet:
            key = row[0].value
            station = row[1].value
            devices.append(station)
            prnt = row[2].value

            dept[station] = deptName
            build[station] = building
            loc[station] = location
            printer[station] = prnt
            mapKey[station] = key

            if printer[station] not in devices:
                devices.append(printer[station])

        print('Finished')

        #adds map key to printer devices and adds printers to printer dictionary.
        #resets key variable to zero and then increments until a key is available
        keyList = []
        key = 0
        for item in mapKey:
            keyList.append(mapKey[item])

        keyList.sort()
        print('adding printers')

        for device in devices:
            try:
                testIfHasKey = mapKey[device]
            except:
                #trys to find available key in key list
                key = 0
                noKey = False
                #sorts lists to maintain list numerically
                keyList.sort()
                for x in range(len(keyList)):
                    if x not in keyList:
                        key = x
                        mapKey[device] = key
                        keyList.append(key)
                        dept[device] = deptName
                        loc[device] = location
                        build[device] = building
                        printer[device] = 'N/A'
                        break
                    elif x == len(keyList) - 1:
                       noKey = True
                if noKey:
                    key = keyList[-1] + 1
                    mapKey[device] = key
                    keyList.append(key)
                    dept[device] = deptName
                    loc[device] = location
                    build[device] = building
                    printer[device] = 'N/A'
        print('Finished')
        #list containing devices that are not in master file
        notInMaster = []

        print('updating data')
        #updates devices in master file with info from map file
        for device in devices:
        #master file format should be (dept, build,,loc,,map,,,,,,,,,,,,,,,,,printer)
            try:
                #checks to see if device is in master file. errors out if not
                testIfInMaster = master[device]

                #succeeded in finding device in master file
                master[device][0].value = dept[device]
                master[device][1].value = build[device]
                master[device][3].value = loc[device]
                master[device][5].value = mapKey[device]
                master[device][22].value = printer[device]
                #update type of device later in one go due to inserting rows to sheet
            except:
                #device not in master file
                notInMaster.append(device)
                print(device, ': not in Master')

        #inserts rows to master file and updates the info with map file data.
        #updates rowlist to correspond to

        notInMaster.sort()
        print(notInMaster)
        print('adding devices not in master')

        index = 0
        #print(notInMaster)
        rowNum = 1
        for device in notInMaster:
            #finds where the device belongs alphabetically
            index = findRowNumber(masterSheet, device)
            print(index, ': ', device)
            #inserts the row above the found row
            masterSheet.insert_rows(index)
            tempRow = masterSheet[index]
            master[device] = tempRow

            master[device][0].value = dept[device]
            master[device][1].value = build[device]
            master[device][3].value = loc[device]
            master[device][5].value = mapKey[device]
            master[device][6].value = device
            master[device][22].value = printer[device]

            rowNum = rowNum + 1

            #insertRow(rowlist, index, tempRow)
        #for dev in notInMaster:
            #print(master[dev][0])

        typeRow = 1
        for row in masterSheet:
            #print(typeRow, row[0])
            try:
                if typeRow < 6:
                    #print(typeRow, ': skipped')
                    typeRow = typeRow + 1
                else:
                    typeFormat = '=IF(ISBLANK($G'+ str(typeRow) +') ,"", IF(OR(ISNUMBER(SEARCH("WSL",$G'+ str(typeRow) +')),ISNUMBER(SEARCH("WSC",$G'+ str(typeRow) +')),ISNUMBER(SEARCH("MFP",$G'+ str(typeRow) +')),ISNUMBER(SEARCH("PRT",$G'+ str(typeRow) +')),ISNUMBER(SEARCH("TC",$G'+ str(typeRow) +'))),IF(OR(ISNUMBER(SEARCH("WSC",$G'+ str(typeRow) +')),ISNUMBER(SEARCH("WSL",$G'+ str(typeRow) +')),ISNUMBER(SEARCH("TC",$G'+ str(typeRow) +'))),IF(ISNUMBER(SEARCH("TC",$G'+ str(typeRow) +')),"Zero","Laptop"),"Printer"),"Workstation"))'
                    #print(typeRow, ': updating')
                    row[7].value = typeFormat
                    typeRow = typeRow + 1
            except:
                print('Row: ', typeRow, ' contains merged cells')
        print('Finished')
        print('Saving...')
        masterFile.save('MasterSpread/CompInventoryTestOutput.xlsx')
        print('Saved')
