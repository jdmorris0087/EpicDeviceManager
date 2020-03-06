from openpyxl import load_workbook as load, Workbook

#finds row in ws given a workstation
def findRowNumber(ws, station):
    test = ''
    i = 1
    for row in ws:
        if i < 6:
            i = i + 1
        else:
            try:
                #row[6] can contatin None values which is not comparable to string values
                #tests if test.value contains none value
                test = row[6]
                if test.value < station:
                    i = i + 1
                else:
                    return i
            except:
                i = i + 1

if __name__ == '__main__':

    #stores list of files to be read in
    fileList = []
    #stores the row in which each station is in. used to find the row in rowlist
    master = {}

    #appends files to file list with structure (path, dept, loc, building)
    with open('toBeMappedFiles.txt') as f:
        for line in f:
            fileList.append(line)

    file = fileList.pop(0)
    file = file.split('\n')[0]
    #loads master file workbook
    masterFile = load(filename= file)
    masterSheet = masterFile.active

    print('Loading Master File')
    #loads info into rowlist and master from master file. info in master file
    #starts in row 6
    rowSkip = 1
    for row in masterSheet:
        if rowSkip < 6:
            rowSkip = rowSkip + 1
        else:
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
        fileTarget = file[0]
        deptName = file[1]
        location = file[2]
        #the last value of line contains carriage return needs to be split
        building = file[3].split('\n')[0]

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
            prnt = row[2].value
            if station not in devices:
                devices.append(station)
                dept[station] = deptName
                build[station] = building
                loc[station] = location
                printer[station] = prnt
                mapKey[station] = key

                if printer[station] not in devices:
                    devices.append(printer[station])
            else:
                pass

        print('Finished')

        #adds map key to printer devices and adds printers to printer dictionary.
        #resets key variable to zero and then increments until a key is available
        keyList = []
        key = 0
        noList = False
        for item in mapKey:
            keyList.append(mapKey[item])
        try:
            keyList.sort()
        except:
            noList = True
        print('adding printers')

        for device in devices:
            try:
                #throws KeyError if not found in Dictionary
                testIfHasKey = mapKey[device]
            except:
                #trys to find available key in key list
                key = 0
                noKey = False
                #sorts lists to maintain list numerically
                if not noList:
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
                    #if x reaches the end of the used keys, it sets noKey to True
                    elif x == len(keyList) - 1:
                       noKey = True
                #check noKey, if true finds the last value in the sorted list and adds one.
                #it then updates the values in the other dictionaries for mapping purposes later on.
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
                #update type of device later in one go due to inserting rows to sheet using excel formulas
            except:
                #device not in master file
                notInMaster.append(device)
                print(device, ': not in Master')

        #inserts rows to master file using index found and updates the info with map file data.

        #sorted so that the entries are added to excel in descending order. keeps row info more consistent
        notInMaster.sort()
        print('adding devices not in master')

        index = 0
        for device in notInMaster:
            #finds where the device belongs alphabetically
            index = findRowNumber(masterSheet, device)
            #inserts the row above the found row
            masterSheet.insert_rows(index)
            #adds row to master dictionary for ease of updating
            tempRow = masterSheet[index]
            master[device] = tempRow

            master[device][0].value = dept[device]
            master[device][1].value = build[device]
            master[device][3].value = loc[device]
            master[device][5].value = mapKey[device]
            master[device][6].value = device
            master[device][22].value = printer[device]

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
                elif row[7].value == 'Cart Workstation':
                    typeRow = typeRow + 1
                else:
                    typeFormat = '=IF(ISBLANK($G'+ str(typeRow) +') ,"", IF(OR(ISNUMBER(SEARCH("WSL",$G'+ str(typeRow) +')),ISNUMBER(SEARCH("WSC",$G'+ str(typeRow) +')),ISNUMBER(SEARCH("MFP",$G'+ str(typeRow) +')),ISNUMBER(SEARCH("PRT",$G'+ str(typeRow) +')),ISNUMBER(SEARCH("TC",$G'+ str(typeRow) +'))),IF(OR(ISNUMBER(SEARCH("WSC",$G'+ str(typeRow) +')),ISNUMBER(SEARCH("WSL",$G'+ str(typeRow) +')),ISNUMBER(SEARCH("TC",$G'+ str(typeRow) +'))),IF(ISNUMBER(SEARCH("TC",$G'+ str(typeRow) +')),"Zero","Laptop"),"Printer"),"Workstation"))'
                    #print(typeRow, ': updating')
                    row[7].value = typeFormat
                    typeRow = typeRow + 1
            except:
                print('Row: ', typeRow, ' contains merged cells')
                typeRow = typeRow + 1
        print('Finished')
        print('Saving...')
        #test file. does not override actual master file. needs to be changed after testing
        masterFile.save('\\\\fileserver\\Shares\\Users\\PCTechs\\Nathaniel\\Current Projects\\Epic Inventory\\Epic Computer Inventory.xlsx')
        print('Saved')
