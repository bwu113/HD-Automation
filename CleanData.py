import Process

packingList = {}
constList = ["orderby", "po", "co", "name", "street", "cityzip", "phone", "sku", "desc", "qty", "sku2", "desc2", "qty2"]
# Clean raw data from output file.
def cleanPackingList(rawData):
    pageList = []
    tempVar = ""
    page = 0
    
    #Cleaning phase 1
    for listItem in rawData.values(): #Cleaning lines using indices and removing excess info
        #print(listItem)
        #listItem.append("#" + str(len(listItem)))
        listItem[1] = listItem[1].split(" ")[2] # Keeps PO# only
        listItem[2] = listItem[2].split(" ")[3] # Keeps Customer Order # Only
        if(len(listItem) == 9):
            tempVar = listItem[7].splitlines()
            if " " in tempVar[0]:
                listItem[7] = tempVar[0].split()[0]
                listItem.insert(8,tempVar[1])
                tempVar = listItem[9].splitlines()
                listItem[8] += " " + tempVar[0]
                listItem[9] = tempVar[1]
            else:
                listItem[7] = tempVar[0]
                listItem.insert(8,tempVar[2])
                tempVar = listItem[9].splitlines()
                listItem[8] += " " + tempVar[0]
                listItem[9] = tempVar[1]
        elif(len(listItem) == 10): #splits description for orders that have left-over sku number and combines split description with left-over description element
            tempVar = listItem[8].splitlines()
            if("New" not in tempVar[0]):
                listItem[7] += tempVar[0]
                listItem[8] = tempVar[2]
                tempVar = listItem[9].splitlines()
                listItem[9] = tempVar[0]
                listItem.insert(10,tempVar[1])
                listItem[8] += " " + listItem[9]
                listItem.remove(listItem[9])
            else:
                listItem[8] = tempVar[2]
                tempVar = listItem[9].splitlines()
                listItem[9] = tempVar[0]
                listItem.insert(10,tempVar[1])
                listItem[8] += " " + listItem[9]
                listItem.remove(listItem[9])
        elif(len(listItem) == 11): #splits lines for 2 sku orders that only have 11 items.
            tempVar = listItem[7].splitlines()
            listItem[7] = tempVar[0]
            listItem.insert(8,tempVar[2])
            tempVar = listItem[9].splitlines()
            listItem[8] += " " + tempVar[0]
            listItem[9] = tempVar[1]
            tempVar = listItem[10].splitlines()
            listItem[10] = tempVar[0]
            listItem.insert(11,tempVar[2])
            tempVar = listItem[12].splitlines()
            listItem[11] += " " + tempVar[0]
            listItem[12] = tempVar[1]
        elif(len(listItem) == 13): #splits descriptions element for orders with 2 skus and combines split description with left-over description element
            tempVar = listItem[8].splitlines() #splits 3 line description element
            listItem[7] += tempVar[0]
            listItem[8] = tempVar[2]
            tempVar = listItem[9].splitlines() #splits leftover description with quantity value
            listItem[8] += " " + tempVar[0]
            listItem[9] = tempVar[1]
            tempVar = listItem[11].splitlines() #splits 3 line description element
            listItem[10] += tempVar[0]
            listItem[11] = tempVar[2]
            tempVar = listItem[12].splitlines() #splits leftover description with quantity value
            listItem[11] += " " + tempVar[0]
            listItem[12] = tempVar[1]
        else:
            listItem.insert(len(listItem),"Error")
        pageList.append(listItem)
        print(listItem)

    #Cleaning Phase 2
    for items in pageList:
        packingList[page] = {constList[0]:items[0]}
        if("#" in items[3]):
            tempVar = items[3].split("-")
            items[3] = tempVar[0]
            for ele in range(1,len(items)):
                if(ele == 4):
                    packingList[page].update({"storenum":"C/O THD Ship to Store " + tempVar[1]})
                packingList[page].update({constList[ele]:items[ele]})
        # elif("Error" in items[len(items)-1]):
        #     packingList[page].update({constList[3]:items[3]})
        #     packingList[page].update({"err":"y"})
        else:
            for ele in range(1,len(items)):
                packingList[page].update({constList[ele]:items[ele]})
        page+=1
    #print(packingList.values())
    # for list in packingList.values():
    #     if("-BS" in list['sku'] or "-MIR" in list['sku']):
    #         splitSKU = list['sku'].split("-")
    #         newSKU = ""
    #         for i in splitSKU:
    #             match i:
    #                 case "BS":
    #                     list.update({"bs":"y"})
    #                 case "MIR":
    #                     list.update({"mir":"y"})
    #                 case _:
    #                     newSKU += i + "-"
    #         list["sku"] = newSKU.rstrip("-")
    #     if("sku2" in list):
    #         if("-BS" in list['sku2'] or "-MIR" in list['sku2']):
    #             newSKU = ""
    #             for i in splitSKU:
    #                 match i:
    #                     case "BS":
    #                         list.update({"bs2":"y"})
    #                     case "MIR":
    #                         list.update({"mir2":"y"})
    #                     case _:
    #                         newSKU += i + "-"
    #             list["sku2"] = newSKU.rstrip("-")
    #     print(list)

    Process.insertExcel(packingList)