import argparse

from apiclient.discovery import build

"""
simple class for reading/write with a single Google Sheets cell programmatically.

would be great to add a flag to force a type of the update...for isntance, so one could
update a formula externally, which is valueType "formulaValue" not "stringValue"...or
so one could read in cmdline args and know if '99' is supposed to be a string or an int.
"""

class Cellmate(object):
    spreadsheets = None

    def __init__(self):
        service = build('sheets', 'v4')
        self.spreadsheets = service.spreadsheets()

    # returns 0 for a, 25 for z
    def getLetterIndex(self, letter):
        return ord(letter[0].upper()) - 65

    # given a cell like "B3" breaks that into row/col
    def convertCellToRowCol(self, cell):
        row = -1
        col = -1

        colPortion = ""
        idx = 0
        for c in cell:
            if (c.isdigit()):
                row = int(cell[idx:]) - 1
                break
            else:
                colPortion = colPortion + c
            idx = idx + 1

        if (row == -1 or colPortion == ""):
            raise Exception("invalid cell '" + str(cell) + "' specified")

        if (len(colPortion) > 1): 
            raise Exception("being lazy; we don't support cols beyond Z")

        col = self.getLetterIndex(colPortion)

        return { "row": row, "col": col }

    def getValue(self, spreadsheetId, sheetName, cell):
        sheetRange = sheetName + "!" + cell + ":" + cell
        resp = self.spreadsheets.get(spreadsheetId=spreadsheetId, ranges=sheetRange, includeGridData=True).execute()
        valueObj = resp["sheets"][0]["data"][0]["rowData"][0]["values"][0]

        # three possible keys - formattedValue is what is displayed. userEnteredValue.???Value is what user
        # actually entered - for instance, a formula. Also effecitveValue.???Value.
        # We'll just make do with formattedValue for now.
        return valueObj["formattedValue"]

    # stuff the value into the specified doc/sheet/cell
    def setValue(self, spreadsheetId, sheetName, cell, value):
        coords = self.convertCellToRowCol(cell)
        sheetId = self.getSheetId(spreadsheetId, sheetName) # would be nice to cache this

        valueType = type(value)
        if (valueType is str):
            userValueType = "stringValue"
        elif (valueType is int or value is float):
            userValueType = "numberValue"
        else:
            raise Exception("Unknown type '" + str(valueType) + "' of value '" + str(value) + "'")

        updateRequest = {
            "updateCells": {
                "start": {
                    "sheetId": sheetId,
                    "rowIndex": coords["row"],
                    "columnIndex": coords["col"]
                    },
                "rows": [
                    {
                        "values": [
                            { "userEnteredValue": {userValueType: value}, }
                        ]
                    }
                ],
                "fields": "userEnteredValue"
                }
            }

        updateBody = {
            "requests" : [ updateRequest ]
        }

        # print updateBody
        resp = self.spreadsheets.batchUpdate(spreadsheetId=spreadsheetId, body=updateBody).execute()

    # given a sheet name, find the id
    def getSheetId(self, spreadsheetId, sheetName):
        resp = self.spreadsheets.get(spreadsheetId=spreadsheetId).execute()
        for sheet in resp["sheets"]:
            if sheet["properties"]["title"] == sheetName:
                return sheet["properties"]["sheetId"]

if __name__ == "__main__":
    cellmate = Cellmate()

    MAGIC_PREFIX = "#Cellmate Configuration Details"

    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--id", help="Spreadsheet ID")
    parser.add_argument("-n", "--name", help="Sheet Name")
    parser.add_argument("-c", "--cell", type=str, help="Target Cell")
    parser.add_argument("-f", "--file", type=str, help="File to Read/Write")
    parser.add_argument("-v", "--value", type=str, help="Value")

    parser.add_argument("-o", "--op", type=str, help="Target Cell", required=True)


    args = vars(parser.parse_args())
    # if anything is missing we won't get past this point

    ssId = args["id"]
    sheetName = args["name"]
    cell = args["cell"]
    filename = args["file"]
    value = args["value"]

    if args["op"] == "store":
        if (ssId == "" or sheetName == "" or cell == "" or filename == ""):
            print "Missing required param - need id/name/cell/file"
        else:
            val = cellmate.getValue(ssId, sheetName, cell)
            print "Storing value from " +  sheetName + "!:" + cell + " in '" + filename + "'"
            headerLine = "{}\t{}\t{}\t{}\t{}\n".format(MAGIC_PREFIX,ssId,sheetName,cell,"1")
            with open(filename, "w") as textFile:
                textFile.write(headerLine)
                textFile.write(val)
    elif args["op"] == "check":
        if (ssId == "" or sheetName == "" or cell == ""):
            print "Missing required param - need id/name/cell"
        else:
            val = cellmate.getValue(ssId, sheetName, cell)
            print val
    elif args["op"] == "update":
        if (ssId == "" or sheetName == "" or cell == "" or value is None):
            print "Missing required param - need id/name/cell/value"
        else:
            cellmate.setValue(ssId, sheetName, cell, value)
    elif args["op"] == "upload":
        if (filename == ""):
            print "Missing required param"
        else:
            with open(filename, "r") as textFile:
                contents = textFile.read()

                segments = contents.split("\n")

                destinationCounter = 0
                destinations = []
                while True:
                    probe = segments[destinationCounter]

                    chunks = probe.split("\t")
                    if (chunks[0] == MAGIC_PREFIX and len(chunks) >= 4):
                        if (len(chunks) == 4 or (len(chunks) >= 5 and int(chunks[4]) == 1)):
                            destinations.append(probe)

                        destinationCounter = destinationCounter + 1
                    else:
                        break

                if (len(destinations) > 0):
                    stuffToUpload = "\n".join(segments[destinationCounter:])

                    for d in destinations:
                        chunks = d.split("\t")

                        ssId = chunks[1]
                        sheetName = chunks[2]
                        cell = chunks[3]

                        try:
                            print "Uploading '" + filename + "' contents to " + sheetName + "!" + cell + " (" + ssId + ")..."
                            cellmate.setValue(ssId, sheetName, cell, stuffToUpload)
                        except:
                            print "Error encountered; check spreadsheet permissions and file metadata."

                else:
                    print "No destinations enabled in file"
    else:
        print "invalid op '" + args["op"] + "' - valid ops are 'store','check','upload','update'"
