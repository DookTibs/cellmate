# import argparse

from apiclient.discovery import build

"""
simple class for reading/write with a single Google Sheets cell programmatically.

would be great to add a flag to force a type of the update...for isntance, so one could
update a formula externally, which is valueType "formulaValue" not "stringValue"...or
so one could read in cmdline args and know if '99' is supposed to be a string or an int
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

    ssId = "1HCfCUGqoKBvf0wsrbJezyqwYyjnPYe-Vd9Xtbx_VZWw"
    sheetName = "Sheet1"
    cell = "A1"

    # val = cellmate.getValue(ssId, sheetName, "C1")
    # cellmate.setValue(ssId, sheetName, "C1", "hello there")

    """
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--id", help="Spreadsheet ID", required=True)
    parser.add_argument("-n", "--name", help="Sheet Name", required=True)
    parser.add_argument("-r", "--row", type=int, help="Target Row", required=True)
    parser.add_argument("-c", "--col", type=int, help="Target Column", required=True)
    parser.add_argument("-v", "--value", help="Update Value", required=True)

    args = vars(parser.parse_args())
    # if anything is missing we won't get past this point

    stuffer = Stuffer()
    spreadsheetId = args["id"]
    sheetId = stuffer.getSheetId(spreadsheetId, args["name"])
    stuffer.stuffValue(spreadsheetId, sheetId, args["row"], args["col"], args["value"])
    """
