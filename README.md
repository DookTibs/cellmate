Note to self - I developed this using virtualenv:
"workon google_uploader"

Go to https://console.developers.google.com/ to set up a Service Account. Save the private key
somewhere like "~/serviceAccount.json"

Grab the email address from that file and share the sheet you want to work on with that address.
	
Then be sure to do this or none of this will work:
    declare -x GOOGLE_APPLICATION_CREDENTIALS=~/serviceAccount.json

Can be run from commandline in one of four modes:

* python cellmate.py -i <spreadsheetId> -n 'SheetName' -c 'A1' -f foo.txt -o store
This will pull the value from a given Google Sheets cell and store it in foo.txt, along
with a special header. This header contains the spreadsheetId/sheetName/cell. The header can
also contain a "1" or "0" at the end to enable/disable upload to this destination. It can also include
a tab-separated comment field at the end.

* python cellmate.py -f foo.txt -o upload
This will check foo.txt to see if it has the special header(s) we expect. If so, it will
figure out the right Google Sheets cell on the appropriate sheet(s) to talk to and attempt to upload the rest of the
file contents into that cell.

* python cellmate.py -i <spreadsheetId> -n 'SheetName' -c 'A1' -o check
Print the value of the cell.

* python cellmate.py -i <spreadsheetId> -n 'SheetName' -c 'A1' -v 'some value' -o update
Store 'some value' in the cell.

I use store/upload for editing AwesomeTable (https://awesome-table.com/) configuration code. AwesomeTable
has support for storing templates/CSS/JavaScript in Google Sheets. This is convenient, but it's
pretty unpleasant editing code directly in Google Sheets! Using Cellmate, I can store the various
configuration details in their own files, keep them in source control, and then edit in Vim/whatever
and get syntax highlighting and real editor tools, and simply push my changes back up to the correct cell
via a macro.
