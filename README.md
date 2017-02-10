Uses virtualenv:
"workon google_uploader"

Go to https://console.developers.google.com/ to set up a Service Account. Save the private key
somewhere like "~/serviceAccount.json"

Grab the email address from that file and share the sheet you want to work on with that address.
	
Then be sure to do this or none of this will work:
    declare -x GOOGLE_APPLICATION_CREDENTIALS=~/serviceAccount.json
