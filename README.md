# DEPRECATED
## code not in use any more (by HuC)


Step 1: Go [here](https://developers.google.com/sheets/api/quickstart/nodejs) and click the big blue button to create credentials

Step 2: run the container with `docker run --rm -it -v "$PWD/credentials.json":/app/credentials.json -v "$PWD/state":/app/state knawhuc/downloadhours -f <last weeknumber to import> -s <google sheetid1> -s <google sheetid2> -d > my_export.csv`

If you got a proper csv back, remove the -d to mark the exported weeks readonly.

 - weeknumber is a 6 digit string with the year and week so week 10 of 2020 becomes 202010
 - google sheetids is the hash in the sheets uri. `https://docs.google.com/spreadsheets/d/<here is the sheet id>/edit#gid=447837688`
 - You can specify more then one sheetid (repeat the -s option). The result is printed as one big csv.

Exported weeks are marked as read-only before exporting. Weeks marked as read-only are not exported again.
If an exception occurred during exporting, or if you did not pipe stdout to a file, you have to remove the read-only ranges in google sheets by going to "protected sheets & ranges" and removing them one by one.

Therefore it's smart to do a dry-run first with the -d option
