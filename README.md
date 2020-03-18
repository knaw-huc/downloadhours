Step 1: Go [https://developers.google.com/sheets/api/quickstart/nodejs](here) and click the big blue button to create credentials

Step 2: run the container with `docker run --rm -it -v "$PWD/credentials.json":/app/credentials.json -v "$PWD/state":/app/state knawhuc/downloadhours <last weeknumber to import> <google sheetids> > my_export.csv`

 - weeknumber is a 6 digit string with the year and week so week 10 of 2020 becomes 202010
 - google sheetids is the hash in the sheets uri. https://docs.google.com/spreadsheets/d/<here is the sheet id>/edit#gid=447837688
 - You can specify more then one sheetid. The result is printed as one big csv.

Exported weeks are marked as read-only before exporting. Weeks marked as read-only are not exported again.
If an exception occurred during exporting, or if you did not pipe stdout to a file, you have to remove the read-only ranges in google sheets by going to "protected sheets & ranges" and removing them one by one
