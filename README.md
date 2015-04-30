# callings-tracker

#### Installing
TBD

#### Using 
##### *Callings* menu items
* *Sort callings*: Sorts each callings sheet (Pending, Current, Archive)
* *Organize callings*: Sorts each callings sheet, then moves each row to the right sheet depending on its status
* *Update calling status*: Forces the statuses of callings on the Pending and Current sheets to be updated. Use this if status didn't update automatically on one or more callings.
* *Print pending callings*: Generates a Google Doc containing the contents of the Pending callings sheet
* *Download members list*: Under construction
* *About*: Shows information about callings-tracker

#### Updating an existing spreadsheet
1. Go [here](https://raw.githubusercontent.com/elesel/callings-tracker/master/Code.gs) and copy all of the text to your clipboard
2. Open the spreadsheet from [Google Drive](https://drive.google.com/drive/my-drive)
3. Click *Tools* -> *Script editor...* from the menu bar
4. Click in the area underneath the *Code.gs* tab
5. Press Ctrl-A to select all of the code
6. Press Ctrl-V to replace the code with what you copied to the clipboard in step #1
7. Press Ctrl-S to save *Cods.gs*
8. Click *Resources* -> *Current project's triggers*
9. Make sure the rows in the table match the following (use the *Add a new trigger* link to add a new row):
  * *onChange* | *From spreadsheet* | *On change*
  * *onEdit* | *From spreadsheet* | *On edit*
10. Click the *Save* button
