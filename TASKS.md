### 1. Setup the app script in google sheets

- Go to Google Sheets: https://docs.google.com/spreadsheets/u/0/ 
- Create a Blank Spreadsheet
- Click on 'Extensions' -> 'Apps Script'
- You can copy paste the .js code from this repo
- Rename the file to Rolodex and begin working

### 2. Test both scripts

- Copy the code from sheetSetup.js file and and paste it in the .gs file
- Save the code 'Ctrl + S' or there is a save 'Save project to drive' button at the top
- Once saved, go back to the blank sheet that you created and reload it.
- You should be able to 'Setup' as another option in the same row as 'Extensions, Help'
- Click on 'Setup' -> 'Setup Sheet'
- This should prompt you to an Authorization window asking advance permissions (since the code is not verified yet) to run the code. Simply click ok and follow the instructions on screen choose a google account -> 'Google hasn't verified this app' screen click 'Advanced' - 'Go to Extension Test(unsafe)' -> 'Continue' -> 'Select All' and click 'Allow'. You should be good to go.
- You can try running it again after the all permissions have been granted.

**Note** - In order to setup the script for 'rolodex.js' the steps are the exact same, just copy and paste the code from 'rolodex.js file into the .gs file.

### 3. Try to combine the code from both scripts into one

### 4. Birthday/Anniversary/Last Meeting/Last Interaction problem

Currently stored in the yyyy/mm/dd format. How can we simplify this from a data entry perspective for the user?

### 5. Phone Number entry problem:

Currently must be entered as '+44 7000 111111, so that the google sheet does not think it's a formula. Is there a way to avoid adding that ' symbol before for this to still work? 