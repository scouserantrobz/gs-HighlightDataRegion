# gs-HighlightDataRegion
Highlight the data region generated by an ARRAY formula in Google Scripts

### Introduction

Google Sheets has a number of formulas that result in an array of data. Examples include SPLIT() and the incredibly powerful QUERY(). It's often very helpful to see the boundaries of these formulas, so you can manage the layout of your sheet better, and ensure they do no overlap into other cells which already have values, causing the formulas to break.

Interestingly, Excel does this when the active cell contains an array formula that gives multiple results - [see screenshot here](https://support.content.office.net/en-us/media/dbd32b6c-e6d0-4289-bee1-a3158107ec1c.png "Screenshot of Array Formula in Excel") (full page [here](https://support.microsoft.com/en-us/office/create-an-array-formula-e43e12e0-afc6-4a12-bc7f-48361075954d "from Microsoft Documentation for 'Create An Array Formula") ).

This script replicates this Excel functionality, but the formatting (background colour and border settings) is permanent, not just displayed when the array formula is the active cell.

### Installation
There is only one code file, which you can copy and paste into a new gs file in your Google Spreadsheet. Save the file and authorise it's use as normal. 

If the script installs succesfully, when the Spreadsheet reload, you'll see an new menu item (Antro's Actions) with two menu options:
1. Format Array Data Region
2. Format Array Data Region & Log

You can change these text labels in the script to suit your preferences.

If you wish to use the second option, you will need to have a sheet called config, which is laid out exactly like this [template sheet here](https://docs.google.com/spreadsheets/d/1ig5UISJtEItlHNgKwLFSuap3hiG5XsgA8QiP4G_J0tY/edit?fbclid=IwAR3G-ZwCEtfXaPTSdJq5p7VrFOciurfWIETbCqJVwNDrnsq5nGBUzV-TXo8#gid=54678740 "Link to Spreadsheet with config template sheet"). Make a copy of the spreadsheet, clear rows 5 and 6, and then copy the empty template to the sheet you are working on (use the **_Copy to_** option from the sheet context menu).

### Use
Both menu options perform the same basic task - highlighting the data region created by the array formula in the active cell.

#### Format Array Data Region
The first option takes the settings of the active cell and applies it to the data region resulting from the array formula. Before applying the border settings, it clears all the current settings (including the active cell). 

#### Format Array Data Region & Log
The second option takes the settings from cell D2 on a sheet called config. This is a special sheet that can be found in this [example Spreadsheet](https://docs.google.com/spreadsheets/d/1ig5UISJtEItlHNgKwLFSuap3hiG5XsgA8QiP4G_J0tY/edit?usp=sharing "Link to Google Spreadsheet"). This sheet records the properties for each array formula processed by the script. Each formula is recorded in it's own row - with column A formmatted exactly the same as the data region. Columns B, C, and D store the sheet name, cell address of the array formula, and these two values combined. Column E keeps a record of the array formula itself, and the final two columns, F and G record the number of rows and columns in the data region.

If the script is run a second time on the same formula, it's row on the config is updated to reflect the new results.

---