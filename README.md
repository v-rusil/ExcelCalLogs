# ExcelCalLogs
This is a POC project to format CalLogs, on porting VBA script developed by Shane Ferrel (shanefe) and converting into an Excel web add-in.

To use this project follow below steps:

Clone this repo to your local dev box
"cd" to the root folder of the project 

Open cmd Prompt
```
npm install
code .
npm run start
```
If you run into issues with JQuery, you may not have the Jquery package installed, then run below at command line
This should install JQuery dependency and solve the issue.
```
npm i --save-dev @types/jquery
```


Above steps will start web server and excel, with a toolbar button labeled "Show Taskpane"
Code . will start Visual Studio Code to edit / browse add-in code  
```
/src/taskpane/taskpane.css - StyleSheet for taskpane
/src/taskpane/taskpane.html - HTML UI
/src/taskpane/taskpane.ts - TypeScript (This will be the main file)
```
Clicking Run in the taskpane will format the CalLogs (Make sure cal logs are pasted to the A1 cell, or else script will not work.

Config JSON File structure
First stage of development is concluded. The goal on this first release is to create the formatting engine as configurable as possible by using JSON config file.
The Config file is composed of 3 main sections (columns, conditionalFormats and filterDefinitions)

```json
{
    "columns": [{
            "autosizeColumn": false,
            "columnName": "LogRow",
            "columnWidth": 65.625,
            "horizontalAlignment": "Center",
            "indentLevel": 0,
            "isMandatory": true,
            "numberFormat": "General",
            "replaceWith": "",
            "searchFor": "",
            "style": "Normal",
            "verticalAlignment": "Bottom",
            "visible": true
        }
       
    ],
    "conditionalFormats": [{
            "CellValueFormula1": null,
            "CellValueFormula2": null,
            "CellValueOperator": null,
            "ColorScaleColorMaximum": null,
            "ColorScaleColorMinimum": null,
            "ColumnName": "Client",
            "ContainsTextOperator": "Contains",
            "ContainsTextSearch": "CRA:CalendarRepairAssistant",
            "CountOccurrences": false,
            "CustomFormula": null,
            "FillColor": "#FF0000",
            "FontColor": "Red",
            "FriendlyName": "Client Contains Text CRA:CalendarRepairAssistant",
            "Style": "Normal",
            "Type": "ConstainsText",
            "WarnIfOccurrencesGTZero": false
        },
        {
            "CellValueFormula1": null,
            "CellValueFormula2": null,
            "CellValueOperator": null,
            "ColorScaleColorMaximum": "#008000",
            "ColorScaleColorMinimum": "#FFFFFF",
            "ColumnName": "ApptSequence",
            "ContainsTextOperator": null,
            "ContainsTextSearch": null,
            "CountOccurrences": false,
            "CustomFormula": null,
            "FillColor": null,
            "FontColor": null,
            "FriendlyName": "ApptSequence Color Scale #008000 Down to #FFFFFF",
            "Style": "Normal",
            "Type": "ColorScale",
            "WarnIfOccurrencesGTZero": false
        }
    ],
    "filterDefinitions": [{
        "ColumnName": "Ignorable",
        "FilterActiveByUIOnly": false,
        "FilterKey": "Ignorable697408",
        "FilterValue": "FALSE",
        "FriendlyName": "Ignorable filter =FALSE"
    }]
}
```

This file configures per column formatting, indentation, values replacement and base style. In above example, if a LogRow column exists it will be formatted as per propery definitions (they are self explanatory, but will cover each property in more  detail in the future).

The "conditionalFormats" section, is used to conditionally highlight portions of column data that needs to visually call attention, either by checkin patterns in the case of color scales, either by highlighting certain cells, if a particular condition is met, such as "Ignorable Column has a property value of TRUE".

The "filterDefinitions" area, will configure a final stage of the formatting process, allowing certain data to be filtered (such as Ignorable rows), and presented in the end of a process in a more condensed stage. (Note: FilterByUIOnly is not yet implemented, and all filter conditions with this property set to True will be ignored by the processor)

This file can be Generated by visually configuring styles. If you want to submit a different config style for any other tabular type of data, this can easilly be achieved by generating a sample config file and testing formatting data against the modified JSON data in the editor.

