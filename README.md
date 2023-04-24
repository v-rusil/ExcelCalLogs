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

```
Above steps will start web server and excel, with a toolbar button labeled "Show Taskpane"
Code . will start Visual Studio Code to edit / browse add-in code  
```
/src/taskpane/taskpane.css - StyleSheet for taskpane
/src/taskpane/taskpane.html - HTML UI
/src/taskpane/taskpane.ts - TypeScript (This will be the main file)
```
Clicking Run in the taskpane will format the CalLogs (Make sure cal logs are pasted to the A1 cell, or else script will not work.
