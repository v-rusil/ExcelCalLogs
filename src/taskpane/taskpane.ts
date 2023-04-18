/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { ColumnDefinition, EnumColumnHorizontalAlignment, EnumColumnVerticalAlignment } from "./columnDefinitions";
import { JsonConfigUtils } from "./jsonConfigUtils";

/* global console, document, Excel, Office */

var ctx:Excel.RequestContext

var tbl:Excel.Table;
var sheet:Excel.Worksheet;
var tblRange:Excel.Range;


//#region JSON properties of Callog
var RAVECDLLOG:string = `
  [  
    {    
      "columnName": "ModifiedDate",   
      "isMandatory": "true", 
      "horizontalAlignment": "Center",    
      "verticalAlignment": "Bottom",    
      "columnWidth": 180,    
      "indentLevel": 1,    
      "style": "Neutral",
      "numberFormat": "MM/dd/yyyy HH:mm:ss"  
    },  
    {    
      "columnName": "Age",  
      "isMandatory": "false",  
      "horizontalAlignment": "center",    
      "verticalAlignment": "middle",    
      "columnWidth": 80,    
      "indentLevel": 1,    
      "style": "italic",  
      "numberFormat": "MM/dd/yyyy HH:mm:ss"  
    }
  ]`

//#endregion


//#region Properties
var _totalTblRows: number = -1;
async function totalTblRows(): Promise<number>
{
  if (_totalTblRows <= 0) {
    tblRange.load(["rowCount"]);
    await ctx.sync();
    
    _totalTblRows = tblRange.rowCount;
      
  }
  return _totalTblRows;
}

// dropdown Type of CDL log
function typeCDL(): string {
  const selectElement = document.getElementById('typeCDL') as HTMLSelectElement;
  return selectElement.value;
}

// Organizer table styles and checkbox
var organizerTableStyle:string = "TableStyleLight13";
var attendeeTableStyle:string = "TableStyleLight10";
async function isOrganizer():Promise<boolean>
{
  const checkbox = document.getElementById("isOrganizer") as HTMLInputElement;
  return checkbox.checked;
}


async function warn1KRows():Promise<boolean>
{
  const checkbox = document.getElementById("warn1KRows") as HTMLInputElement;
  return checkbox.checked;
}

//#endregion

//#region Helper methods
async function showSpinner(show:boolean)
{
  var element = document.getElementById("spinner");
  if (show) {
    element.classList.remove("invisible");
  } else {
    element.classList.add("invisible");
  }
}

function AddMessage(message: string) {
  const p = document.getElementById("statusMessage");
  p.textContent = message;
}

enum enumTypeAnalysis
{
  Warning, 
  Action, 
  Danger, 
  Success
}
async function addAnalysisInfo(title:string, badge:number, message:string, smallfooter:string, typeanalysis:enumTypeAnalysis)
{
  // <a href="#" class="list-group-item list-group-item-action list-group-item-warning">
  //    <div class="d-flex w-100 justify-content-between">
  //      <h5 class="mb-1">Row limit</h5>
  //      <span class="badge badge-primary badge-pill">1002</span>
  //    </div>
  //    <p class="mb-1">If rows returned are close to 1K</p>
  //    <small>Get-CalendarDiagnosticObjects</small>
  // </a>
    const analysisDiv = document.getElementById("analysisInfo");
  
    // Create the <a> element with the appropriate class based on the enum value
    const aElement = document.createElement("a");
    switch (typeanalysis) {
      case enumTypeAnalysis.Warning:
        aElement.classList.add("list-group-item", "list-group-item-action", "list-group-item-warning");
        break;
      case enumTypeAnalysis.Action:
        aElement.classList.add("list-group-item", "list-group-item-action");
        break;
      case enumTypeAnalysis.Danger:
        aElement.classList.add("list-group-item", "list-group-item-action", "list-group-item-danger");
        break;
      case enumTypeAnalysis.Success:
        aElement.classList.add("list-group-item", "list-group-item-action", "list-group-item-success");
        break;
    }
  
    // Create the <div> element with the appropriate classes and contents
    const divElement = document.createElement("div");
    divElement.classList.add("d-flex", "w-100", "justify-content-between");
    const h5Element = document.createElement("h5");
    h5Element.classList.add("mb-1");
    h5Element.innerText = title;
    const spanElement = document.createElement("span");
    spanElement.classList.add("badge", "badge-primary", "badge-pill");
    if (badge == 0) {
      spanElement.classList.add("invisible");
    }
    spanElement.innerText = badge.toString();
    divElement.appendChild(h5Element);
    divElement.appendChild(spanElement);
  
    // Create the <p> element with the message
    const pElement = document.createElement("p");
    pElement.classList.add("mb-1");
    pElement.innerText = message;
  
    // Create the <small> element with the footer text
    const smallElement = document.createElement("small");
    smallElement.innerText = smallfooter;
  
    // Add the child elements to the <a> element
    aElement.appendChild(divElement);
    aElement.appendChild(pElement);
    aElement.appendChild(smallElement);
  
    // Add the <a> element to the analysis div
    analysisDiv.appendChild(aElement);
  
  }

  async function resetAnalysisInfo()
  {
    const analysisInfoDiv = document.getElementById('analysisInfo');
    analysisInfoDiv.innerHTML = '';
  }
//#endregion



//#region Init OfficeJS
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("createConfig").onclick = createConfig;
    resetAnalysisInfo();
  }
});

async function ClearAllTables() {
  try {
    await Excel.run(async (context) => {
      
      // Get all the tables in the worksheet
      const tables = sheet.tables;

      // Load the items property of the tables object
      tables.load("items");

      // Synchronize the document state by executing the queued commands
      await context.sync();

      // Loop through each table and remove its formatting
      tables.items.forEach((table) => {
        table.getRange().clear("Formats");
      });

      // Synchronize the document state by executing the queued commands
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
//#endregion

async function CreateTable(context) {
  try {

    // await ClearAllTables();

    const _isOrganizer = await isOrganizer();

    // Define the range of cells you want to select
    const range = sheet.getUsedRange();
    range.load("address");
    
    // Create a table from the selected range
    let tbl = sheet.tables.getItemOrNullObject("CDL");
    let tblRange = tbl.getRange();
    await context.sync();

    if (tbl.isNullObject) {
      tbl = sheet.tables.add(range, true /* hasHeaders */); 
      tbl.name = "CDL"; 
      tblRange=tbl.getRange();
      await context.sync();
    }

    tblRange.clear("Formats");
    await context.sync();

    if (_isOrganizer) {
      tbl.style = organizerTableStyle;  
    }
    else{
      tbl.style = attendeeTableStyle;  
    }
    
    tbl.load('tableStyle');
    tbl.columns.load();
    tblRange = tbl.getRange();
    await context.sync();
    
    await context.sync();

    


    // // Update table style 
    // //Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" 
    // //through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11"
    // tbl.style = "TableStyleLight10";
    // tbl.load('tableStyle');
    // await context.sync();
    
  } catch (error) {
    console.error(error);
    AddMessage(error);
  }
}


async function FormatCells(context) 
 {
  // // Get a reference to the active worksheet
  // const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Format cells
  const columnsA = sheet.getRange("A:A");
  columnsA.format.horizontalAlignment = "Center";
  columnsA.format.verticalAlignment = "Bottom";
  columnsA.format.columnWidth = 6;
  columnsA.format.autofitColumns();
  
  const columnsEF = sheet.getRange("E:F");
  columnsEF.format.autofitColumns();
  columnsEF.format.indentLevel = 1;

  const columnC = sheet.getRange("C:C");
  columnC.format.columnWidth = 8;
  const columnE = sheet.getRange("E:E");
  columnE.format.columnWidth = 18;
  const columnG = sheet.getRange("G:G");
  columnG.format.autofitColumns();
  columnG.format.columnWidth = 25;

  const columnsHI:Excel.Range = sheet.getRange("H:I");
  columnsHI.format.horizontalAlignment = "Center";
  columnsHI.format.verticalAlignment = "Bottom";
  columnsHI.format.columnWidth = 6;

  const columnsKM:Excel.Range = sheet.getRange("K:M");
  columnsKM.format.indentLevel = 1;
  //sheet.getUsedRange().getOffsetRange(-1, 0).getOffsetRange(1, 0).activate();
  const columnP:Excel.Range = sheet.getRange("P:P");
  columnP.format.columnWidth = 6.33;

  // const tableName = "MyTable";
  // const table = sheet.tables.getItem(tableName).range;
  // const filterRange = table.getOffsetRange(1, 15).getIntersection(table.getUsedRange());
  // filterRange.autoFilter(1, "TRUE");

  const columnQ = sheet.getRange("Q:Q");
  columnQ.format.columnWidth = 11.67;
  
  const columnN = sheet.getRange("N:N");
  columnN.format.horizontalAlignment = "Center";
  columnN.format.verticalAlignment = "Bottom";
  columnN.format.columnWidth = 10;
  
  // const filterRange2 = table.getOffsetRange(1, 24).getIntersection(table.getUsedRange());
  // filterRange2.autoFilter(1, "ForwardedAppointment");

  const columnY = sheet.getRange("Y:Y");
  columnY.style = "Neutral";
  columnY.format.columnWidth = 10.56;
  const columnW = sheet.getRange("W:W");
  columnW.format.indentLevel = 1;
  columnW.format.columnWidth = 11.22;

  await context.sync();
  // Done
  console.log("Cells formatted.");
}

async function FilterIgnorable(value:string) {
    tbl.columns.load();
    await ctx.sync();
    let ignorableFilter = tbl.columns.getItemOrNullObject("Ignorable").filter;
    await ctx.sync();
    ignorableFilter.apply({
      filterOn: Excel.FilterOn.values,
      values: [value]
    });
  await ctx.sync();
  console.log("Cells filtered. ");
}


async function FormatDateColumn( context, columnName:string){
  //const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Format cells
  const col:Excel.TableColumn = tbl.columns.getItemOrNullObject(columnName);
  const colRange:Excel.Range = col.getDataBodyRange();
  await context.sync();
  await context.sync();


  const criteria: Excel.ReplaceCriteria = {
    completeMatch: false, /* Use a complete match to skip cells that already say "okay". */
    matchCase: true /* Ignore case when comparing strings. */
  };
  colRange.replaceAll("Z", "", criteria);

  await context.sync();

  // Apply horizontal alignment as "center", vertical alignment as "bottom" and wrap text as "false"
  colRange.format.horizontalAlignment = "Center";
  colRange.format.verticalAlignment = "Bottom";
  colRange.format.wrapText = false;
  // colRange.format.columnWidth=19;
  
  const rangeWidth:Excel.Range = colRange.getEntireColumn();
  rangeWidth.load("format"); await context.sync();
  rangeWidth.format.autofitColumns();
 
  await context.sync();

}

async function HighlightIgnorable(){

  const colRange:Excel.Range = tbl.columns.getItemOrNullObject("Ignorable").getRange();

  const conditionalFormat = colRange.conditionalFormats.add(
    Excel.ConditionalFormatType.containsText
  );

  // Color the font of every cell containing "Delayed".
  conditionalFormat.textComparison.format.font.color = "blue";
  conditionalFormat.textComparison.format.fill.color="#ADD8E6";
  // conditionalFormat.textComparison.style = "20% - Accent5";
  conditionalFormat.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: "TRUE"
  };
  await ctx.sync();

  return await CountFilterOccurrences(conditionalFormat.getRange());


  // await FilterIgnorable("TRUE").then(async ()=>{
  //     var filter = tbl.columns["Ignorable"].filter;
      
  //     const visibleRange:Excel.RangeView =tbl.getDataBodyRange().getVisibleView();
  //     var vr:Excel.Range = visibleRange.getRange();
  //     vr.load("address"); 
  //     await ctx.sync();
  //     vr.format.font.color = "blue";
  //     vr.format.fill.tintAndShade = 0.399975585192419;
  //     vr.style = "20% - Accent5";
  // });

  // await ctx.sync();
}

async function HighlightApptSequence(context){
    const col = tbl.columns.getItemOrNullObject("ApptSequence");
    
    const range = col.getDataBodyRange();
    await ctx.sync();
    const conditionalFormat = range.conditionalFormats
        .add(Excel.ConditionalFormatType.colorScale);
    const criteria = {
        minimum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "white" },
        // midpoint: { formula: "50", type: Excel.ConditionalFormatColorCriterionType.percent, color: "yellow" },
        maximum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "green" }
    };
    conditionalFormat.colorScale.criteria = criteria;

    await context.sync();
}

//format columns SentRepresentingEmailAddress	ResponsibleUserName	SenderEmailAddress

async function FormatRawFrom(context){

  const colRepresentingEmailAddress:Excel.Range =
        tbl.columns.getItemOrNullObject("SentRepresentingEmailAddress").getRange();
  const colResponsibleUserName:Excel.Range =
        tbl.columns.getItemOrNullObject("ResponsibleUserName").getRange();
  const colSenderEmailAddress:Excel.Range =
        tbl.columns.getItemOrNullObject("SenderEmailAddress").getRange();


colRepresentingEmailAddress.format.horizontalAlignment = "Center";
colRepresentingEmailAddress.format.verticalAlignment = "Bottom";
colRepresentingEmailAddress.format.indentLevel = 1;
colRepresentingEmailAddress.format.autofitColumns();
await context.sync();

colResponsibleUserName.format.horizontalAlignment = "Center";
colResponsibleUserName.format.verticalAlignment = "Bottom";
colResponsibleUserName.format.indentLevel = 1;
colResponsibleUserName.format.autofitColumns();
await context.sync();

colSenderEmailAddress.format.horizontalAlignment = "Center";
colSenderEmailAddress.format.verticalAlignment = "Bottom";
colSenderEmailAddress.format.autofitColumns();
colSenderEmailAddress.format.indentLevel = 1;
await context.sync();

}

async function CountFilterOccurrences( filterRange:Excel.Range ):Promise<number>
{
  await ctx.sync();
  // Get the range of cells affected by the conditional format
  const affectedRange = await filterRange.getIntersectionOrNullObject(tbl.getRange());
  affectedRange.load(["rowCount"]); await ctx.sync();
  // Get the number of rows affected by the conditional format
  const rowCount = affectedRange ? affectedRange.rowCount : 0;
  return rowCount;
}

async function HighlightCRA(context):Promise<number>
{

    const colRange:Excel.Range = tbl.columns.getItemOrNullObject("Client").getRange();

    const conditionalFormat = colRange.conditionalFormats.add(
      Excel.ConditionalFormatType.containsText
    );

    // Color the font of every cell containing "Delayed".
    conditionalFormat.textComparison.format.font.color = "red";
    conditionalFormat.textComparison.rule = {
      operator: Excel.ConditionalTextOperator.contains,
      text: "CRA:CalendarRepairAssistant"
    };
    await context.sync();

    return await CountFilterOccurrences(conditionalFormat.getRange());
  
}

async function HighLightCreates(context):Promise<number>
{
  const colRange:Excel.Range = tbl.columns.getItemOrNullObject("Trigger").getRange();

  const conditionalFormat = colRange.conditionalFormats.add(
    Excel.ConditionalFormatType.containsText
  );

  // Color the font of every cell containing "Delayed".
  conditionalFormat.textComparison.format.fill.color = "Green";
  conditionalFormat.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: "Create"
  };

  await context.sync();

  return await CountFilterOccurrences(conditionalFormat.getRange());

}


async function validateCDLStructure(jsonString: string): Promise<boolean> {
  try {
    var jsonArray = JSON.parse(jsonString);

    if (!Array.isArray(jsonArray)) {
      throw new Error("Input JSON is not an array.");
    }

    tbl.columns.load();
    await ctx.sync();

    for (const element of jsonArray) 
    {
      var tblCol:Excel.TableColumn;
      var tblColRange:Excel.Range;
      if (element.columnName !== undefined && element.columnName !== "") {
        tblCol = tbl.columns.getItemOrNullObject(element.columnName);
        tblColRange = tblCol.getRange();
        await ctx.sync();

        if (tblCol.isNullObject) 
        {
          if (element.isMandatory !== undefined && element.isMandatory !== "" && element.isMandatory == "false") {
            console.log(`isMandatory: ${element.isMandatory}`);
            continue;
          }
          console.log(`Column Name does not exist: ${element.columnName}`);
          addAnalysisInfo("columnName",0,`Column Name does not exist: ${element.columnName}`, "ValidateJSONStruct", enumTypeAnalysis.Danger);
          return false;
        }
        else
        {
          console.log(`Column Name: ${element.columnName}`);

          if (element.horizontalAlignment !== undefined && element.horizontalAlignment !== "") {
            console.log(`Horizontal Alignment: ${element.horizontalAlignment}`);
            tblColRange.format.horizontalAlignment = element.horizontalAlignment;await ctx.sync();
            await ctx.sync();
          }
    
          if (element.verticalAlignment !== undefined && element.verticalAlignment !== "") {
            console.log(`Vertical Alignment: ${element.verticalAlignment}`);
            tblColRange.format.verticalAlignment = element.verticalAlignment;
            await ctx.sync();
          }
    
          if (element.columnWidth !== undefined && element.columnWidth !== null) {
            console.log(`Column Width: ${element.columnWidth}`);
            tblColRange.format.columnWidth = element.columnWidth;
            await ctx.sync();
          }
    
          if (element.identLevel !== undefined && element.identLevel !== null) {
            console.log(`Ident Level: ${element.indentLevel}`);
            tblColRange.format.indentLevel = element.indentLevel;
            await ctx.sync();
          }
    
          if (element.style !== undefined && element.style !== "") {
            console.log(`Style: ${element.style}`);
            tblColRange.style = element.style;
            await ctx.sync();
          }

          if (element.numberFormat !== undefined && element.numberFormat !== "") {
            console.log(`Style: ${element.numberFormat}`);
            tblColRange.numberFormat = element.numberFormat;
            await ctx.sync();
          }

          await ctx.sync();
              
        }
        
      }
    }

    return true;

  } catch (error) {
    addAnalysisInfo("columnName",0,`Error traversing JSON array: ${error}`, "ValidateJSONStruct", enumTypeAnalysis.Danger);
    console.error(`Error traversing JSON array: ${error}`);
    return false;
  }
}




async function createColumnDefinitionsFromTable(): Promise<JsonConfigUtils> 
{
  var colDefinitions:JsonConfigUtils = new JsonConfigUtils();

  // Get the table object from the global variable 'tbl'
  // Load all columns and their format properties
  const columns = tbl.columns.load([
    'name',
    'values/format',
    'values/horizontalAlignment',
    'values/verticalAlignment',
    'values/columnWidth',
    'values/indentLevel',
    'values/style',
    'values/numberFormat',
    'values/autosizeColumn'
  ]);

  // Synchronize with the document
  await ctx.sync();

   // Iterate through the columns
   for (const column of columns.items) 
   {
      var r: Excel.Range = column.getRange();
      r.load(['style', 'numberFormat'])
      r.format.load(['format','horizontalAlignment', 'verticalAlignment', 'columnWidth', 'indentLevel', 'style', 'numberFormat']);
      await ctx.sync();
       // Access the loaded properties
       const name = column.name;
       const format = r.format;
       const horizontalAlignment = r.format.horizontalAlignment;
       const verticalAlignment = r.format.verticalAlignment;
       const columnWidth = r.format.columnWidth;
       const indentLevel = r.format.indentLevel;
       const style = r.style;
       const numberFormat = r.numberFormat[0].toString(); //numberFormat is an array of all the cells format, we'll check just the first row
       const autosizeColumn = false;
       
       // Create a column definition object and add it to your array
       const colDef: ColumnDefinition = {
         columnName: name,
         isMandatory: true, // Set this to whatever your default is
         horizontalAlignment:horizontalAlignment,
         verticalAlignment:verticalAlignment,
         columnWidth:columnWidth,
         indentLevel:indentLevel,
         style:style,
         numberFormat:numberFormat,
         visible: true, // Set this to whatever your default is
         autosizeColumn:false
       };
       colDefinitions.addColumn(colDef);
   }

  return colDefinitions;
}






async function PerformAnalysis(context) {
  await resetAnalysisInfo();
  await CheckNumberOfRows();
  await context.sync();
}

async function CheckNumberOfRows() {
  
  if (!CheckNumberOfRows) return; 

  const rowCount = await totalTblRows();
  
  if (rowCount >= 950) {
    addAnalysisInfo("Row count number", rowCount, "Number of rows is very close(or above) the Diag Limit of 1000Rows", "CheckNumberOfRows", enumTypeAnalysis.Warning);
    AddMessage("Number of rows is very close to the Diag Limit of 1000Rows returned($tblRange.rowCount)");
  }   
  await ctx.sync();
}


async function freezeColumns(columnName: string)
{
    const column = tbl.columns.getItem(columnName);
    sheet.freezePanes.freezeRows(1);
    sheet.freezePanes.freezeColumns(3);
    await ctx.sync();
}








async function getJsonData(): Promise<any> {
  const response = await fetch("./RaveCDLconfig.json");
  const jsonData = await response.json();
  return jsonData;
}











//#region Config section
export async function createConfig()
{
    var err = await getJsonData();
    
      var colDefs:JsonConfigUtils =  await createColumnDefinitionsFromTable();
      document.getElementById("jsonConfig").textContent = colDefs.convertColumnDefinitionsToJson();
}
//#endregion

//Main Function
export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      
      await showSpinner(true);
      await resetAnalysisInfo();
      await AddMessage("Starting Processing");
      

      ctx = context;
      sheet = context.workbook.worksheets.getActiveWorksheet();
      tbl = sheet.tables.getItemOrNullObject("CDL");
      tblRange = tbl.getRange();

      sheet.getRange().clear("Formats");
      sheet.getRange().conditionalFormats.clearAll();
      tbl.clearFilters();
      sheet.freezePanes.unfreeze();
      
  

      await context.sync() ;


      await  CreateTable(context).then(()=>{AddMessage("Create Table Done")});
      
      var isTableValid:boolean = await validateCDLStructure(RAVECDLLOG);
      if (!isTableValid) {
        addAnalysisInfo("CDL Invalid", 0, "CDL Structure is invalid (check previous exceptions)", "CDLInvalid", enumTypeAnalysis.Danger);
        showSpinner(false);
        return;
      }

      //format section
      await freezeColumns("Ignorable");
      await FormatCells(context).then(()=>{AddMessage("Format cells Done")});
      await FormatRawFrom(context).then(()=>{AddMessage("Format raw from Done")});
      await FormatDateColumn(context, "ModifiedDate").then(()=>{AddMessage("Format ModifiedDate Done")}); //ModifiedDate
      await FormatDateColumn(context, "StartTime").then(()=>{AddMessage("Create StartTime Done")}); //StartTime
      await FormatDateColumn(context, "EndTime").then(()=>{AddMessage("Create End Done")}); //EndTime

      //highlight section
      await HighlightIgnorable().then(()=>{AddMessage("Highlight Ignorable Done")});
      await HighlightApptSequence(context).then(()=>{AddMessage("Highlight  Done")});
      await HighlightCRA(context).then(()=>{AddMessage("Highlight CRA Done")});
      
      //await addAnalysisInfo("CRA Found", rowCount, "CRA Events were found, meaning calendar state was not 100%","HighlightCRA",enumTypeAnalysis.Warning);
      await HighLightCreates(context).then(()=>{AddMessage("Highlight Create Done")});

      //Filters section
      await FilterIgnorable("FALSE").then(()=>{AddMessage("Filter Ignorable Done")});

      await context.sync();
      
      await PerformAnalysis(context).then(()=>{AddMessage("Perform Analysis Done")});
      console.log(`Processing done.`);
      
      AddMessage("Done!");
      showSpinner(false);
      
      const urlCDLVideo:string = "https://msit.microsoftstream.com/video/4221a4ff-0400-9fb2-4805-f1eb0f28f09b";
      addAnalysisInfo("Success", 0,`Process executed successfully, check the video on CDL analysis ${urlCDLVideo} `, "success", enumTypeAnalysis.Success)
    });
  } 
  catch (error) 
  {
    showSpinner(false);
    console.error(error);
    AddMessage(error);
  }
  
  
}


