/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var tbl:Excel.Table;
var sheet:Excel.Worksheet;
var tblRange:Excel.Range;


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});



async function CreateTable(context) {
  try {
    
    // Define the range of cells you want to select
    const range = sheet.getUsedRange();
    range.load("address");
    
    // Create a table from the selected range
    tbl = sheet.tables.getItemOrNullObject("CDL");
    tblRange = tbl.getRange();
    await context.sync();
    if (tbl.isNullObject) {
      let tbl = sheet.tables.add(range, true /* hasHeaders */); 
      tbl.name = "CDL"; 
      tbl.style = "TableStyleLight10";
      tbl.load('tableStyle');
      tblRange = tbl.getRange();
      await context.sync();
    }
    
    
    // Update table style 
    //Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" 
    //through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11"
    tbl.style = "TableStyleLight10";
    tbl.load('tableStyle');
    await context.sync();
    
  } catch (error) {
    console.error(error);
  }
}

async function FormatCells(context) 
 {
  // Get a reference to the active worksheet
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Format cells
  const columnsA = sheet.getRange("A:A");
  columnsA.format.horizontalAlignment = "Center";
  columnsA.format.verticalAlignment = "Bottom";
  columnsA.columnWidth = 6;
  columnsA.format.autofitColumns();
  
  const columnsEF = sheet.getRange("E:F");
  columnsEF.format.autofitColumns();
  columnsEF.format.indentLevel = 1;

  const columnC = sheet.getRange("C:C");
  columnC.columnWidth = 8;
  const columnE = sheet.getRange("E:E");
  columnE.columnWidth = 18;
  const columnG = sheet.getRange("G:G");
  columnG.format.autofitColumns();
  columnG.columnWidth = 25;

  const columnsHI = sheet.getRange("H:I");
  columnsHI.format.horizontalAlignment = "Center";
  columnsHI.format.verticalAlignment = "Bottom";
  columnsHI.columnWidth = 6;

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
  columnQ.columnWidth = 11.67;
  
  const columnN = sheet.getRange("N:N");
  columnN.format.horizontalAlignment = "Center";
  columnN.format.verticalAlignment = "Bottom";
  columnN.columnWidth = 10;
  
  // const filterRange2 = table.getOffsetRange(1, 24).getIntersection(table.getUsedRange());
  // filterRange2.autoFilter(1, "ForwardedAppointment");

  const columnY = sheet.getRange("Y:Y");
  columnY.style = "Neutral";
  columnY.columnWidth = 10.56;
  const columnW = sheet.getRange("W:W");
  columnW.format.indentLevel = 1;
  columnW.columnWidth = 11.22;

  await context.sync();
  // Done
  console.log("Cells formatted.");
}

async function FilterIgnorable(context) {
  // let sheet = context.workbook.worksheets.getActiveWorksheet();
  // let CDLTable = sheet.tables.getItem("CDL");

    // Queue a command to apply a filter on the Category column.
    let ignorableFilter = tbl.columns.getItemOrNullObject("Ignorable").filter;
    ignorableFilter.apply({
      filterOn: Excel.FilterOn.values,
      values: ["FALSE"]
    });
  await context.sync();
  console.log("Cells filtered.");
}


async function FormatDateColumn( context, columnName:string){
  //const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Format cells
  const col:Excel.TableColumn = tbl.columns.getItemOrNullObject(columnName);
  const colRange:Excel.Range = col.getDataBodyRange();
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

async function HighlighIgnorable(context){

  let ignorableFilter = tbl.columns.getItemOrNullObject("Ignorable").filter;
  ignorableFilter.clear();
  await context.sync();
  const isFilterNull:boolean = ignorableFilter.isNullObject;
  ignorableFilter.apply({
    filterOn: Excel.FilterOn.values,
    values: ["TRUE"]
  });

  await context.sync();

  // Highlight the filtered data
  
  const range:Excel.Range = tbl.getDataBodyRange();
  tbl.load("address");await context.sync();
  range.format.font.color = "blue";
  // range.format.fill.tintAndShade = 0.399975585192419;
  // column.format.font.style = "20% - Accent5";

  // Clear the filter
  ignorableFilter.clear();

  await context.sync();
}

async function HighlightApptSequence(context){
    const col = tbl.columns.getItemOrNullObject("ApptSequence");
    
    const range = col.getDataBodyRange();
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


async function HighlightCRA(context){

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
}

async function HighLightCreates(context){
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
}


//#region Status Update (HTML)
function addStatus(action) {
  const ul = document.createElement("ul");
  ul.classList.add("ms-List", "ms-welcome__status");

  const li = document.createElement("li");
  li.classList.add("ms-ListItem", "ms-font-m");
  li.textContent = action;

  ul.appendChild(li);

  const main = document.getElementById("app-body");
  main.appendChild(ul);
}
//#endregion


//Main Function
export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      addStatus("Starting Processing");
      sheet = context.workbook.worksheets.getActiveWorksheet();
      
      await context.sync();
      
      await CreateTable(context);
      await FormatCells(context);
      await HighlighIgnorable(context);
      await HighlightApptSequence(context);
      await HighlightCRA(context);
      await HighLightCreates(context);
      await FormatRawFrom(context);

      await FormatDateColumn(context, "ModifiedDate"); //ModifiedDate
      await FormatDateColumn(context, "StartTime"); //StartTime
      await FormatDateColumn(context, "EndTime"); //EndTime

      await FilterIgnorable(context);

      await context.sync();
      console.log(`Processing done.`);

      await PerformAnalysis(context);
      console.log(`Processing done.`);
      
      addStatus("Done!");
    });
  } catch (error) {
    console.error(error);
  }
}
async function PerformAnalysis(context) {
  await CheckNumberOfRows(context);
  await context.sync();
}

async function CheckNumberOfRows(context: any) {
  tblRange.load(["rowCount"]);
  await context.sync();
  
  if (tblRange.rowCount >= 950) {
    AddMessage("Number of rows is very close to the Diag Limit of 1000Rows returned($tblRange.rowCount)");
  }   
  await context.sync();
}

function AddMessage(message: string) {

  const ul = document.getElementById("message");
  
  const li = document.createElement("li");
  li.classList.add("ms-ListItem", "ms-font-m");
  li.textContent = message;

  ul.appendChild(li);

}

