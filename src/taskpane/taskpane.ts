/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

async function CreateTable(context) {
  try {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    
    // Define the range of cells you want to select
    const range = sheet.getUsedRange();
    range.load("address");
    
    // Create a table from the selected range
    const table = sheet.tables.add(range, true /* hasHeaders */);
    table.name = "CDL";
    // Update table style 
    //Possible values are: "TableStyleLight1" through "TableStyleLight21", "TableStyleMedium1" 
    //through "TableStyleMedium28", "TableStyleDark1" through "TableStyleDark11"
    table.style = "TableStyleLight10";
    table.load('tableStyle');
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
  
  const columnsEF = sheet.getRange("E:F");
  columnsEF.format.indentLevel = 1;
  const columnC = sheet.getRange("C:C");
  columnC.columnWidth = 8;
  const columnE = sheet.getRange("E:E");
  columnE.columnWidth = 18;
  const columnG = sheet.getRange("G:G");
  columnG.columnWidth = 25;

  const columnsHI = sheet.getRange("H:I");
  columnsHI.format.horizontalAlignment = "Center";
  columnsHI.format.verticalAlignment = "Bottom";
  columnsHI.columnWidth = 6;

  // const columnsKM = sheet.getRange("K:M");
  // columnsKM.format.indentLevel = 1;
  // sheet.activate();
  // sheet.getUsedRange().getOffsetRange(-1, 0).getOffsetRange(1, 0).activate();
  // const columnP = sheet.getRange("P:P");
  // columnP.columnWidth = 6.33;

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

async function FilterCells(context) {
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  let CDLTable = sheet.tables.getItem("CDL");

    // Queue a command to apply a filter on the Category column.
    let ignorableFilter = CDLTable.columns.getItem("Ignorable").filter;
    ignorableFilter.apply({
      filterOn: Excel.FilterOn.values,
      values: ["FALSE"]
    });
  console.log("Cells filtered.");
}

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      await CreateTable(context);
      await FormatCells(context);
      await FilterCells(context);

      
      await context.sync();
      console.log(`Processing done.`);
    });
  } catch (error) {
    console.error(error);
  }
}
