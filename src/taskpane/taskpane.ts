/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { ConditionalFormat, enumCellValueOperator, enumConditionalFormatTextOperator, enumConditionalFormatType } from './ConditionalFormats';
import { FilterDefinition } from './FilterDefinitions';
import { ColumnDefinition, EnumColumnHorizontalAlignment, EnumColumnVerticalAlignment } from "./columnDefinitions";
import { JsonConfigUtils } from './jsonConfigUtils';
import { myConsole } from "./myConsole";


/* global console, document, Excel, Office */
module.exports = ctx;


var ctx:Excel.RequestContext

var tbl:Excel.Table;
var sheet:Excel.Worksheet;
var tblRange:Excel.Range;
var jsonConfigUtils:JsonConfigUtils;


//#region Constants
// Organizer table styles and checkbox
const organizerTableStyle:string = "TableStyleLight10";
const attendeeTableStyle:string = "TableStyleLight13";

const organizerTabColor:string = "#FFA500"; //orange
const attendeeTabColor:string = "#ADD8E6";  //light blue
//#endregion


//#region JSON properties of Callog
var jsonLog:string = `
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

async function hideLessRelevants():Promise<boolean>
{
  const checkbox = document.getElementById("hideLessRelevants") as HTMLInputElement;
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
  myConsole.log(message);
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
  
    AddMessage(message);
  }

  async function resetAnalysisInfo()
  {
    const analysisInfoDiv = document.getElementById('analysisInfo');
    analysisInfoDiv.innerHTML = '';
  }

  async function freezeColumns(columnName: string)
{
    const column = tbl.columns.getItem(columnName);
    sheet.freezePanes.freezeRows(1);
    sheet.freezePanes.freezeColumns(3);
    await ctx.sync();
}
//#endregion

//#region Init OfficeJS
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    jsonConfigUtils = new JsonConfigUtils();

    document.getElementById("sideload-msg").style.display = "none";
    
    document.getElementById("run").onclick = run;
    document.getElementById("createConfig").onclick = createConfig;
    document.getElementById("testConfig").onclick = testConfig;
    document.getElementById("testJsonFile").onclick = testJsonFile;
    
    $(function () {
      $('[data-toggle="tooltip"]').tooltip()
    });

    resetAnalysisInfo();
  }
});

//#endregion


//#region WorkSheet Custom Properties
async function resetCustomProperties()
{
  
  sheet.customProperties.load();
  await ctx.sync();

  // Iterate over all custom properties and delete them
  for (const key in sheet.customProperties.items) {
    if (sheet.customProperties.items.hasOwnProperty(key)) {
      const customProperty = sheet.customProperties.items[key];
      customProperty.delete();
    }
  }
  await ctx.sync();
}

async function addCustomProperty(key:string, value:string):Promise<Excel.CustomProperty>
{
    sheet.load(["customProperties"]);await ctx.sync();
    sheet.customProperties.add(key, value); await ctx.sync();
    sheet.customProperties.load(["key", "value", "type"]);await ctx.sync();
    var cp:any = sheet.customProperties.getItemOrNullObject(key); await ctx.sync();
    return cp as Excel.CustomProperty;
}

async function GetCustomPropertyValue(key:string):Promise<string>
{
    sheet.load(["customProperties"]);await ctx.sync();
    sheet.customProperties.load(["items"]);await ctx.sync();
    var cp = sheet.customProperties.getItemOrNullObject(key); await ctx.sync();
    cp.load(["key", "value", "type"]);await ctx.sync();
    return cp.value ?? "" ;
}
//#endregion


//#region Create Table
async function CreateTable(context, keepFormats:boolean=false) : Promise<boolean>
{ 
  try {

    sheet.load(["tables"]); await ctx.sync();

    const tblCount = sheet.tables.getCount(); await ctx.sync();
    if (tblCount.value > 1) {
      addAnalysisInfo("CreateTable Error", tblCount.value,`There is more than 1 table in the current worksheet. This is not supported. `,"None, or 1 table only supported",enumTypeAnalysis.Danger);
      return false;
    }

    
    if (tblCount.value == 1) 
    {
      AddMessage(`1 table found in current worksheet. Clearing formats`);
      
      tbl = sheet.tables.getItemAt(0); await ctx.sync();
      tblRange = tbl.getRange(); await ctx.sync();
      tblRange.clear("Formats"); await ctx.sync();
      tbl.convertToRange();await ctx.sync();
      resetCustomProperties();
      await context.sync();
      AddMessage(`Table cleared!`);
    }

    // Define the range of cells you want to select
    const range = sheet.getUsedRange();
    range.load("address");
    await ctx.sync();

    // // Create a table from the selected range
    // tbl = sheet.tables.getItemOrNullObject("CDL");
    // tblRange = tbl.getRange();
    // await context.sync().catch((error) => {
    //   addAnalysisInfo("error", 0, error, "create table", enumTypeAnalysis.Danger);
    //   AddMessage(error);
    // });

    // if (tbl.isNullObject) {
      const randomNumber = Math.ceil(Math.random() * 999999);
      tbl = sheet.tables.add(range, true /* hasHeaders */); await ctx.sync();
      // tbl.name = `CDL${randomNumber}`; 
      // tbl.name = "CDLRS"; 
      tblRange=tbl.getRange();
      await context.sync();
    // }

    if (keepFormats) return; //just create table and leave

    tblRange.clear("Formats");
    resetCustomProperties();
    await context.sync();

    var _isOrganizer = await isOrganizer();
    if (_isOrganizer) {
      tbl.style = organizerTableStyle;  
      sheet.tabColor = organizerTabColor;
      addCustomProperty("Organizer", "true");
    }
    else{
      tbl.style = attendeeTableStyle;  
      sheet.tabColor = attendeeTabColor;
      
      addCustomProperty("Organizer", "false");
    }

    tbl.load('tableStyle');
    tbl.columns.load();
    tblRange = tbl.getRange();
    await context.sync();

    return true;
    
  } catch (error) {
    console.error(error);
    addAnalysisInfo("create Table", 0, `Error creating table ${error}`, "Create Table", enumTypeAnalysis.Danger );   
    return false;  
  }
}

//#endregion 

//#region Filters methods
async function ClearFilters(ColumnName:string, value:string)
{
  tbl.clearFilters();
  AddMessage(`Table filters cleared`);  
}

async function SetFilter(ColumnName:string, value:string)
{
  tbl.columns.load();
  await ctx.sync();
  let columnFilter = tbl.columns.getItemOrNullObject(ColumnName).filter;
  await ctx.sync();
  columnFilter.apply({
    filterOn: Excel.FilterOn.values,
    values: [value]
  });
await ctx.sync();


AddMessage(`Column ${ColumnName} Filtered for value ${value}`);
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


AddMessage("Cells filtered. ");
}

//#endregion


//#region FormatDate (Legacy to remove)
async function FormatDateColumn( context, columnName:string){
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


//#endregion

//#region Highlights methods

async function HighlightIgnorable(){

  const colRange:Excel.Range = tbl.columns.getItemOrNullObject("Ignorable").getDataBodyRange();

  const conditionalFormat = colRange.conditionalFormats.add(
    Excel.ConditionalFormatType.containsText
  );

  conditionalFormat.textComparison.format.font.color = "blue";
  conditionalFormat.textComparison.format.fill.color="#ADD8E6";
  
  conditionalFormat.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: "TRUE"
  };
  await ctx.sync();

  return await CountFilterOccurrences(conditionalFormat.getRange());
  
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
  var r:Excel.Range = conditionalFormat.getRange();
  var n: number =  await CountFilterOccurrences(r);
  return n;

}

//#endregion

//#region Count filter Occurrences
async function CountFilterOccurrences( filterRange:Excel.Range ):Promise<number>
{

// var visibleTableRange: Excel.RangeView = tbl.getDataBodyRange().getVisibleView();
// visibleTableRange.load(["rowCount"]); await ctx.sync();
// var rows:number = visibleTableRange.rowCount;
// console.log(`rows Filtered  %d`, rows);

  await ctx.sync();
  // Get the range of cells affected by the conditional format
  const affectedRange = await filterRange.getIntersectionOrNullObject(tbl.getRange()); await ctx.sync();
  affectedRange.load(["rowCount"]); await ctx.sync();
  // Get the number of rows affected by the conditional format
  const rowCount = affectedRange ? affectedRange.rowCount : 0;
  return rowCount;
}

//#endregion


//#region json Methods
function isJSONString(str: any): boolean {
  try {
    const jsonObj = JSON.parse(str);
    return typeof jsonObj === "object" && jsonObj !== null;
  } catch (e) {
    return false;
  }
}

async function applyJsonConfig(json: any, hideLessRelevants:boolean=false): Promise<boolean> {
  
  var jsonArray;
  var  retval:boolean=true;
  
  if (typeof json === "string") {
    if (!isJSONString(json)) {
      throw new Error("Invalid JSON string");
    }
    jsonArray = JSON.parse(json);
  } else if (typeof json === "object") {
    jsonArray = json;
  } else {
    throw new Error("Parameter must be a string or an object");
  }
  
  if (typeof jsonArray != "object") 
  {
    throw new Error("Input JSON is not an Object.");
  }

  retval = retval && await applyJsonColDefinitions(jsonArray, hideLessRelevants);
  retval = retval && await applyJSONHighlights(jsonArray);
  retval = retval && await applyJSONFilters(jsonArray);
  return retval;
}

async function applyJsonColDefinitions(jsonArray:any, hideLessRelevants:boolean=false): Promise<boolean>
{
  try 
  {
    //jsonArray = JSON.parse(json);

    for (const element of jsonArray.columns) 
    {
        var tblCol:Excel.TableColumn;
        var tblColRange:Excel.Range;
        var tblColFormat:Excel.RangeFormat;
        if (element.columnName == undefined || element.columnName == "") 
        {
          AddMessage("Skipping json element as it is undefined ColumnName");
          continue;
        }

        tblCol = tbl.columns.getItemOrNullObject(element.columnName);
        tblCol.load(["isNullObject"]);
        tblColRange = tblCol.getDataBodyRange();
        tblColRange.load(["format"]);
        tblColFormat = tblColRange.format;
        tblColFormat.load(["horizontalAlignment", "verticalAlignment"]);
        ctx.trackedObjects.add([tblCol, tblColRange]);;
        await ctx.sync();

        if (tblCol.isNullObject) 
        {
          if (element.isMandatory !== undefined && element.isMandatory !== "" && element.isMandatory == "false") {
            AddMessage(`isMandatory: ${element.isMandatory}`);
            continue;
          }
          AddMessage(`Column Name does not exist: ${element.columnName}`);
          addAnalysisInfo("columnName",0,`Column Name does not exist: ${element.columnName}`, "ValidateJSONStruct", enumTypeAnalysis.Danger);
          return false;
        }


        AddMessage(`Column Name: ${element.columnName}`);

        if (element.style !== undefined && element.style !== "") {
          //style must be the first prop to set as it overrides all the below props
          AddMessage(`Style: ${element.style}`);
          tblColRange.style = element.style;
          // await ctx.sync();
        }

        if (element.searchFor !== undefined && element.searchFor !== "") {
          //Replacing values before applying remaining styles (dates come with Z )
          AddMessage(`SearchFor/ReplaceWith: ${element.searchFor} / ${element.ReplaceWith}`);
          const criteria: Excel.ReplaceCriteria = {
            completeMatch: false, /* Use a complete match to skip cells that already say "okay". */
            matchCase: true /* Ignore case when comparing strings. */
          };
          tblColRange.replaceAll(element.searchFor, element.replaceWith, criteria);
          await ctx.sync();
        }
      

        if (element.horizontalAlignment !== undefined && element.horizontalAlignment !== "") {
          AddMessage(`Horizontal Alignment: ${element.horizontalAlignment}`);
          tblColRange.format.horizontalAlignment =jsonConfigUtils.convertToHorizontalAlignment(element.horizontalAlignment);
        }
  
        if (element.verticalAlignment !== undefined && element.verticalAlignment !== "") {
            AddMessage(`Vertical Alignment: ${element.verticalAlignment}`);
            tblColRange.format.verticalAlignment = jsonConfigUtils.convertToVerticalAlignment(element.verticalAlignment);
        }
  
        if (element.columnWidth !== undefined && element.columnWidth !== null) {
          AddMessage(`Column Width: ${element.columnWidth}`);
          tblColRange.format.columnWidth = element.columnWidth;
          // await ctx.sync();
        }
  
        if (element.indentLevel !== undefined && element.indentLevel !== null) {
          AddMessage(`Indent Level: ${element.indentLevel}`);
          tblColRange.format.indentLevel = element.indentLevel;
          // await ctx.sync();
        }
  
       
        if (element.numberFormat !== undefined && element.numberFormat !== "") {
          AddMessage(`Style: ${element.numberFormat}`);
          tblColRange.numberFormat = element.numberFormat;
          // await ctx.sync();
        }

        if (element.visible !== undefined && element.visible !== null) {
          AddMessage(`Visible: ${element.visible}`);
          if (!element.visible && hideLessRelevants)
          {
            tblColRange.columnHidden = true;
          }  
          // await ctx.sync();
        }

        if (element.autosizeColumn !== undefined && element.autosizeColumn !== null) {
          AddMessage(`autosizeColumn: ${element.autosizeColumn}`);
          if (element.autosizeColumn==="true")  tblColRange.format.autofitColumns();
          // await ctx.sync();
        }

        await ctx.sync();
        AddMessage(`removing tracked objects for ${element.columnName}`);
        ctx.trackedObjects.remove([tblCol, tblColRange]);;               
    }

    return true;

  } catch (error) {
    addAnalysisInfo("columnName",0,`Error traversing JSON array: ${error}`, "ValidateJSONStruct", enumTypeAnalysis.Danger);
    console.error(`Error traversing JSON array: ${error}`);
    return false;
  }

}

async function applyJSONHighlights(jsonArray) : Promise<boolean>
{
    var retval:boolean=true;
    AddMessage(jsonArray);

    for (const element of jsonArray.conditionalFormats) 
    {
      
      var e:ConditionalFormat = element;
      

      switch (e.Type) {
        case enumConditionalFormatType.ColorScale:
          await createConditionalFormatColorScale(e);
          break;

        case enumConditionalFormatType.ContainsText:
          await createConditionalFormatContainsText(e);
          break;

        case enumConditionalFormatType.CellValue:
          break;

        case enumConditionalFormatType.Custom:
          break;
          
      
        default:
          break;
      }
    }

    return retval;
}


async function applyJSONFilters(jsonArray) : Promise<boolean>
{
  var retval:boolean=true;
  AddMessage(jsonArray);

  for (const element of jsonArray.filterDefinitions) 
  {
    var e:FilterDefinition = element;
    
    if (e.FilterActiveByUIOnly == false) {
      createColumnFilter(e);
    }
  }
  return retval;
}



async function createConditionalFormatColorScale(cf:ConditionalFormat): Promise<boolean>
{
    AddMessage(`Formatting column ${cf.FriendlyName}`);
    var retval:boolean = true;
    const col:Excel.TableColumn = tbl.columns.getItemOrNullObject(cf.ColumnName);

    col.load(["isNullObject"]); await ctx.sync();
    if (col.isNullObject) {
      return false;
    }

    const r:Excel.Range = col.getDataBodyRange();
    const excelCF = r.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
    excelCF.load(["colorScale"]); await ctx.sync();
    excelCF.colorScale.criteria.maximum.color = cf.ColorScaleColorMaximum;
    excelCF.colorScale.criteria.minimum.color = cf.ColorScaleColorMinimum;
    await ctx.sync();
    return retval;
}

async function createConditionalFormatContainsText(cf:ConditionalFormat): Promise<boolean>
{
  AddMessage(`Creating Conditional Format ${cf.FriendlyName}`);
  var retval:boolean = true;
  const col:Excel.TableColumn = tbl.columns.getItemOrNullObject(cf.ColumnName);

  col.load(["isNullObject"]); await ctx.sync();
  if (col.isNullObject) {
    return false;
  }

  const r:Excel.Range = col.getDataBodyRange();
  const excelCF = r.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
  excelCF.load(["textComparison", "format", "format/fill", "format/font"]); await ctx.sync();
  
  excelCF.textComparison.rule = {
    operator: Excel.ConditionalTextOperator.contains,
    text: cf.ContainsTextSearch
  };
  
  if(cf.FillColor !== undefined && cf.FillColor !== null && cf.FillColor.toLowerCase() !== "null") {
    excelCF.textComparison.format.fill.color = cf.FillColor;
  }

  if(cf.FontColor !== undefined && cf.FontColor !== null && cf.FontColor.toLowerCase() !== "null") {
    excelCF.textComparison.format.font.color = cf.FontColor;
  }

  await ctx.sync();
  return retval;

}

async function createColumnFilter(f:FilterDefinition): Promise<boolean>
{
    AddMessage(`Creating Filter ${f.FriendlyName}`);
    var retval:boolean = true;

    const col:Excel.TableColumn = tbl.columns.getItemOrNullObject(f.ColumnName);

    col.load(["isNullObject"]); await ctx.sync();
    if (col.isNullObject) {
      return false;
    }
  
    const r:Excel.Range = col.getDataBodyRange();

    const excelFilter = col.filter;
    await ctx.sync();
    excelFilter.apply({
      filterOn: Excel.FilterOn.values,
      values: [f.FilterValue]
    });
  await ctx.sync();

  return retval;

}

async function createColumnDefinitionsFromTable(jsonConfigUtils:JsonConfigUtils): Promise<JsonConfigUtils> 
{
  
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
      const headerCell:Excel.Range = column.getRange().getCell(0, 0);
      headerCell.load(["columnHidden"]);
      await ctx.sync();
      const isVisible = headerCell.columnHidden;

      var r: Excel.Range = column.getDataBodyRange();
      r.load(['style', 'numberFormat', 'columnHidden']);
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
       const visible:boolean = !headerCell.columnHidden;
       
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
         visible: visible, 
         autosizeColumn:false,
         searchFor:"",
         replaceWith:""
       };
       jsonConfigUtils.addColumn(colDef);

       AddMessage(`Adding column ${name}/format ${numberFormat}/ style ${style}`);
       
   }

  return jsonConfigUtils;
}

async function createFiltersFromTable(jsonConfig:JsonConfigUtils): Promise<any>
{
  ctx.trackedObjects.add(tbl); await ctx.sync();
  tblRange.conditionalFormats.load();await ctx.sync();
  const conditionalFormats = tblRange.conditionalFormats;await ctx.sync();
  const results = [];

  for (const column of tbl.columns.items) 
   {
      column.load(["name", "filter"]);await ctx.sync();

      const filter = column.filter;

      if (filter.criteria != null) {
        const randomNumber = Math.ceil(Math.random() * 999999);
        const columnName:string = column.name;
        const key:string = columnName + randomNumber.toString();
        var values:string;
        if (filter.criteria.filterOn == Excel.FilterOn.custom) 
        {
          const criterion2 = filter.criteria.criterion2 ?? '';
          const correctedString = criterion2.toString() !== '' ? `, ${criterion2.toString()}` : '';

          values = filter.criteria.criterion1.toString() + correctedString;         
        }
        else  
        {
           values = filter.criteria.values.join(", ") ;
        }
        jsonConfig.addFilterCondition(columnName, values,key);

        AddMessage(`Add filter ${columnName} / ${values} / ${key}`);
        
      }

   }
}

async function createConditionalFormatsFromTable(jsonConfig:JsonConfigUtils): Promise<any[]> {
  ctx.trackedObjects.add(tbl); await ctx.sync();
  tblRange.conditionalFormats.load();await ctx.sync();
  const conditionalFormats = tblRange.conditionalFormats;await ctx.sync();
  const results = [];
  
  
  for (const column of tbl.columns.items) 
   {
    const cfs:Excel.ConditionalFormatCollection = column.getDataBodyRange().conditionalFormats;
    cfs.load("items");await ctx.sync();
    for (const cf of cfs.items) {
        const randomNumber = Math.ceil(Math.random() * 999999);
        AddMessage(`Column ${column.name} has CF of type ${cf.type}`);  
        
        switch (cf.type) 
        {
            case Excel.ConditionalFormatType.cellValue:
              cf.cellValue.load(["rule", "format", "format/font", "format/fill"]);await ctx.sync();
              jsonConfig.addConditionalFormatCellValue(`${column.name} Cell Value ${cf.cellValue.rule.formula1}`, 
                            column.name, false,false,
                            "Normal", cf.cellValue.format.font.color, cf.cellValue.format.fill.color,
                             cf.cellValue.rule.formula1,cf.cellValue.rule.formula2,
                             CellValueOperatorToJsonEnum(cf.cellValue.rule.operator));

              break;
            case Excel.ConditionalFormatType.containsText:
              cf.textComparison.load(["rule", "format", "format/font", "format/fill"]);await ctx.sync();
              
              jsonConfig.addConditionalFormatContainsText(`${column.name} Contains Text ${cf.textComparison.rule.text}`, 
                            column.name, false,false,
                            "Normal",cf.textComparison.format.fill.color,cf.textComparison.format.font.color,
                            cf.textComparison.rule.text, 
                            cf.textComparison.rule.operator == Excel.ConditionalTextOperator.contains ? 
                                                                enumConditionalFormatTextOperator.Contains : 
                                                                enumConditionalFormatTextOperator.NotContains);
              break;
            case Excel.ConditionalFormatType.custom:
              break;
            case Excel.ConditionalFormatType.colorScale:
              cf.colorScale.load(["rule", "format", "format/font", "format/fill", "criteria"]);await ctx.sync();
              
              jsonConfig.addConditionalFormatColorScale(`${column.name} Color Scale ${cf.colorScale.criteria.maximum.color} Down to ${cf.colorScale.criteria.minimum.color}`, 
                            column.name, false,false,
                            "Normal",null,null,
                            cf.colorScale.criteria.minimum.color, cf.colorScale.criteria.maximum.color);
              break;
            case Excel.ConditionalFormatType.dataBar:
              break;
              
          default:
            break;
        }
    }
    
   }
  
return results;
}


async function getJsonData(): Promise<any> 
{
  var jsonType:string = await typeCDL();
  var response;

  switch (jsonType) {
    case "rave-diag-log":
      response = await fetch("./RaveCDLconfig.json");
      break;

    case "exo-cdl":
      response = await fetch("./EXOCDLconfig.json");
      break;

    default:
      response = await fetch("./RaveCDLconfig.json");
      break;
  }
  
  const jsonData = await response.json();
  return jsonData;
}

//#endregion



//#region Analysys
async function PerformAnalysis(context) {
  //await resetAnalysisInfo();
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

//#endregion


//#region Config section
export async function createConfig()
{
  myConsole.reset();
  await Excel.run(async (context) => {
    ctx = context;
    sheet = context.workbook.worksheets.getActiveWorksheet();
    tbl = sheet.tables.getItemOrNullObject("CDL");
    tblRange = tbl.getRange();
    jsonConfigUtils = new JsonConfigUtils();
    

    await context.sync() ;

    const validTable:boolean = await  CreateTable(context, true); //keep formatting for json creation

    if (!validTable) {
      return;
    }

    await createColumnDefinitionsFromTable(jsonConfigUtils);
    await createConditionalFormatsFromTable(jsonConfigUtils);
    await createFiltersFromTable(jsonConfigUtils);
    document.getElementById("jsonConfig").textContent = jsonConfigUtils.getValue();

  });
  
}


export async function testJsonFile()
{
  myConsole.reset();
  var tempJson = await getJsonData();
  document.getElementById("jsonConfig").textContent = JSON.stringify(tempJson);
  var isTableValid:boolean = await applyJsonConfig(tempJson);
}

export async function testConfig()
{
  myConsole.reset();
  var textbox:any = document.getElementById("jsonConfig");
  var tempJson = textbox.value;
  var isTableValid:boolean = await applyJsonConfig(tempJson);
}
//#endregion

//Main Function
export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      myConsole.reset();
      await showSpinner(true);
      await resetAnalysisInfo();
      await AddMessage("Starting Processing");

      ctx = context;
      sheet = context.workbook.worksheets.getActiveWorksheet();
      jsonConfigUtils = new JsonConfigUtils();

      // tbl = sheet.tables.getItemOrNullObject("CDL");
      // tblRange = tbl.getRange();
      // tbl.clearFilters();
      sheet.getRange().clear("Formats");
      sheet.getRange().conditionalFormats.clearAll();
      
      sheet.freezePanes.unfreeze();
      await context.sync();


      const validTable:boolean = await  CreateTable(context);
      if (!validTable) {
        return;
      }
      AddMessage("Create Table Done");

      jsonLog = await getJsonData();
      var isTableValid:boolean = await applyJsonConfig(jsonLog);
      if (!isTableValid) {
        addAnalysisInfo("CDL Invalid", 0, "CDL Structure is invalid (check previous exceptions)", "CDLInvalid", enumTypeAnalysis.Danger);
        showSpinner(false);
        return;
      }

      //format section
      await freezeColumns("Ignorable");
      // await FormatCells(context).then(()=>{AddMessage("Format cells Done")});
      // await FormatRawFrom(context).then(()=>{AddMessage("Format raw from Done")});
      // await FormatDateColumn(context, "ModifiedDate").then(()=>{AddMessage("Format ModifiedDate Done")}); //ModifiedDate
      // await FormatDateColumn(context, "StartTime").then(()=>{AddMessage("Create StartTime Done")}); //StartTime
      // await FormatDateColumn(context, "EndTime").then(()=>{AddMessage("Create End Done")}); //EndTime

      //highlight section
      // await HighlightIgnorable().then(()=>{AddMessage("Highlight Ignorable Done")});
      // await HighlightApptSequence(context).then(()=>{AddMessage("Highlight  Done")});
      // await HighlightCRA(context).then(()=>{AddMessage("Highlight CRA Done")});
      
      // //await addAnalysisInfo("CRA Found", rowCount, "CRA Events were found, meaning calendar state was not 100%","HighlightCRA",enumTypeAnalysis.Warning);
      // await HighLightCreates(context).then(()=>{AddMessage("Highlight Create Done")});

      //Filters section
      //await FilterIgnorable("FALSE").then(()=>{AddMessage("Filter Ignorable Done")});

      await context.sync();
      
      await PerformAnalysis(context).then(()=>{AddMessage("Perform Analysis Done")});
      AddMessage(`Processing done.`);
      
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
    addAnalysisInfo("Error",0,error, "Run/Catch", enumTypeAnalysis.Danger);
  }
  
  
}


function CellValueOperatorToJsonEnum(operator: string): enumCellValueOperator {
  //throw new Error('Function not implemented.');
  return enumCellValueOperator.EQ;
}

