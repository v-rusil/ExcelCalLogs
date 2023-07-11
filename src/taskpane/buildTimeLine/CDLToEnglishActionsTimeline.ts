//read worksheet table row by row and compare with previous row. output differences according to below rules
//1. if the row is new, output the row (Subject: col.NormalizedSubject, Organizer: col.SentRepresentingDisplayName, Sender: col.SentRepresentingDisplayName, To List: col.DisplayAttendeesAll, Location: col.Location)
//generate code below

import { IPreviousRowChanged, getTableCellRange, isPreviousRowChanged, tryReadTblValue } from "../excelUtils";
import { myConsole } from "../myConsole";
import { calendarItemTypes, rt, scn } from "./hashTablesCDLTimeline";
import { timelineDataObject } from "./timelineDataObject";
import { timelineDataObjectDetail } from "./timelineDataObjectDetail";
import { timelineDataObjectDetailChangedProperties } from "./timelineDataObjectDetailChangedProperties";

export class CDLToEnglishActionsTimeline 
{
   
    
    //create property tbl:Excel.Table;
    private _tbl : Excel.Table;
    public get tbl() : Excel.Table {
        return this._tbl;
    }
    public set tbl(v : Excel.Table) {
        this._tbl = v;
    }

    
    private _ctx : Excel.RequestContext;
    public get ctx() : Excel.RequestContext {
        return this._ctx;
    }
    public set ctx(v : Excel.RequestContext) {
        this._ctx = v;
    }
    
    

    constructor(tbl: Excel.Table, ctx: Excel.RequestContext) {
        this._tbl = tbl;
        this._ctx = ctx;
    }

    public async processCDLToEnglishActionsTimeline():Promise<timelineDataObject> 
    {
        let timelineData: timelineDataObject = new timelineDataObject();
        var i:number = 0;

        this.tbl.load("rows"); await this.ctx.sync();
        this.tbl.rows.load("items"); await this.ctx.sync();
        
        var totalRows:number = this.tbl.rows.count;

        if (totalRows < 1) {
          return timelineData;
        }
      
        await this.readHeaderData(this.ctx, this.tbl, timelineData, i);
        //read header data for first row
        var row:Excel.TableRow = this.tbl.rows.getItemAt(i);
        
        i++;
        while (i < totalRows) {
          //read header data for next row

          i++;
        }
        
        return timelineData;
    }
    
    
    public async readHeaderData(ctx:Excel.RequestContext, tbl:Excel.Table, timelineData:timelineDataObject, i:number):Promise<void> {
      timelineData.header.subject = await tryReadTblValue(this.ctx, this.tbl, i, "NormalizedSubject");
      timelineData.header.startTime = await tryReadTblValue(this.ctx, this.tbl, i, "StartTime");
      timelineData.header.endTime = await tryReadTblValue(this.ctx, this.tbl, i, "EndTime");
      timelineData.header.organizer = await tryReadTblValue(this.ctx, this.tbl, i, "SentRepresentingDisplayName");
      timelineData.header.sender = await tryReadTblValue(this.ctx, this.tbl, i, "SentRepresentingDisplayName");
      timelineData.header.to = await tryReadTblValue(this.ctx, this.tbl, i, "DisplayAttendeesAll");
      timelineData.header.location = await tryReadTblValue(this.ctx, this.tbl, i, "Location");
      timelineData.header.timeZone = await tryReadTblValue(this.ctx, this.tbl, i, "Timezone");
      timelineData.header.isRecurring = await tryReadTblValue(this.ctx, this.tbl, i, "AppointmentRecurring");
      timelineData.header.recurrencePattern = await tryReadTblValue(this.ctx, this.tbl, i, "RecurrencePattern");
      timelineData.header.recurrenceStartTime = await tryReadTblValue(this.ctx, this.tbl, i, "ViewStartTime");
      timelineData.header.recurrenceEndTime = await tryReadTblValue(this.ctx, this.tbl, i, "ViewEndTime");
    }

    public async readDetailData(ctx:Excel.RequestContext, tbl:Excel.Table, timelineData:timelineDataObject, i:number, previousRow:number):Promise<void>
    {
      let detail: timelineDataObjectDetail = new timelineDataObjectDetail();
      detail.rowIndex = i;
      detail.logRow = i + 1;
      detail.columnName = "NormalizedSubject";
      detail.description = await tryReadTblValue(this.ctx, this.tbl, i, "NormalizedSubject");
      detail.changedProperties = await this.readChangedProperties(this.ctx, this.tbl, timelineData, i, previousRow);
      timelineData.details.push(detail);
    }


    public async readChangedProperties(ctx:Excel.RequestContext, tbl:Excel.Table, timelineData:timelineDataObject, i:number, previousRow:number):Promise<timelineDataObjectDetailChangedProperties[] | undefined> 
    {
      let changedProperties: timelineDataObjectDetailChangedProperties[] = [];
        try
        {
            myConsole.log(`Reading changed properties for row ${i} Column NormalizedSubject`);
            var prevChangedData: timelineDataObjectDetailChangedProperties | undefined = undefined;
            await isPreviousRowChanged(ctx, tbl, i, previousRow, "NormalizedSubject").then((isChanged: IPreviousRowChanged | undefined) => {
                if (isChanged !== undefined && isChanged.isChanged) {
                    prevChangedData = new timelineDataObjectDetailChangedProperties();
                    prevChangedData.columnName = "NormalizedSubject";
                    prevChangedData.description = "Subject changed from " + isChanged.previousValue + " to " + isChanged.currentValue;
                    changedProperties.push(prevChangedData);
                }
            });
            myConsole.log(`Reading changed properties for row ${i} Column SentRepresentingDisplayName`);
            await isPreviousRowChanged(ctx, tbl, i, previousRow, "SentRepresentingDisplayName").then((isChanged: IPreviousRowChanged | undefined) => {
                if (isChanged !== undefined && isChanged.isChanged) {
                    prevChangedData = new timelineDataObjectDetailChangedProperties();
                    prevChangedData.columnName = "SentRepresentingDisplayName";
                    prevChangedData.description = "Organizer changed from " + isChanged.previousValue + " to " + isChanged.currentValue;
                    changedProperties.push(prevChangedData);
                }
            });
            myConsole.log(`Reading changed properties for row ${i} Column DisplayAttendeesAll`);
            await isPreviousRowChanged(ctx, tbl, i, previousRow, "DisplayAttendeesAll").then((isChanged: IPreviousRowChanged | undefined) => {
                if (isChanged !== undefined && isChanged.isChanged) {
                    prevChangedData = new timelineDataObjectDetailChangedProperties();
                    prevChangedData.columnName = "DisplayAttendeesAll";
                    prevChangedData.description = "To List changed from " + isChanged.previousValue + " to " + isChanged.currentValue;
                    changedProperties.push(prevChangedData);
                }
            });
            myConsole.log(`Reading changed properties for row ${i} Column Location`);
            await isPreviousRowChanged(ctx, tbl, i, previousRow, "Location").then((isChanged: IPreviousRowChanged | undefined) => {
                if (isChanged !== undefined && isChanged.isChanged) {
                    prevChangedData = new timelineDataObjectDetailChangedProperties();
                    prevChangedData.columnName = "Location";
                    prevChangedData.description = "Location changed from " + isChanged.previousValue + " to " + isChanged.currentValue;
                    changedProperties.push(prevChangedData);
                }
            });
            myConsole.log(`Reading changed properties for row ${i} Column StartTime`);
            await isPreviousRowChanged(ctx, tbl, i, previousRow, "StartTime").then((isChanged: IPreviousRowChanged | undefined) => { 
                if (isChanged !== undefined && isChanged.isChanged) {
                    prevChangedData = new timelineDataObjectDetailChangedProperties();
                    prevChangedData.columnName = "StartTime";
                    prevChangedData.description = "Start Time changed from " + isChanged.previousValue + " to " + isChanged.currentValue;
                    changedProperties.push(prevChangedData);
                }
            });
            myConsole.log(`Reading changed properties for row ${i} Column EndTime`);
            await isPreviousRowChanged(ctx, tbl, i, previousRow, "EndTime").then((isChanged: IPreviousRowChanged | undefined) => {
                if (isChanged !== undefined && isChanged.isChanged) {
                    prevChangedData = new timelineDataObjectDetailChangedProperties();
                    prevChangedData.columnName = "EndTime";
                    prevChangedData.description = "End Time changed from " + isChanged.previousValue + " to " + isChanged.currentValue;
                    changedProperties.push(prevChangedData);
                }
            });
            myConsole.log(`Reading changed properties for row ${i} Column Timezone`);
            await isPreviousRowChanged(ctx, tbl, i, previousRow, "Timezone").then((isChanged: IPreviousRowChanged | undefined) => {
                if (isChanged !== undefined && isChanged.isChanged) {
                    prevChangedData = new timelineDataObjectDetailChangedProperties();
                    prevChangedData.columnName = "Timezone";
                    prevChangedData.description = "Timezone changed from " + isChanged.previousValue + " to " + isChanged.currentValue;
                    changedProperties.push(prevChangedData);
                }
            });
            myConsole.log(`Reading changed properties for row ${i} Column AppointmentRecurring`);
            await isPreviousRowChanged(ctx, tbl, i, previousRow, "AppointmentRecurring").then((isChanged: IPreviousRowChanged | undefined) => {
                if (isChanged !== undefined && isChanged.isChanged) {
                    prevChangedData = new timelineDataObjectDetailChangedProperties();
                    prevChangedData.columnName = "AppointmentRecurring";
                    prevChangedData.description = "Is Recurring changed from " + isChanged.previousValue + " to " + isChanged.currentValue;
                    changedProperties.push(prevChangedData);
                }
            });
            myConsole.log(`Reading changed properties for row ${i} Column RecurrencePattern`);
            await isPreviousRowChanged(ctx, tbl, i, previousRow, "RecurrencePattern").then((isChanged: IPreviousRowChanged | undefined) => {
                if (isChanged !== undefined && isChanged.isChanged) { 
                    prevChangedData = new timelineDataObjectDetailChangedProperties();
                    prevChangedData.columnName = "RecurrencePattern";
                    prevChangedData.description = "Recurrence Pattern changed from " + isChanged.previousValue + " to " + isChanged.currentValue;
                    changedProperties.push(prevChangedData);
                }
            });
            myConsole.log(`Reading changed properties for row ${i} Column ViewStartTime`);
            await isPreviousRowChanged(ctx, tbl, i, previousRow, "ViewStartTime").then((isChanged: IPreviousRowChanged | undefined) => {
                if (isChanged !== undefined && isChanged.isChanged) { 
                    prevChangedData = new timelineDataObjectDetailChangedProperties();
                    prevChangedData.columnName = "ViewStartTime";
                    prevChangedData.description = "Recurrence Start Time changed from " + isChanged.previousValue + " to " + isChanged.currentValue;
                    changedProperties.push(prevChangedData);
                }
            });
            myConsole.log(`Reading changed properties for row ${i} Column ViewEndTime`);
            await isPreviousRowChanged(ctx, tbl, i, previousRow, "ViewEndTime").then((isChanged: IPreviousRowChanged | undefined) => {
                if (isChanged !== undefined && isChanged.isChanged) { 
                    prevChangedData = new timelineDataObjectDetailChangedProperties();
                    prevChangedData.columnName = "ViewEndTime";
                    prevChangedData.description = "Recurrence End Time changed from " + isChanged.previousValue + " to " + isChanged.currentValue;
                    changedProperties.push(prevChangedData);
                }
            });

        }
        catch (error)
        {
            myConsole.log(`Error in readChangedProperties: ${error}`);
        }
        return changedProperties; ;
    }


    populateCardPane(timelineData: timelineDataObject): void {
      const cardTitleElement = document.querySelector("#timeline-subject");
      if (cardTitleElement) {
        cardTitleElement.textContent = timelineData.header.subject;
      }
    
      // const cardDateElement = document.querySelector("#timeline-date");
      // if (cardDateElement) {
      //   cardDateElement.textContent = timelineData.header.date.toDateString();
      // }
    
      const timelineBodyElement = document.querySelector("#timeline-body");
      if (timelineBodyElement) {
        timelineBodyElement.innerHTML = ""; // Clear existing content
    
        for (const detail of timelineData.details) {
          const rowElement = document.createElement("div");
          rowElement.classList.add("row");
          rowElement.addEventListener("click", () => {
            this.selectExcelTableRow(detail.rowIndex, detail.columnName);
          });
    
          const rowIndexElement = document.createElement("div");
          rowIndexElement.classList.add("col");
          rowIndexElement.textContent = detail.rowIndex.toString();
          rowElement.appendChild(rowIndexElement);
    
          const columnNameElement = document.createElement("div");
          columnNameElement.classList.add("col");
          columnNameElement.textContent = detail.columnName;
          rowElement.appendChild(columnNameElement);

          const descriptionElement = document.createElement("div");
          descriptionElement.classList.add("col");
          descriptionElement.textContent = detail.description;
          rowElement.appendChild(descriptionElement);

          const logRowElement = document.createElement("div");
          logRowElement.classList.add("col");
          logRowElement.textContent = detail.logRow.toString();
          rowElement.appendChild(logRowElement);

          const changedPropertiesElement = document.createElement("div");
          changedPropertiesElement.classList.add("col");
          changedPropertiesElement.textContent = detail.changedPropertiesToString();
          rowElement.appendChild(changedPropertiesElement); 

    
          timelineBodyElement.appendChild(rowElement);
        }
      }
    }
    
    public async selectExcelTableRow(rowIndex: number, columnName: string): Promise<void> {
      
      try
      {
        myConsole.log(`Selecting row ${rowIndex} in column ${columnName}`);
        const column = this.tbl.columns.getItem(columnName);
        this.tbl.load("rows");await this.ctx.sync();
        const row = this.tbl.rows.getItemAt(rowIndex - 1); // Adjusting to zero-based index
        await this.ctx.sync();

        const r:Excel.Range = await getTableCellRange(this.ctx,this.tbl, rowIndex, columnName);
        r.select();
        await this.ctx.sync();
        

        await this.ctx.sync();
        myConsole.log(`Selected row ${rowIndex} in column ${columnName}`);
        return;
    
      } 
      catch (error)
      {
        myConsole.log(error);
      }
    }

     

}


