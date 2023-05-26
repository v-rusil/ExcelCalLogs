//read worksheet table row by row and compare with previous row. output differences according to below rules
//1. if the row is new, output the row (Subject: col.NormalizedSubject, Organizer: col.SentRepresentingDisplayName, Sender: col.SentRepresentingDisplayName, To List: col.DisplayAttendeesAll, Location: col.Location)
//generate code below

import { calendarItemTypes, rt, scn } from "./hashTablesCDLTimeline";
import { timelineDataObject } from "./timelineDataObject";
import { timelineDataObjectDetail } from "./timelineDataObjectDetail";

export class CDLToEnglishActionsTimeline 
{
   
    
    //create property tbl:Excel.Table;
    private __tbl : Excel.Table;
    public get _tbl() : Excel.Table {
        return this.__tbl;
    }
    public set _tbl(v : Excel.Table) {
        this.__tbl = v;
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

    
    public async processCDLToEnglishActionsTimeline():Promise<timelineDataObject> {
        let timelineData: timelineDataObject = new timelineDataObject();
        timelineData.generateDummyData();
        return timelineData;
    }





    MeetingSummary(
        time: any[],
        meetingChanges: any[],
        entry: any,
        longVersion: boolean,
        shortVersion: boolean
      ): void {
        let initialSubject = "Subject: " + entry.NormalizedSubject;
        let initialOrganizer = "Organizer: " + entry.SentRepresentingDisplayName;
        let initialSender = "Sender: " + entry.SentRepresentingDisplayName;
        let initialToList = "To List: " + entry.DisplayAttendeesAll;
        let initialLocation = "Location: " + entry.Location;
      
        if (shortVersion || longVersion) {
          let initialStartTime = "StartTime: " + entry.StartTime.toString();
          let initialEndTime = "EndTime: " + entry.EndTime.toString();
        }
      
        let initialTimeZone = "";
        if (longVersion && entry.Timezone !== "") {
          initialTimeZone = "Time Zone: " + entry.Timezone;
        } else {
          initialTimeZone = "Time Zone: Not Populated";
        }
      
        let initialRecurring = "";
        if (entry.AppointmentRecurring) {
          initialRecurring = "Recurring: Yes - Recurring";
        } else {
          initialRecurring = "Recurring: No - Single instance";
        }
      
        let initialRecurrencePattern = "";
        let initialSeriesStartTime = "";
        let initialSeriesEndTime = "";
        let initialEndDate = "";
        if (longVersion && entry.AppointmentRecurring) {
          initialRecurrencePattern = "RecurrencePattern: " + entry.RecurrencePattern;
          initialSeriesStartTime = "Series StartTime: " + entry.ViewStartTime.toString();
          initialSeriesEndTime = "Series EndTime: " + entry.ViewStartTime.toString();
          if (!entry.ViewEndTime) {
            initialEndDate = "Meeting Series does not have an End Date.";
          }
        }
      
        if (!time) {
          //time = CalLog.LastModifiedTime.toString();
        }
      
        if (!meetingChanges) {
          meetingChanges = [];
          meetingChanges.push(
            initialSubject,
            initialOrganizer,
            initialSender,
            initialToList,
            initialLocation,
            //initialStartTime,
            //initialEndTime,
            initialTimeZone,
            initialRecurring,
            initialRecurrencePattern,
            initialSeriesStartTime,
            initialSeriesEndTime,
            initialEndDate
          );
        }
      
        if (shortVersion) {
          meetingChanges = [];
          meetingChanges.push(
            initialToList,
            initialLocation,
            //initialStartTime,
            //initialEndTime,
            initialRecurring
          );
        }
      
        //ConvertData(["Time", "MeetingChanges"]);
      }
      

}