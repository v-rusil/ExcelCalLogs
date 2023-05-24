
export interface CalendarItemTypes {
    [key: string]: string;
  }
  
  export const calendarItemTypes: CalendarItemTypes = {
    'IPM.Schedule.Meeting.Request.AttendeeListReplication': "AttendeeList",
    'IPM.Schedule.Meeting.Canceled': "Canceled",
    'IPM.OLE.CLASS.{00061055-0000-0000-C000-000000000046}': "ExceptionMsgClass",
    'IPM.Schedule.Meeting.Notification.Forward': "ForwardNotification",
    'IPM.Appointment': "IpmAppointment",
    'IPM.Schedule.Meeting.Request': "MeetingRequest",
    'IPM.CalendarSharing.EventUpdate': "SharingCFM",
    'IPM.CalendarSharing.EventDelete': "SharingDelete",
    'IPM.Schedule.Meeting.Resp': "RespAny",
    'IPM.Schedule.Meeting.Resp.Neg': "RespNeg",
    'IPM.Schedule.Meeting.Resp.Tent': "RespTent",
    'IPM.Schedule.Meeting.Resp.Pos': "RespPos",
  };
  
  export interface SCN {
    [key: string]: string;
  }
  
  export const scn: SCN = {
    'Client=Hub Transport': "Transport",
    'Client=MSExchangeRPC': "Outlook",
    'Lync for Mac': "LyncMac",
    'AppId=00000004-0000-0ff1-ce00-000000000000': "SkypeMMS",
    'MicrosoftNinja': "Teams",
    'Remove-CalendarEvents': "RemoveCalendarEvent",
    'Client=POP3/IMAP4': "PopImap",
    'Client=OWA': "OWA",
    'PublishedBookingCalendar': "BookingAgent",
    'LocationAssistantProcessor': "LocationProcessor",
    'AppId=6326e366-9d6d-4c70-b22a-34c7ea72d73d': "CalendarReplication",
    'AppId=1e3faf23-d2d2-456a-9e3e-55db63b869b0': "CiscoWebex",
    'AppId=1c3a76cc-470a-46d7-8ba9-713cfbb2c01f': "Time Service",
    'AppId=48af08dc-f6d2-435f-b2a7-069abd99c086': "RestConnector",
    'GriffinRestClient': "GriffinRestClient",
    'MacOutlook': "MacOutlookRest",
    'Outlook-iOS-Android': "OutlookMobile",
    'Client=OutlookService;Outlook-Android': "OutlookAndroid",
    'Client=OutlookService;Outlook-iOS': "OutlookiOS",
  };
  
  export interface RT {
    [key: string]: string;
  }
  
  export const rt: RT = {
    '0': "None",
    '1': "Organizer",
    '2': "Tentative",
    '3': "Accept",
    '4': "Decline",
    '5': "Not Responded",
  };
  