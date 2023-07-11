import { timelineDataObjectDetail } from "./timelineDataObjectDetail";
import { timelineDataObjectHeader } from "./timelineDataObjectHeader";

export class timelineDataObject
{
 
    
    private _header : timelineDataObjectHeader;
    public get header() : timelineDataObjectHeader {
        return this._header;
    }
    public set header(v : timelineDataObjectHeader) {
        this._header = v;
    }
    
    
    private _details : timelineDataObjectDetail[];
    public get details() : timelineDataObjectDetail[] {
        return this._details;

    }
   public set details(v : timelineDataObjectDetail[]){
        this._details = v;
    }

    constructor() {
        this._header = new timelineDataObjectHeader();
        this._details = [];
    }
    
    public async generateDummyData(): Promise<timelineDataObject> {
        this.header.subject = "Subject: Test Subject";
        var d = new timelineDataObjectDetail();
        
        d.columnName = "Trigger";
        d.rowIndex = 13;
        this.details.push(d);

        d = new timelineDataObjectDetail();
        
        d.columnName = "LogRow";
        d.rowIndex = 113;
        this.details.push(d);

        for (let i = 0; i < 10; i++) {
            let detail: timelineDataObjectDetail = new timelineDataObjectDetail();
            detail.columnName = "Column " + i.toString();
            detail.description = "Detail " + i.toString();
            detail.rowIndex = i;
            this.details.push(detail);
        }
        return this;
    }
}