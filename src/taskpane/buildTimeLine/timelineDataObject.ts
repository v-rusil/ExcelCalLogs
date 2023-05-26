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
        this.header.subjec = "Subject: Test Subject";
        for (let i = 0; i < 10; i++) {
            let detail: timelineDataObjectDetail = new timelineDataObjectDetail();
            detail.columnName = "Column " + i.toString();
            detail.detail = "Detail " + i.toString();
            detail.rowIndex = i;
            this.details.push(detail);
        }
        return this;
    }
}