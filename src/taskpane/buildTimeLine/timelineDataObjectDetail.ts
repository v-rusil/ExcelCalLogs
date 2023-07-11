import { timelineDataObjectDetailChangedProperties } from "./timelineDataObjectDetailChangedProperties";

export class timelineDataObjectDetail
{
    
    constructor() {
        this._rowIndex = 0;
        this._logRow = 0;
        this._columnName = "";
        this._description = "";
        this._changedProperties = [];
    }

    private _rowIndex : number;
    public get rowIndex() : number {
        return this._rowIndex;
    }
    public set rowIndex(v : number) {
        this._rowIndex = v;
    }
    
    private _logRow : number;
    public get logRow() : number {  
        return this._logRow;
    }
    public set logRow(v : number) {
        this._logRow = v;
    }
    
    private _columnName : string;
    public get columnName() : string {
        return this._columnName;
    }
    public set columnName(v : string) {
        this._columnName = v;
    }
    
    private _description : string;
    public get description() : string {
        return this._description;
    }
    public set description(v : string) {
        this._description = v;
    }
    
    private _changedProperties : timelineDataObjectDetailChangedProperties[];
    public get changedProperties() : timelineDataObjectDetailChangedProperties[] {
        return this._changedProperties;
    }
    public set changedProperties(v : timelineDataObjectDetailChangedProperties[]) { 
        this._changedProperties = v;
    }

    public changedPropertiesToString(): string {    
        let result: string = "";
        this.changedProperties.forEach(element => {
            result += element.columnName + " " + element.description + "\n";
        });
        return result;
    }
}