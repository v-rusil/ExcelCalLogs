export class timelineDataObjectDetail
{
    
    private _rowIndex : number;
    public get rowIndex() : number {
        return this._rowIndex;
    }
    public set rowIndex(v : number) {
        this._rowIndex = v;
    }
    

    
    private _columnName : string;
    public get columnName() : string {
        return this._columnName;
    }
    public set columnName(v : string) {
        this._columnName = v;
    }
    

    
    private _detail : string;
    public get detail() : string {
        return this._detail;
    }
    public set detail(v : string) {
        this._detail = v;
    }
    

}