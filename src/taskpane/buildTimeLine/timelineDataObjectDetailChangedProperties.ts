export class timelineDataObjectDetailChangedProperties {
    private _columnName: string;
    private _description: string;

    constructor() {
        this._columnName = "";
        this._description = "";
    }
    
    get columnName() {
        return this._columnName;
    }
    set columnName(v) {
        this._columnName = v;
    }

    get description() {
        return this._description;
    }
    set description(v) {
        this._description = v;
    }

}