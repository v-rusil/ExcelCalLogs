import { ColumnDefinition, EnumColumnHorizontalAlignment, EnumColumnVerticalAlignment } from './columnDefinitions';

export class JsonConfigUtils 
{
  private columns: ColumnDefinition[] = [];
  private json: string;
  
  constructor() {
    // this.json = value;
  }
  
  getValue(): string {
    return this.json;
  }

  addColumn(col: ColumnDefinition): ColumnDefinition {
    this.columns.push(col);
    return col;
  }

  addColumnByName( columnName: string,isMandatory: boolean, horizontalAlignment?: string, verticalAlignment?: EnumColumnVerticalAlignment,
                    columnWidth?: number, indentLevel?: number, style?: string, numberFormat?: string, visible: boolean = true, autosizeColumn: boolean = false
                 ): ColumnDefinition {

    const col: ColumnDefinition = {
                                    columnName,
                                    isMandatory,
                                    horizontalAlignment,
                                    verticalAlignment,
                                    columnWidth,
                                    indentLevel,
                                    style,
                                    numberFormat,
                                    visible,
                                    autosizeColumn
                                };
    this.columns.push(col);
    return col;
  }

  convertColumnDefinitionsToJson(): string {
    // Convert the array of ColumnDefinition objects to a JSON string
    const json: string = JSON.stringify(this.columns);
    return json;
  }

}
