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



  //#region  conversion utils

convertToHorizontalAlignment(value: string): Excel.HorizontalAlignment {
  switch (value) {
    case "Center":
      return Excel.HorizontalAlignment.center;
    case "Left":
      return Excel.HorizontalAlignment.left;
    case "Right":
      return Excel.HorizontalAlignment.right;
    case "Justify":
      return Excel.HorizontalAlignment.justify
    case"General":
        return Excel.HorizontalAlignment.general
    default:
      return Excel.HorizontalAlignment.general
  }
}

  convertToVerticalAlignment(value: string): Excel.VerticalAlignment {
    switch (value) {
      case "Center":
        return Excel.VerticalAlignment.bottom;
      case "Left":
        return Excel.VerticalAlignment.center;
      case "Right":
        return Excel.VerticalAlignment.distributed;
      case "Justify":
        return Excel.VerticalAlignment.justify;
      case"General":
          return Excel.VerticalAlignment.top;
      default:
        return Excel.VerticalAlignment.top;
    }
  }

  //#endregion


}
