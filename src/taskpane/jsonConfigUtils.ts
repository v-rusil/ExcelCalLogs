import { ColumnDefinition, EnumColumnHorizontalAlignment, EnumColumnVerticalAlignment } from './columnDefinitions';
import {ConditionalFormat, enumConditionalFormatType, enumConditionalFormatTextOperator, enumCellValueOperator} from './ConditionalFormats'
import { FilterDefinition } from './FilterDefinitions';


export class JsonConfigUtils 
{
  public columns: ColumnDefinition[] = [];
  public conditionalFormats: ConditionalFormat[]=[];
  public filterDefinitions: FilterDefinition[]=[];

  
  
  constructor() {
    // this.json = value;
  }
  
  getValue(): string {
    return JSON.stringify(this);
  }

  //#region Column Definitions
  addColumn(col: ColumnDefinition): ColumnDefinition {
    this.columns.push(col);
    return col;
  }

  addColumnByName( columnName: string,isMandatory: boolean, horizontalAlignment?: string, verticalAlignment?: EnumColumnVerticalAlignment,
                    columnWidth?: number, indentLevel?: number, style?: string, numberFormat?: string, visible: boolean = true, autosizeColumn: boolean = false,
                    searchFor:string="", replaceWith:string=""
                 ): ColumnDefinition {

    const col: ColumnDefinition = 

                                  { columnName,
                                    isMandatory,
                                    horizontalAlignment,
                                    verticalAlignment,
                                    columnWidth,
                                    indentLevel,
                                    style,
                                    numberFormat,
                                    visible,
                                    autosizeColumn,
                                    searchFor,
                                    replaceWith
                                };
    this.columns.push(col);
    return col;
  }

  convertColumnDefinitionsToJson(): string {
    // Convert the array of ColumnDefinition objects to a JSON string
    const json: string = JSON.stringify(this.columns);
    return json;
  }


convertToHorizontalAlignment(value: string) : Excel.HorizontalAlignment
{
  switch (value) {
    case "Center":
      return Excel.HorizontalAlignment.center;
      break;
    case "Left":
      return Excel.HorizontalAlignment.left;
      break;
    case "Right":
      return Excel.HorizontalAlignment.right;
      break;
    case "Justify":
      return Excel.HorizontalAlignment.justify
      break;
    case"General":
    return Excel.HorizontalAlignment.general
      break;
    default:
      return Excel.HorizontalAlignment.general
      break;
  }
}

convertToVerticalAlignment(value: string): Excel.VerticalAlignment {
  switch (value) {
    case "Bottom":
      return Excel.VerticalAlignment.bottom;
    case "Center":
      return Excel.VerticalAlignment.center;
    case "Distributed":
      return Excel.VerticalAlignment.distributed;
    case "Justify":
      return Excel.VerticalAlignment.justify;
    case"Top":
        return Excel.VerticalAlignment.top;
    default:
      return Excel.VerticalAlignment.top;
  }
}

//#endregion

//#region ConditionalFormats Definitions
public addConditionalFormat(cf: ConditionalFormat): void {
  this.conditionalFormats.push(cf);
}

public addConditionalFormatByName(friendlyName: string, columnName:string, countOccurrences: boolean, warnIfOccurrencesGTZero: boolean, 
                                  type: enumConditionalFormatType, style: string, fontColor: string, fillColor: string, 
                                  colorScaleColorMinimum: string, colorScaleColorMaximum: string, 
                                  containsTextSearch: string, containsTextOperator: enumConditionalFormatTextOperator, 
                                  cellValueFormula1: string, cellValueFormula2: string, cellValueOperator: enumCellValueOperator, 
                                  customFormula: string): void 
{
  const cf: ConditionalFormat = {
      FriendlyName: friendlyName,
      ColumnName:columnName,
      CountOccurrences: countOccurrences,
      WarnIfOccurrencesGTZero: warnIfOccurrencesGTZero,
      Type: type,
      Style: style,
      FontColor: fontColor,
      FillColor: fillColor,
      ColorScaleColorMinimum: colorScaleColorMinimum,
      ColorScaleColorMaximum: colorScaleColorMaximum,
      ContainsTextSearch: containsTextSearch,
      ContainsTextOperator: containsTextOperator,
      CellValueFormula1: cellValueFormula1,
      CellValueFormula2: cellValueFormula2,
      CellValueOperator: cellValueOperator,
      CustomFormula: customFormula
  };
  this.conditionalFormats.push(cf);
}


public addConditionalFormatColorScale(friendlyName: string, columnName:string, countOccurrences: boolean, warnIfOccurrencesGTZero: boolean, 
  style: string, fontColor: string, fillColor: string, 
  colorScaleColorMinimum: string, colorScaleColorMaximum: string ): void 
{
  const cf: ConditionalFormat = {
    FriendlyName: friendlyName,
    ColumnName: columnName,
    CountOccurrences: countOccurrences,
    WarnIfOccurrencesGTZero: warnIfOccurrencesGTZero,
    Type: enumConditionalFormatType.ColorScale,
    Style: style,
    FontColor: fontColor,
    FillColor: fillColor,
    ColorScaleColorMinimum: colorScaleColorMinimum,
    ColorScaleColorMaximum: colorScaleColorMaximum,
    ContainsTextSearch: null,
    ContainsTextOperator: null,
    CellValueFormula1: null,
    CellValueFormula2: null,
    CellValueOperator: null,
    CustomFormula: null
  };
  this.conditionalFormats.push(cf);
}


public addConditionalFormatContainsText(friendlyName: string, columnName:string, countOccurrences: boolean, warnIfOccurrencesGTZero: boolean, 
  style: string, fontColor: string, fillColor: string, 
  containsTextSearch: string, containsTextOperator: enumConditionalFormatTextOperator): void 
{
    const cf: ConditionalFormat = {
      FriendlyName: friendlyName,
      ColumnName: columnName,
      CountOccurrences: countOccurrences,
      WarnIfOccurrencesGTZero: warnIfOccurrencesGTZero,
      Type: enumConditionalFormatType.ContainsText,
      Style: style,
      FontColor: fontColor,
      FillColor: fillColor,
      ColorScaleColorMinimum: null,
      ColorScaleColorMaximum: null,
      ContainsTextSearch: containsTextSearch,
      ContainsTextOperator: containsTextOperator,
      CellValueFormula1: null,
      CellValueFormula2: null,
      CellValueOperator: null,
      CustomFormula: null
    };
    this.conditionalFormats.push(cf);
}


public addConditionalFormatCellValue(friendlyName: string, ColumnName:string, countOccurrences: boolean, warnIfOccurrencesGTZero: boolean, 
  style: string, fontColor: string, fillColor: string, 
  cellValueFormula1: string, cellValueFormula2: string, cellValueOperator: enumCellValueOperator ): void 
{
  const cf: ConditionalFormat = {
    FriendlyName: friendlyName,
    ColumnName: ColumnName,
    CountOccurrences: countOccurrences,
    WarnIfOccurrencesGTZero: warnIfOccurrencesGTZero,
    Type: enumConditionalFormatType.CellValue,
    Style: style,
    FontColor: fontColor,
    FillColor: fillColor,
    ColorScaleColorMinimum: null,
    ColorScaleColorMaximum: null,
    ContainsTextSearch: null,
    ContainsTextOperator: null,
    CellValueFormula1: cellValueFormula1,
    CellValueFormula2: cellValueFormula2,
    CellValueOperator: cellValueOperator,
    CustomFormula: null
  };
  this.conditionalFormats.push(cf);
}


public addConditionalFormatCustom(friendlyName: string, columnName:string, countOccurrences: boolean, warnIfOccurrencesGTZero: boolean, 
  style: string, fontColor: string, fillColor: string, 
  customFormula: string): void 
{
    const cf: ConditionalFormat = {
      FriendlyName: friendlyName,
      ColumnName: columnName,
      CountOccurrences: countOccurrences,
      WarnIfOccurrencesGTZero: warnIfOccurrencesGTZero,
      Type: enumConditionalFormatType.Custom,
      Style: style,
      FontColor: fontColor,
      FillColor: fillColor,
      ColorScaleColorMinimum: null,
      ColorScaleColorMaximum: null,
      ContainsTextSearch: null,
      ContainsTextOperator: null,
      CellValueFormula1: null,
      CellValueFormula2: null,
      CellValueOperator: null,
      CustomFormula: customFormula
    };
    this.conditionalFormats.push(cf);
}


//#endregion

//#region Add filter Condition 
addFilterCondition( columnName:string, value:string, key:string):FilterDefinition
{
  const newFilter:FilterDefinition = {
    FilterActiveByUIOnly:false,
    FilterKey:key,
    FilterValue:value,
    ColumnName:columnName,
    FriendlyName:`${columnName} filter ${value}`
  }
  this.filterDefinitions.push(newFilter);
  return newFilter;
}
//#endregion

}
