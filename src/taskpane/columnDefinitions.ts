export enum EnumColumnHorizontalAlignment {
    Top = 'Top',
    Middle = 'Middle',
    Bottom = 'Bottom'
  }
  
  export enum EnumColumnVerticalAlignment {
    Left = 'Left',
    Middle = 'Middle',
    Right = 'Right'
  }
  
  export interface ColumnDefinition {
    columnName: string;
    isMandatory: boolean;
    horizontalAlignment?: string;
    verticalAlignment?: string;
    columnWidth?: number;
    indentLevel?: number;
    style?: string;
    numberFormat: string;
    visible: boolean;
    autosizeColumn?: boolean;
  }
  