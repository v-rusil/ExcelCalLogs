export enum enumConditionalFormatType {
    ColorScale = "ColorScale",
    ContainsText = "ConstainsText",
    CellValue = "CellValue",
    Custom ="Custom"
  }
  
  export enum enumConditionalFormatTextOperator {
    Contains = "Contains",
    NotContains ="NotContains"
  }

  export enum enumCellValueOperator{
    LT = "LessThan",
    GT ="GreaterThan",
    EQ = "Equal",
    BETWEEN="Between"
  }
  
  export interface ConditionalFormat {
    FriendlyName:string;
    ColumnName:string;
    CountOccurrences:boolean;
    WarnIfOccurrencesGTZero:boolean;
    Type:enumConditionalFormatType;
    Style:string;
    FontColor:string;
    FillColor:string;

    ColorScaleColorMinimum:string;
    ColorScaleColorMaximum:string;

    ContainsTextSearch:string;
    ContainsTextOperator:enumConditionalFormatTextOperator;

    CellValueFormula1:string;
    CellValueFormula2:string;
    CellValueOperator:enumCellValueOperator;

    CustomFormula:string;
  }
  

  function CellValueOperatorToJsonEnum(op:Excel.ConditionalCellValueOperator) : enumCellValueOperator
  {
      var retval:enumCellValueOperator = enumCellValueOperator.EQ;
      switch (op) {
          case Excel.ConditionalCellValueOperator.equalTo:
            retval = enumCellValueOperator.EQ;
            break;
        
          case Excel.ConditionalCellValueOperator.between:
            retval = enumCellValueOperator.BETWEEN;
            break;

          case Excel.ConditionalCellValueOperator.greaterThan:
            retval = enumCellValueOperator.GT;
            break;
      

          case Excel.ConditionalCellValueOperator.lessThan:
            retval = enumCellValueOperator.LT;
            break;

      default:
          retval = enumCellValueOperator.EQ;
          break;
      }

      return retval;
  }


