import { myConsole } from "./myConsole";

export async function tblToJson(ctx:Excel.RequestContext, tbl:Excel.Table) : Promise<string>
{
    try {
        // Read the table data
        tbl.load(["name", "address","values"]);
        
        await ctx.sync();

        const tableR = tbl.getDataBodyRange();
        tableR.load("values"); 
        await ctx.sync();
        // Convert table data to a JavaScript array
        const tableData = tableR.values;

        const tableH = tbl.getHeaderRowRange();
        tableH.load("values");
        await ctx.sync();

        // Convert the table data to a JSON array
        const jsonArray = tableData.map(row => {
        const jsonObject = {};
        
        tableH.values[0].forEach((header, index) => {
                jsonObject[header] = row[index];
            });
            return jsonObject;
        });

        // Convert the JSON array to a JSON string
        
        return JSON.stringify(jsonArray, null, 2);
    } catch (error) {
        console.log(error);
        myConsole.log(error);
        return "";
    }
}


export async function getTableCellRange(ctx:Excel.RequestContext, tbl:Excel.Table, rowIndex: number, columnName: string): Promise<Excel.Range | undefined> 
{
    let cellRange: Excel.Range | undefined = undefined;
  
    try 
    {
        myConsole.log(`Getting cell range for row ${rowIndex} in column ${columnName}`);
        const col:Excel.TableColumn = tbl.columns.getItemOrNullObject(columnName);
        col.load(["index","name"]);
        await ctx.sync();
        const colIndex:number = col.index;
        cellRange = col.getRange().getCell(rowIndex, colIndex);
    }    
    catch(error) 
    {
      myConsole.log(`Error getting table cell range: ${error}`);
    }
  
    return cellRange;
  }
  
  export async function tryReadTblValue(ctx:Excel.RequestContext, tbl:Excel.Table, rowIndex: number, columnName: string): Promise<string> 
  {
    let result: string = "";
    try 
    {
      myConsole.log(`Reading value for row ${rowIndex} in column ${columnName}`);
      const cellRange:Excel.Range | undefined = await getTableCellRange(ctx, tbl, rowIndex, columnName);
      if (cellRange !== undefined) 
      {
        cellRange.load(["values"]);
        await ctx.sync();
        result = cellRange.values[0][0].toString();
      }
    } 
    catch (error) 
    {
      myConsole.log(`Error tryReadTblValue: ${error}`);
    }
    return result;
  }

  export interface IPreviousRowChanged 
  {
    isChanged: boolean;
    currentValue: string;
    previousValue: string;
  }

    export async function isPreviousRowChanged(ctx:Excel.RequestContext, tbl:Excel.Table, rowIndex: number, previousRowIndex:number, columnName: string): Promise<IPreviousRowChanged | undefined>
    {
        let result: IPreviousRowChanged | undefined = undefined;
        try 
        {
            myConsole.log(`Previous row changed? ${rowIndex} in column ${columnName}`);
            const cellCurrentRowValue:string = await tryReadTblValue(ctx, tbl, rowIndex, columnName);
            const cellPreviousRowValue:string = await tryReadTblValue(ctx, tbl, previousRowIndex, columnName);
            result = {  
                isChanged: cellCurrentRowValue !== cellPreviousRowValue,
                currentValue: cellCurrentRowValue,
                previousValue: cellPreviousRowValue
            };
        }
        catch (error) 
        {
            myConsole.log(`Error isPreviousRowChanged: ${error}`);
        }
        return result;
    }
        
        

