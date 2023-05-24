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
        
        return JSON.stringify(jsonArray);
    } catch (error) {
        console.log(error);
        myConsole.log(error);
        return "";
    }
}