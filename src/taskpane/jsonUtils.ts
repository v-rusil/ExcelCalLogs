const ctx:Excel.RequestContext = require('./taskpane.ts');



async function testJsonUtil(): Promise<boolean>
{
    await ctx.sync();
    return true;
}