export class myConsole
{

static count:number = 0;

  constructor() {
    myConsole.count = 0;
  }

  public static addCounter(): void {
    myConsole.count++;
  }

  public static log(message:string):void 
  {
    console.log(message);
    myConsole.addCounter();
    myConsole.addRow(myConsole.count.toString(), message);
  }

    private static addRow(first: string, second: string): void {
        const myConsole = document.querySelector('#myConsole');
        const newRow = document.createElement('div');
        newRow.classList.add('row');
        const firstCell = document.createElement('div');
        firstCell.classList.add('col-2');
        firstCell.innerHTML = `<small> ${first}</small>`;
        const secondCell = document.createElement('div');
        secondCell.classList.add('col-10');
        secondCell.innerHTML = `<small>${second}</small>`;
        newRow.appendChild(firstCell);
        newRow.appendChild(secondCell);
        myConsole.appendChild(newRow);
      }
     
      public static reset()
      {
        myConsole.count=0;
        const consoleDiv = document.querySelector('#myConsole');
        consoleDiv.innerHTML="";

      }
}