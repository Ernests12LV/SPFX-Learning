export class HelloWorldLibrary {
  public name(): string {
    return 'HelloWorldLibrary';
  }

  public getCurrentTime(): string {
    let currentDate: Date;
    let str: string;

    currentDate = new Date();

    str="<br>Todays date is: " + currentDate.toLocaleDateString() + "<br>";
    str += "Current time is: " + currentDate.toLocaleTimeString() + "<br>";
    
    return(str);
  }

}