declare interface IMyCommandSetCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'MyCommandSetCommandSetStrings' {
  const strings: IMyCommandSetCommandSetStrings;
  export = strings;
}
