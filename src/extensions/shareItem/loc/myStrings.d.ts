declare interface IShareItemCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ShareItemCommandSetStrings' {
  const strings: IShareItemCommandSetStrings;
  export = strings;
}
