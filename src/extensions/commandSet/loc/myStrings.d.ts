declare interface ICommandSetCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'CommandSetCommandSetStrings' {
  const strings: ICommandSetCommandSetStrings;
  export = strings;
}
