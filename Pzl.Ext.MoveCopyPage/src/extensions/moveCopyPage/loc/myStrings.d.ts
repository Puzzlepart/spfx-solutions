declare interface IMoveCopyPageCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'MoveCopyPageCommandSetStrings' {
  const strings: IMoveCopyPageCommandSetStrings;
  export = strings;
}
