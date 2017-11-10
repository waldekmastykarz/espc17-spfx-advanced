declare interface ISampleCommandCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'SampleCommandCommandSetStrings' {
  const strings: ISampleCommandCommandSetStrings;
  export = strings;
}
