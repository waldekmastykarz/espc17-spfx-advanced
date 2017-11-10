declare interface IHelloWorldBlankWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'HelloWorldBlankWebPartStrings' {
  const strings: IHelloWorldBlankWebPartStrings;
  export = strings;
}
