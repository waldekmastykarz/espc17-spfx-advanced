declare interface IHelloWorldTenantWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'HelloWorldTenantWebPartStrings' {
  const strings: IHelloWorldTenantWebPartStrings;
  export = strings;
}
