declare interface IHelloWorldLibraryStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'HelloWorldLibraryStrings' {
  const strings: IHelloWorldLibraryStrings;
  export = strings;
}
