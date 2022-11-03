declare interface IReactDirectoryWebPartStrings {
  SearchPlaceHolder: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string; 
  PagingLabel: string;
  DirectoryMessage: string;
  LoadingText: string;
}

declare module 'ReactDirectoryWebPartStrings' {
  const strings: IReactDirectoryWebPartStrings;
  export = strings;
}
