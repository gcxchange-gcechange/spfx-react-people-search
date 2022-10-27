declare interface IReactDirectoryWebPartStrings {
  DropDownPlaceLabelMessage: string;
  DropDownPlaceHolderMessage: string;
  SearchPlaceHolder: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  DirectoryMessage: string;
  LoadingText: string;
  SearchPropsLabel: string;
  SearchPropsDesc: string;
  ClearTextSearchPropsLabel: string;
  ClearTextSearchPropsDesc: string;
  PagingLabel: string;
  FirstName: string;
  LastName: string;
}

declare module 'ReactDirectoryWebPartStrings' {
  const strings: IReactDirectoryWebPartStrings;
  export = strings;
}
