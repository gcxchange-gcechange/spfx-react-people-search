declare interface IReactDirectoryWebPartStrings {
  SearchPlaceHolder: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  PagingLabel: string;
  DirectoryMessage: string;
  LoadingText: string;
  SearchBoxLabel: string;
  SendEmailLabel:string;
  NoUserFoundLabelText: string;
  NoUserFoundImageAltText: string;
  NoUserFoundEmailSubject:string;
  NoUserFoundEmailBody:string;
  NoUserFoundEmail:string;


}

declare module 'ReactDirectoryWebPartStrings' {
  const strings: IReactDirectoryWebPartStrings;
  export = strings;
}
