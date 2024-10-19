declare interface IContactCsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'ContactCsWebPartStrings' {
  const strings: IContactCsWebPartStrings;
  export = strings;
}
