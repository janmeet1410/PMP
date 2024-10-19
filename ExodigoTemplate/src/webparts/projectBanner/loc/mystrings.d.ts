declare interface IProjectBannerWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'ProjectBannerWebPartStrings' {
  const strings: IProjectBannerWebPartStrings;
  export = strings;
}
