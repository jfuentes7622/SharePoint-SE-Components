declare interface ITabComponentWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  LayoutGroupName: string;
  ManageTabs: string;
  Tabs: string;
  DisableColor: string;
  useGlobalCSS:strings;
  OverrideCSS: string;
  SelectedColor: string;
  ErrorClassicSharePoint: string;
  TabType:string;
  DiagnosticsGroupName: string;
  PropEnableDiagnosticsLabel: string;
}

declare module 'TabComponentWebPartStrings' {
  const strings: ITabComponentWebPartStrings;
  export = strings;
}
