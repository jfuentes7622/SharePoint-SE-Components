declare interface ICalendarWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;

  DataSourceGroupName: string;
  ListNameFieldLabel: string;
  DefaultViewFieldLabel: string;
  ViewMonthLabel: string;
  ViewWeekLabel: string;
  ViewDayLabel: string;
  ViewListLabel: string;
  ShowWeekendsFieldLabel: string;
  CalendarHeightFieldLabel: string;

  DiagnosticsGroupName: string;
  PropEnableDiagnosticsLabel: string;

  NoListIconText: string;
  NoListConfigured: string;
  NoListBtn: string;
  LoadingMessage: string;
  LoadErrorMessage: string;
}

declare module 'CalendarWebPartStrings' {
  const strings: ICalendarWebPartStrings;
  export = strings;
}
