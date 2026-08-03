declare interface IWorldClockWebPartStrings {
  // Property Pane labels and strings
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  DigitalLabel: string;
  ListNameFieldLabel: string;

  // Placeholder labels and strings
  PlaceholderIconName: string;
  PlaceholderIconText: string;
  PlaceholderDescription: string;
  PlaceholderButtonLabel: string;
}

declare module "WorldClockWebPartStrings" {
  const strings: IWorldClockWebPartStrings;
  export = strings;
}
