declare interface IAccordionWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ListNameFieldLabel: string;
  ItemNameFieldLabel: string;
  ItemContentFieldLabel: string;
  OptionChoiceFieldLabel: string;
}

declare module 'AccordionWebPartStrings' {
  const strings: IAccordionWebPartStrings;
  export = strings;
}