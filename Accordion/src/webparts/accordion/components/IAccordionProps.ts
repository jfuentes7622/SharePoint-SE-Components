import {  WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAccordionProps {
  listName: string;
  itemName: string;
  itemContent: string;
  optionChoice: string;
  spfxContext: WebPartContext;
  headerBackgroundColor: string;
  headerTextColor: string;
  headerFontBold: boolean;
  contentBackgroundColor: string;
  contentTextColor: string;
  contentFontBold: boolean;
  fontFamily: string;
  fontStyle: string;
  overrideCssUrl: string;
  enableDiagnostics: boolean;
  }