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
  headerFontSize: number;
  contentBackgroundColor: string;
  contentTextColor: string;
  contentFontBold: boolean;
  contentFontSize: number;
  cornerStyle: string;
  cornerRadius: number;
  showCloseBar: boolean;
  closeBarText: string;
  closeBarBackgroundColor: string;
  closeBarTextColor: string;
  closeBarFontSize: number;
  closeBarFontBold: boolean;
  closeBarAlignment: string;
  fontFamily: string;
  fontStyle: string;
  overrideCssUrl: string;
  enableDiagnostics: boolean;
  }