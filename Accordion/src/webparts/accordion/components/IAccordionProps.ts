import {  WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAccordionProps {
  listName: string;
  itemName: string;
  itemContent: string;
  optionChoice: string;
  spfxContext: WebPartContext;
  }