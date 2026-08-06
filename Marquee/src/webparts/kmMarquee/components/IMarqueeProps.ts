import {DisplayMode} from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IKmMarqueeProps {
  description: string;
  dMode: DisplayMode;
  spfxContext:WebPartContext;
  marqueeTextColor: string;
  marqueeBackColor: string;
  marqueeActive:boolean;
  fontFamily: string;
  fontSize: string;
  fontStyle: string;
  fontBold: boolean;
  marqueeHeight: string;
  scrollSpeed: number;
  scrollDirection: string;
  listName: string;
  messageField: string;
  messageDuration: number;
  enableDiagnostics: boolean;
}

