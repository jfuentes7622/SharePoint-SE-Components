import {DisplayMode} from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IKmMarqueeProps {
  description: string;
  dMode: DisplayMode;
  spfxContext:WebPartContext;
  marqueeTextColor: string;
  marqueeBackColor: string;
  marqueeActive:boolean;
}

