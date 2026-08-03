//import { SPHttpClient } from "@microsoft/sp-http";
import { ISPClockList } from "./ISPClockListElement";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { SPHttpClient } from '@microsoft/sp-http';

export interface IWorldClockProps {
  //whether the clock should be only in digital format, as opposed to the default (digital & analog) 
  digitalOnly: boolean;
  //whether the clock should be in the 12hr format (12AM -> 12PM -> 12AM) or the default format (00hr -> 23hr)  
  hour12: boolean;
  //whether the clock should display the weekday 
  displayDay: boolean;
  //name of the sharepoint worldclock list
  spList: string;
  //error handler for the spfxfetching
  errorHandler: (errorMessage: string) => void;
  //http client for the sharepoint context
  context: ApplicationCustomizerContext;
  //base url of site page
  absoluteUrl: string;
  //spHttpClinet
  spHttpClient: SPHttpClient;
}

export interface IWorldClockState {
  //list of clock data from sharepoint
  spClockList: ISPClockList;
}
