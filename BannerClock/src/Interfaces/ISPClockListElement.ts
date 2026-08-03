export interface ISPClockList {
  //list of clock data from sharepoint
  value: Array<ISPClockListElement>;
}
export interface ISPClockListElement {
  //the value of the Title column in the sharepoint worldclock list 
  Title: string;
  //the value of the Timezone column in the sharepoint worldclock list (IANA time zone descriptor)
  Timezone: string;
}
