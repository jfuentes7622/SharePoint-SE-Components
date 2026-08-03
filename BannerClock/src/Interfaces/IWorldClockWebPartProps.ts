export interface IWorldClockWebPartProps {
  //whether the clock should be only in digital format, as opposed to the default (digital & analog)
  digitalOnly: boolean;
  //whether the clock should be in the 12hr format (12AM -> 12PM -> 12AM) or the default format (00hr -> 23hr
  hour12: boolean;
  //whether the clock should display the weekday 
  displayDay: boolean;
  //root path of site collection to search for lists
  collectionRoot: string;
  //name of the sharepoint worldclock list
  spList: string;

}
