export interface IClockProps {
    //time zone descriptor as seen in IANA time zone database
    timeZone: string;
    //whether the clock should be only in digital format, as opposed to the default (digital & analog) 
    digitalOnly: boolean;
    //whether the clock should be in the 12hr format (12AM -> 12PM -> 12AM) or the default format (00hr -> 23hr)
    hour12: boolean;
    //whether the clock should display the weekday 
    displayDay: boolean;
}
export interface IClockState {
    //the current date/time to be shown on the clock (is loosely equivalent to Date.now())
    date: Date;
}