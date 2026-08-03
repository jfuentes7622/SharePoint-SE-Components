/**
 * Custom interface definition: as of 04/12/2021 Typescript does not have support for the 'hourCycle' option for DateTimeFormatOptions; 
 * this option is needed in this project, so we have added it as shown below. 
 * The solution in 'Clock.tsx' utilizes this definition rather than Typescripts built in one.
 */

export interface IdateTimeFormatOptions {
    localeMatcher?: "best fit" | "lookup";
    weekday?: "long" | "short" | "narrow";
    era?: "long" | "short" | "narrow";
    year?: "numeric" | "2-digit";
    month?: "numeric" | "2-digit" | "long" | "short" | "narrow";
    day?: "numeric" | "2-digit";
    hour?: "numeric" | "2-digit";
    minute?: "numeric" | "2-digit";
    second?: "numeric" | "2-digit";
    timeZoneName?: "long" | "short";
    formatMatcher?: "best fit" | "basic";
    //***************************************/
    hourCycle?: 'h11' | 'h12' | 'h23' | 'h24';
    //***************************************/
    hour12?: boolean;
    timeZone?: string;
}