import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
// import {
//     BaseComponentContext
//   } from '@microsoft/sp-component-base';

export interface ILoaderProps{
    context: ApplicationCustomizerContext;
    rootUrl: string;
}