import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { SPHttpClient } from '@microsoft/sp-http';

export interface IWorldClockProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: ApplicationCustomizerContext;
  spHttpClient: SPHttpClient;
}
