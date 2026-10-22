import { SPHttpClient } from '@microsoft/sp-http';
import { IRegistrationSelection } from '../models';

export interface IEventRegistrationProps {
  title: string;
  description: string;
  webUrl: string;
  spHttpClient: SPHttpClient;
  listPrefix: string;
  organizerGroupId: number;
  buttonDisplayMode: string;
  enableDiagnostics: boolean;
  initialEventId: number;
  initialSessionId: number;
  sourceUrl: string;
  onRegistrationSelected: (selection: IRegistrationSelection) => void;
}
