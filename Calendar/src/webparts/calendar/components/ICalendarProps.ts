import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICalendarProps {
  title: string;
  displayMode: DisplayMode;
  spfxContext: WebPartContext;
  listName: string;
  defaultView: string;
  showWeekends: boolean;
  calendarHeight: number;
  enableDiagnostics: boolean;
  fUpdateProperty: (value: string) => void;
  fPropertyPaneOpen: () => void;
}
