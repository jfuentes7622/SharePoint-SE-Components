import * as React from 'react';
import styles from './Calendar.module.scss';
import { ICalendarProps } from './ICalendarProps';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import * as strings from 'CalendarWebPartStrings';

function escapeODataText(value: string): string {
  return value.replace(/'/g, "''");
}

export interface ICalendarState {
  loading: boolean;
  errorMessage: string;
}

export default class Calendar extends React.Component<ICalendarProps, ICalendarState> {
  private _calendarEl: HTMLDivElement;
  private _calendar: any;

  constructor(props: ICalendarProps) {
    super(props);
    this.state = { loading: false, errorMessage: '' };
  }

  public componentWillUnmount(): void {
    this.logDiagnostic('componentWillUnmount called.');
    if (this._calendar) {
      this._calendar.destroy();
      this._calendar = undefined;
    }
  }

  public componentDidUpdate(prevProps: ICalendarProps): void {
    if (!this._calendar) {
      return;
    }

    if (prevProps.listName !== this.props.listName) {
      this.loadEvents();
    }
    if (prevProps.defaultView !== this.props.defaultView) {
      this._calendar.changeView(this.props.defaultView || 'dayGridMonth');
    }
    if (prevProps.showWeekends !== this.props.showWeekends) {
      this._calendar.setOption('weekends', this.props.showWeekends !== false);
    }
    if (prevProps.calendarHeight !== this.props.calendarHeight) {
      this._calendar.setOption('height', this.props.calendarHeight || 600);
    }
  }

  private logDiagnostic(message: string): void {
    if (this.props.enableDiagnostics === false) {
      return;
    }

    console.log('[Calendar] ' + message);
  }

  private _setCalendarRef = (el: HTMLDivElement): void => {
    if (!el && this._calendar) {
      this._calendar.destroy();
      this._calendar = undefined;
    }

    this._calendarEl = el;

    if (el && !this._calendar) {
      this._initCalendar();
    }
  }

  private _initCalendar(): void {
    if (!this._calendarEl) {
      return;
    }

    this.logDiagnostic('Initializing FullCalendar. defaultView=' + this.props.defaultView);

    // Loaded via require() rather than import so TypeScript never has to parse
    // FullCalendar's own .d.ts files (they target a newer TypeScript than this project uses).
    const FullCalendarCore = require('@fullcalendar/core');
    const dayGridModule = require('@fullcalendar/daygrid');
    const timeGridModule = require('@fullcalendar/timegrid');
    const listModule = require('@fullcalendar/list');
    const interactionModule = require('@fullcalendar/interaction');

    const dayGridPlugin = dayGridModule.default || dayGridModule;
    const timeGridPlugin = timeGridModule.default || timeGridModule;
    const listPlugin = listModule.default || listModule;
    const interactionPlugin = interactionModule.default || interactionModule;

    this._calendar = new FullCalendarCore.Calendar(this._calendarEl, {
      plugins: [dayGridPlugin, timeGridPlugin, listPlugin, interactionPlugin],
      initialView: this.props.defaultView || 'dayGridMonth',
      height: this.props.calendarHeight || 600,
      weekends: this.props.showWeekends !== false,
      headerToolbar: {
        left: 'prev,next today',
        center: 'title',
        right: 'dayGridMonth,timeGridWeek,timeGridDay,listWeek'
      },
      eventDidMount: (info: any) => {
        const location = info.event.extendedProps && info.event.extendedProps.location;
        info.el.title = info.event.title + (location ? ' \u2014 ' + location : '');
      }
    });

    this._calendar.render();
    this.loadEvents();
  }

  private async loadEvents(): Promise<void> {
    if (!this._calendar) {
      return;
    }

    if (!this.props.listName) {
      this._calendar.removeAllEventSources();
      return;
    }

    this.setState({ loading: true, errorMessage: '' });

    try {
      this.logDiagnostic('Loading events from list "' + this.props.listName + '"');
      const webUrl = this.props.spfxContext.pageContext.web.absoluteUrl.replace(/\/$/, '');
      const url = webUrl + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items"
        + '?$select=Id,Title,EventDate,EndDate,fAllDayEvent,Location&$orderby=EventDate&$top=500';

      let response = await this.props.spfxContext.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) {
        response = await this.props.spfxContext.spHttpClient.get(url, SPHttpClient.configurations.v1, {
          headers: { Accept: 'application/json;odata=nometadata' }
        });
      }
      if (!response.ok) {
        throw new Error('Request failed. HTTP ' + String(response.status) + ' ' + response.statusText);
      }

      const data = await response.json();
      const items = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      const events = items.map((item: any) => {
        return {
          id: String(item.Id),
          title: item.Title || '',
          start: item.EventDate,
          end: item.EndDate,
          allDay: !!item.fAllDayEvent,
          extendedProps: { location: item.Location || '' }
        };
      });

      this.logDiagnostic('Loaded events successfully. Count=' + String(events.length));
      this._calendar.removeAllEventSources();
      this._calendar.addEventSource(events);
      this.setState({ loading: false, errorMessage: '' });
    } catch (error) {
      const message = error && error.message ? error.message : String(error);
      console.error('[Calendar] Failed to load events from list "' + this.props.listName + '": ' + message);
      this.setState({ loading: false, errorMessage: strings.LoadErrorMessage + ' ' + message });
    }
  }

  public render(): React.ReactElement<ICalendarProps> {
    this.logDiagnostic('render() called. listName=' + (this.props.listName || '(none)'));
    return (
      <div className={styles.calendar}>
        <WebPartTitle displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.fUpdateProperty} />

        {
          this.props.listName ? (
            <div>
              {this.state.errorMessage && <div className={styles.error}>{this.state.errorMessage}</div>}
              {this.state.loading && <div className={styles.message}>{strings.LoadingMessage}</div>}
              <div className={styles.calendarContainer} ref={this._setCalendarRef} />
            </div>
          ) : (
            <Placeholder
              iconName='Calendar'
              iconText={strings.NoListIconText}
              description={strings.NoListConfigured}
              buttonLabel={strings.NoListBtn}
              onConfigure={this.props.fPropertyPaneOpen} />
          )
        }
      </div>
    );
  }
}
