import * as React from 'react';
import styles from './Calendar.module.scss';
import { ICalendarDataSource, ICalendarProps } from './ICalendarProps';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import * as strings from 'CalendarWebPartStrings';

const RRuleEngine: any = require('rrule/dist/es5/rrule').RRule;
const DOMPurify: any = require('dompurify');

function escapeODataText(value: string): string {
  return value.replace(/'/g, "''");
}

interface ICalendarCondition {
  sourceListName?: string;
  field: string;
  operator: string;
  logical: string;
  valueType?: string;
  value: any;
}

interface ICalendarStyleRule {
  enabled: boolean;
  priority: number;
  conditions: ICalendarCondition[];
  style: any;
}

interface ICalendarOccurrence {
  id: string;
  item: any;
  start: Date;
  end?: Date;
}

function appendQueryParam(url: string, key: string, value: string): string {
  var hash = '';
  var base = url;
  var hashIndex = url.indexOf('#');
  if (hashIndex >= 0) {
    base = url.substring(0, hashIndex);
    hash = url.substring(hashIndex);
  }

  return base + (base.indexOf('?') >= 0 ? '&' : '?')
    + encodeURIComponent(key) + '=' + encodeURIComponent(value) + hash;
}

function getQueryParam(url: string, paramName: string): string {
  var queryIndex = url.indexOf('?');
  if (queryIndex < 0) {
    return '';
  }

  var query = url.substring(queryIndex + 1).split('#')[0];
  var parts = query.split('&');
  for (var i = 0; i < parts.length; i += 1) {
    var pair = parts[i].split('=');
    if (decodeURIComponent(String(pair[0] || '')).toLowerCase() === paramName.toLowerCase()) {
      return decodeURIComponent(String(pair[1] || ''));
    }
  }
  return '';
}

function resolveEmbeddedPageUrl(url: string): string {
  var source = getQueryParam(url, 'source');
  var id = getQueryParam(url, 'id');
  return source && /\.aspx/i.test(source) ? source : (id && /\.aspx/i.test(id) ? id : '');
}

export interface ICalendarState {
  loading: boolean;
  errorMessage: string;
  events: any[];
  laneStartDate: Date;
  selectedEvent: any;
  selectedItemId: number;
  selectedEventKey: string;
  legendOpen: boolean;
}

export default class Calendar extends React.Component<ICalendarProps, ICalendarState> {
  private _calendarEl: HTMLDivElement;
  private _calendar: any;
  private _eventLoadSequence: number = 0;

  constructor(props: ICalendarProps) {
    super(props);
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    this.state = {
      loading: false,
      errorMessage: '',
      events: [],
      laneStartDate: today,
      selectedEvent: undefined,
      selectedItemId: 0,
      selectedEventKey: '',
      legendOpen: false
    };
  }

  public componentDidMount(): void {
    if (typeof document !== 'undefined') {
      document.addEventListener('keydown', this.handleDocumentKeyDown);
    }
  }

  public componentWillUnmount(): void {
    this.logDiagnostic('componentWillUnmount called.');
    if (typeof document !== 'undefined') {
      document.removeEventListener('keydown', this.handleDocumentKeyDown);
    }
    this._eventLoadSequence += 1;
    if (this._calendar) {
      this._calendar.destroy();
      this._calendar = undefined;
    }
  }

  public componentDidUpdate(prevProps: ICalendarProps): void {
    if (!this._calendar) {
      return;
    }

    var shouldReloadEvents = prevProps.listName !== this.props.listName
      || JSON.stringify(prevProps.dataSources || []) !== JSON.stringify(this.props.dataSources || []);
    if (shouldReloadEvents && this.state.selectedItemId > 0) {
      this.setState({ selectedEvent: undefined, selectedItemId: 0, selectedEventKey: '' });
    }
    if (prevProps.defaultView !== this.props.defaultView) {
      this._calendar.changeView(this.props.defaultView || 'dayGridMonth');
    }
    if (prevProps.showWeekends !== this.props.showWeekends) {
      this._calendar.setOption('weekends', this.props.showWeekends !== false);
      shouldReloadEvents = true;
    }
    if (prevProps.calendarHeight !== this.props.calendarHeight) {
      this._calendar.updateSize();
    }
    if (prevProps.enableSwimlanes !== this.props.enableSwimlanes && this._calendar.updateSize) {
      this._calendar.updateSize();
      shouldReloadEvents = true;
    }
    if (prevProps.swimlaneDays !== this.props.swimlaneDays) {
      shouldReloadEvents = true;
    }
    if (prevProps.filterJson !== this.props.filterJson
      || prevProps.conditionalStyleJson !== this.props.conditionalStyleJson
      || prevProps.eventBackgroundColor !== this.props.eventBackgroundColor
      || prevProps.eventTextColor !== this.props.eventTextColor
      || prevProps.eventBorderColor !== this.props.eventBorderColor) {
      shouldReloadEvents = true;
    }
    if (prevProps.selectedEventBackgroundColor !== this.props.selectedEventBackgroundColor
      || prevProps.selectedEventTextColor !== this.props.selectedEventTextColor
      || prevProps.selectedEventBorderColor !== this.props.selectedEventBorderColor
      || prevProps.selectedEventFontFamily !== this.props.selectedEventFontFamily
      || prevProps.selectedEventFontSize !== this.props.selectedEventFontSize
      || prevProps.selectedEventFontStyle !== this.props.selectedEventFontStyle
      || prevProps.selectedEventFontBold !== this.props.selectedEventFontBold
      || prevProps.selectedEventBorderWidth !== this.props.selectedEventBorderWidth
      || prevProps.selectedEventCornerRadius !== this.props.selectedEventCornerRadius) {
      this.rerenderSelectedEvents();
    }
    if (shouldReloadEvents) {
      this.loadEvents();
    }
  }

  private handleDocumentKeyDown = (keyboardEvent: KeyboardEvent): void => {
    if (keyboardEvent.key === 'Escape' && this.state.selectedEvent) {
      this.closeEventDetails();
    } else if (keyboardEvent.key === 'Escape' && this.state.legendOpen) {
      this.setState({ legendOpen: false });
    }
  }

  private logDiagnostic(message: string): void {
    if (this.props.enableDiagnostics === false) {
      return;
    }

    console.log('[Calendar] ' + message);
  }

  private logRestItems(source: string, requestUrl: string, items: any[]): void {
    if (this.props.enableDiagnostics === false) {
      return;
    }

    console.log('[Calendar] REST response.', {
      source: source,
      requestUrl: requestUrl,
      itemCount: items.length
    });
    items.forEach((item: any, index: number) => {
      var snapshot = item;
      try {
        snapshot = JSON.parse(JSON.stringify(item));
      } catch (_error) {
        // The REST item is still logged if a custom field cannot be serialized.
      }
      console.log('[Calendar] REST item.', {
        source: source,
        index: index,
        id: item && item.Id,
        title: this.getItemTitle(item),
        item: snapshot
      });
    });
  }

  private parseRestItemResponse(responseText: string): any {
    var value = String(responseText || '').trim();
    if (!value) {
      return undefined;
    }
    if (value.charAt(0) !== '<') {
      var jsonValue = JSON.parse(value);
      return jsonValue && jsonValue.d ? jsonValue.d : jsonValue;
    }
    if (typeof DOMParser === 'undefined') {
      return undefined;
    }

    var documentValue = new DOMParser().parseFromString(value, 'text/xml');
    if (documentValue.getElementsByTagName('parsererror').length > 0) {
      return undefined;
    }
    var item: any = {};
    ['Title', 'EventDate', 'EndDate', 'RecurrenceData', 'UID'].forEach((fieldName: string) => {
      var elements = documentValue.getElementsByTagNameNS
        ? documentValue.getElementsByTagNameNS('*', fieldName)
        : documentValue.getElementsByTagName('d:' + fieldName);
      if (elements.length > 0) {
        item[fieldName] = String(elements[0].textContent || '');
      }
    });
    return item;
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
      height: 'auto',
      weekends: this.props.showWeekends !== false,
      fixedWeekCount: true,
      showNonCurrentDates: true,
      expandRows: true,
      dayMaxEvents: false,
      displayEventTime: false,
      headerToolbar: {
        left: 'prev,next today',
        center: 'title',
        right: 'dayGridMonth,timeGridWeek,timeGridDay,listWeek'
      },
      datesSet: () => {
        if (this.props.enableSwimlanes !== true) {
          this.loadEvents();
        }
      },
      eventDidMount: (info: any) => {
        const location = info.event.extendedProps && info.event.extendedProps.location;
        const eventStyle = info.event.extendedProps && info.event.extendedProps.eventStyle;
        const sourceItemId = Number(info.event.extendedProps && info.event.extendedProps.sourceItemId || 0);
        const sourceItemKey = String(info.event.extendedProps && info.event.extendedProps.sourceItemKey || info.event.id || '');
        info.el.title = info.event.title + (location ? ' \u2014 ' + location : '');
        info.el.setAttribute('data-calendar-event-id', String(info.event.id || ''));
        this.applyEventElementStyle(info.el, eventStyle, sourceItemId, sourceItemKey);
        this.logDiagnostic('Mounted event. id=' + String(info.event.id) + ', title="' + String(info.event.title)
          + '", start=' + String(info.event.start ? info.event.start.toISOString() : '(none)')
          + ', elementText="' + String(info.el.textContent || '').trim() + '"');
      },
      eventsSet: (calendarEvents: any[]) => {
        this.logDiagnostic('FullCalendar accepted events. Count=' + String(calendarEvents.length));
      },
      eventClick: (info: any) => {
        const itemId = Number(info.event.extendedProps && info.event.extendedProps.sourceItemId || 0);
        this.handleEventActivated(itemId, {
          title: info.event.title,
          start: info.event.start,
          end: info.event.end,
          allDay: info.event.allDay,
          extendedProps: info.event.extendedProps
        });
      }
    });

    this._calendar.render();
  }

  private getWebUrl(): string {
    return String(this.props.spfxContext.pageContext.web.absoluteUrl || '').replace(/\/$/, '');
  }

  private handleEventActivated(itemId: number, event?: any): void {
    if (isNaN(itemId) || itemId <= 0) {
      return;
    }

    this.logDiagnostic('Event selected. itemId=' + String(itemId));
    this.props.onEventSelectionChange(itemId, 'view');
    if (this.props.enableEventDetails === true && event) {
      var eventKey = String(event.extendedProps && event.extendedProps.sourceItemKey || '');
      this.setState({ selectedEvent: event, selectedItemId: itemId, selectedEventKey: eventKey }, () => this.rerenderSelectedEvents());
    } else if (this.props.showLinkToItem === true) {
      this.setState({ selectedItemId: itemId, selectedEventKey: String(event && event.extendedProps && event.extendedProps.sourceItemKey || '') });
      var itemUrl = this.getItemLinkUrl(itemId, event);
      if (itemUrl && typeof window !== 'undefined') {
        window.location.href = itemUrl;
      }
    } else {
      this.setState({ selectedItemId: itemId, selectedEventKey: String(event && event.extendedProps && event.extendedProps.sourceItemKey || '') }, () => this.rerenderSelectedEvents());
    }
  }

  private rerenderSelectedEvents(): void {
    if (!this._calendarEl || !this._calendar || !this._calendar.getEventById) {
      return;
    }

    var eventElements = this._calendarEl.querySelectorAll('[data-calendar-event-id]');
    for (var index = 0; index < eventElements.length; index += 1) {
      var eventElement = eventElements[index] as HTMLElement;
      var eventId = String(eventElement.getAttribute('data-calendar-event-id') || '');
      var calendarEvent = this._calendar.getEventById(eventId);
      if (!calendarEvent) {
        continue;
      }
      var extendedProps = calendarEvent.extendedProps || {};
      var sourceItemId = Number(extendedProps.sourceItemId || calendarEvent.id || 0);
      var sourceItemKey = String(extendedProps.sourceItemKey || calendarEvent.id || '');
      this.applyEventElementStyle(eventElement, extendedProps.eventStyle || {}, sourceItemId, sourceItemKey);
    }
  }

  private applyEventElementStyle(element: HTMLElement, eventStyle: any, sourceItemId: number, sourceItemKey?: string): void {
    var isSelected = sourceItemId > 0 && sourceItemId === this.state.selectedItemId
      && (!this.state.selectedEventKey || String(sourceItemKey || '') === this.state.selectedEventKey);
    var backgroundColor = isSelected ? this.props.selectedEventBackgroundColor
      : (eventStyle.backgroundColor || this.props.eventBackgroundColor);
    var borderColor = isSelected ? this.props.selectedEventBorderColor
      : (eventStyle.borderColor || this.props.eventBorderColor);
    var textColor = isSelected ? this.props.selectedEventTextColor
      : (eventStyle.textColor || this.props.eventTextColor);
    element.style.setProperty('background-color', backgroundColor || (isSelected ? '#ffb900' : '#3788d8'), 'important');
    element.style.setProperty('border-color', borderColor || (isSelected ? '#8a4b00' : '#2c6faa'), 'important');
    element.style.setProperty('color', textColor || (isSelected ? '#201f1e' : '#ffffff'), 'important');
    element.style.setProperty('border-width', String(isSelected ? this.props.selectedEventBorderWidth : this.props.eventBorderWidth) + 'px', 'important');
    element.style.setProperty('border-radius', String(isSelected ? this.props.selectedEventCornerRadius : this.props.eventCornerRadius) + 'px', 'important');
    element.style.fontFamily = isSelected ? (this.props.selectedEventFontFamily || 'inherit')
      : (eventStyle.fontFamily || this.props.eventFontFamily || 'inherit');
    element.style.fontSize = isSelected ? (this.props.selectedEventFontSize || 'inherit')
      : (eventStyle.fontSize || this.props.eventFontSize || 'inherit');
    element.style.fontStyle = isSelected ? (this.props.selectedEventFontStyle || 'normal')
      : (eventStyle.fontStyle || this.props.eventFontStyle || 'normal');
    element.style.fontWeight = isSelected ? (this.props.selectedEventFontBold !== false ? 'bold' : 'normal')
      : (eventStyle.fontWeight || (this.props.eventFontBold === true ? 'bold' : 'normal'));
    if (isSelected) {
      element.setAttribute('aria-selected', 'true');
    } else {
      element.removeAttribute('aria-selected');
    }
  }

  private closeEventDetails(): void {
    this.setState({ selectedEvent: undefined }, () => this.rerenderSelectedEvents());
  }

  private getEventDetailFields(event: any): any[] {
    var defaultFields = [
      { field: 'Title', label: 'Title' },
      { field: 'EventDate', label: 'Start time' },
      { field: 'EndDate', label: 'End time' },
      { field: 'Location', label: 'Location' },
      { field: 'Description', label: 'Description' },
      { field: 'Category', label: 'Category' }
    ];
    var configuredFields: any[] = [];
    if (Array.isArray(this.props.eventDetailsFields) && this.props.eventDetailsFields.length > 0) {
      configuredFields = this.props.eventDetailsFields;
    } else {
      var source = String(this.props.eventDetailsFieldsJson || '').trim();
      if (source) {
        try {
          var parsed = JSON.parse(source);
          configuredFields = Array.isArray(parsed) ? parsed : [];
        } catch (_error) {
          configuredFields = [];
        }
      }
    }
    if (configuredFields.length === 0) {
      return defaultFields;
    }
    var eventSource = String(event && event.extendedProps && event.extendedProps.sourceListName
      || this.props.listName || '').toLowerCase();
    var primarySource = this.props.dataSources && this.props.dataSources.length > 0
      ? String(this.props.dataSources[0].listName || this.props.listName || '') : String(this.props.listName || '');
    return configuredFields.map((entry: any) => {
      var selection = entry && entry.field && typeof entry.field === 'object' ? entry.field : undefined;
      var normalizedEntry: any = {};
      Object.keys(entry || {}).forEach((key: string) => {
        normalizedEntry[key] = entry[key];
      });
      normalizedEntry.field = String(selection && selection.field || entry && entry.field || '');
      normalizedEntry.sourceListName = String(selection && selection.sourceListName
        || entry && entry.sourceListName || primarySource);
      return normalizedEntry;
    }).filter((entry: any) => entry.field
      && String(entry.sourceListName || '').toLowerCase() === eventSource);
  }

  private formatEventDetailValue(event: any, entry: any): string {
    var fieldName = String(entry && entry.field || '');
    var normalizedField = String(fieldName || '').toLowerCase();
    var value: any;
    if (normalizedField === 'title') {
      value = event.title;
    } else if (normalizedField === 'eventdate' || normalizedField === 'start') {
      value = event.start;
    } else if (normalizedField === 'enddate' || normalizedField === 'end') {
      value = event.end;
    } else {
      value = this.getItemValue(event.extendedProps && event.extendedProps.item, fieldName);
    }
    var format = String(entry && entry.format || 'auto').toLowerCase();
    var isDateValue = value instanceof Date || (value && !isNaN(Date.parse(String(value)))
      && (format === 'date' || format === 'datetime' || format === 'time' || /date|time/i.test(fieldName)));
    if (isDateValue) {
      var dateValue = value instanceof Date ? value : new Date(value);
      if (format === 'date') {
        return dateValue.toLocaleDateString();
      }
      if (format === 'time') {
        return dateValue.toLocaleTimeString();
      }
      return dateValue.toLocaleString();
    }
    if (format === 'yesno' || value === true || value === false) {
      var normalizedValue = String(value).toLowerCase();
      return value === true || value === 1 || normalizedValue === 'true' || normalizedValue === 'yes' ? 'Yes' : 'No';
    }
    return String(value === undefined || value === null || value === '' ? '\u2014' : value);
  }

  private isRichTextDetailField(entry: any): boolean {
    var fieldName = String(entry && entry.field || '').toLowerCase();
    var metadata = this.props.fieldMetadata && this.props.fieldMetadata[fieldName];
    var format = String(entry && entry.format || 'auto').toLowerCase();
    return format !== 'text' && !!metadata && metadata.type === 'note' && metadata.richText === true;
  }

  private sanitizeRichText(value: string): string {
    return DOMPurify && typeof DOMPurify.sanitize === 'function' ? DOMPurify.sanitize(value) : '';
  }

  private getRecurrenceDescription(event: any): string {
    var item = event && event.extendedProps && event.extendedProps.item;
    var recurrenceItem = this.resolveRecurrenceItem(item, event);
    if (!recurrenceItem || !this.getRecurrenceData(recurrenceItem)) {
      return '';
    }

    var rule = this.createRecurrenceRule(recurrenceItem);
    if (!rule || typeof rule.toText !== 'function') {
      return strings.RecurringEventFallback;
    }
    try {
      var description = String(rule.toText() || '').trim();
      return description ? description.charAt(0).toUpperCase() + description.substring(1) : strings.RecurringEventFallback;
    } catch (_error) {
      return strings.RecurringEventFallback;
    }
  }

  private resolveRecurrenceItem(item: any, event?: any): any {
    if (!item || this.getRecurrenceData(item)) {
      return item;
    }
    var masterId = String(item.MasterSeriesItemID || item.MasterSeriesItemId || '');
    if (!masterId) {
      return undefined;
    }
    var sourceListName = String(event && event.extendedProps && event.extendedProps.sourceListName || '').toLowerCase();
    for (var index = 0; index < this.state.events.length; index += 1) {
      var candidate = this.state.events[index];
      var candidateProps = candidate && candidate.extendedProps;
      var candidateItem = candidateProps && candidateProps.item;
      var candidateId = String(candidateProps && candidateProps.sourceItemId || candidateItem && candidateItem.Id || '');
      var candidateSourceListName = String(candidateProps && candidateProps.sourceListName || '').toLowerCase();
      if (candidateId === masterId && (!sourceListName || candidateSourceListName === sourceListName)
        && this.getRecurrenceData(candidateItem)) {
        return candidateItem;
      }
    }
    return undefined;
  }

  private getColorWithOpacity(color: string, opacityPercent: number): string {
    var value = String(color || '#000000').replace('#', '');
    if (value.length === 3) {
      value = value.charAt(0) + value.charAt(0) + value.charAt(1) + value.charAt(1) + value.charAt(2) + value.charAt(2);
    }
    var red = parseInt(value.substring(0, 2), 16);
    var green = parseInt(value.substring(2, 4), 16);
    var blue = parseInt(value.substring(4, 6), 16);
    if (isNaN(red) || isNaN(green) || isNaN(blue)) {
      return 'rgba(0, 0, 0, .28)';
    }
    var opacity = Math.max(0, Math.min(100, opacityPercent)) / 100;
    return 'rgba(' + String(red) + ', ' + String(green) + ', ' + String(blue) + ', ' + String(opacity) + ')';
  }

  private renderEventDetails(): React.ReactElement<any> {
    var event = this.state.selectedEvent;
    if (!event || this.props.enableEventDetails !== true) {
      return undefined;
    }
    var itemId = Number(event.extendedProps && event.extendedProps.sourceItemId || 0);
    var showItemLink = this.props.showLinkToItem === true && itemId > 0;
    var itemUrl = showItemLink ? this.getItemLinkUrl(itemId, event) : '';
    var linkTitle = showItemLink && this.props.detailsLinkPresentation === 'title';
    var recurrenceDescription = this.getRecurrenceDescription(event);
    return (
      <div className={styles.eventDetailsBackdrop} role='presentation' onClick={() => this.closeEventDetails()}>
        <section className={styles.eventDetailsPanel} role='dialog' aria-modal='true'
          aria-label={event.title || 'Event details'} onClick={(clickEvent: React.MouseEvent<HTMLElement>) => clickEvent.stopPropagation()}>
          <button type='button' className={styles.eventDetailsClose} aria-label='Close event details'
            title='Close event details' onClick={() => this.closeEventDetails()}>&times;</button>
          <h2 className={styles.eventDetailsTitle}>
            {linkTitle ? <a href={itemUrl}>{event.title}</a> : event.title}
          </h2>
          <dl className={styles.eventDetailsList}>
            {this.getEventDetailFields(event).map((entry: any, index: number) => {
              var displayValue = this.formatEventDetailValue(event, entry);
              var renderAsHtml = displayValue !== '\u2014' && this.isRichTextDetailField(entry);
              if (entry.hideWhenEmpty === true && displayValue === '\u2014') {
                return null;
              }
              return (
                <div className={entry.showLabel === false ? styles.eventDetailsRowNoLabel : styles.eventDetailsRow}
                  key={String(entry.field) + '-' + String(index)}>
                  {entry.showLabel !== false && <dt>{String(entry.label || entry.field)}</dt>}
                  {renderAsHtml
                    ? <dd style={{ whiteSpace: 'normal' }}
                      dangerouslySetInnerHTML={{ __html: this.sanitizeRichText(displayValue) } as any} />
                    : <dd>{displayValue}</dd>}
                </div>
              );
            })}
            {recurrenceDescription
              && <div className={styles.eventDetailsRow} key='recurrence-information'>
                <dt>{strings.RecurrenceDetailsLabel}</dt>
                <dd>{recurrenceDescription}</dd>
              </div>}
          </dl>
          {showItemLink && !linkTitle
            && <a className={styles.eventDetailsOpenButton} href={itemUrl}>{strings.OpenEventDetailsButton}</a>}
        </section>
      </div>
    );
  }

  private getItemLinkUrl(itemId: number, event?: any): string {
    var extendedProps = event && event.extendedProps ? event.extendedProps : {};
    var targetPageUrl = String(extendedProps.sourceTargetPageUrl || this.props.linkTargetPageUrl || '').trim();
    var sourceListName = String(extendedProps.sourceListName || this.props.listName || '').trim();
    var targetIdParam = String(this.props.linkTargetIdParam || 'itemid').trim() || 'itemid';

    if (targetPageUrl) {
      var normalizedUrl = targetPageUrl;
      if (!/^https?:\/\//i.test(normalizedUrl)) {
        normalizedUrl = this.getWebUrl() + (normalizedUrl.charAt(0) === '/' ? normalizedUrl : '/' + normalizedUrl);
      }

      var embeddedUrl = resolveEmbeddedPageUrl(normalizedUrl);
      if (embeddedUrl) {
        normalizedUrl = /^https?:\/\//i.test(embeddedUrl)
          ? embeddedUrl
          : this.getWebUrl() + (embeddedUrl.charAt(0) === '/' ? embeddedUrl : '/' + embeddedUrl);
      }

      var customUrl = appendQueryParam(normalizedUrl, targetIdParam, String(itemId));
      if (this.props.includeReturnUrlParam === true && typeof window !== 'undefined' && window.location) {
        customUrl = appendQueryParam(customUrl, 'return', window.location.href);
      }
      return customUrl;
    }

    return this.getWebUrl() + '/Lists/' + encodeURIComponent(sourceListName) + '/DispForm.aspx?ID=' + String(itemId);
  }

  private getItemValue(item: any, fieldName: string): any {
    if (!item || !fieldName) {
      return '';
    }

    var target = fieldName.toLowerCase();
    var key: string;
    for (key in item) {
      if (Object.prototype.hasOwnProperty.call(item, key) && key.toLowerCase() === target) {
        var value = item[key];
        if (value && value.results) {
          return value.results.join(', ');
        }
        if (value && typeof value === 'object') {
          return value.Title || value.Name || value.LookupValue || value.Email || JSON.stringify(value);
        }
        return value;
      }
    }

    var textValues = item.FieldValuesAsText || {};
    for (key in textValues) {
      if (Object.prototype.hasOwnProperty.call(textValues, key) && key.toLowerCase() === target) {
        return textValues[key];
      }
    }
    return '';
  }

  private getLaneValues(item: any): string[] {
    var fieldName = String(this.props.swimlaneFieldName || '').toLowerCase();
    var rawValue: any = undefined;
    var key: string;
    for (key in item) {
      if (Object.prototype.hasOwnProperty.call(item, key) && key.toLowerCase() === fieldName) {
        rawValue = item[key];
        break;
      }
    }

    if (rawValue === undefined && item.FieldValuesAsText) {
      for (key in item.FieldValuesAsText) {
        if (Object.prototype.hasOwnProperty.call(item.FieldValuesAsText, key) && key.toLowerCase() === fieldName) {
          rawValue = item.FieldValuesAsText[key];
          break;
        }
      }
    }

    var values: any[] = rawValue && Array.isArray(rawValue.results) ? rawValue.results : [rawValue];
    var labels = values.map((value: any) => {
      if (value && typeof value === 'object') {
        return String(value.Title || value.Name || value.LookupValue || value.Email || '').trim();
      }
      return String(value === undefined || value === null ? '' : value).trim();
    }).filter((value: string) => !!value);

    return labels.length > 0 ? labels : [this.props.swimlaneUnassignedLabel || strings.SwimlaneUnassignedDefault];
  }

  private getLaneDates(): Date[] {
    var dates: Date[] = [];
    var cursor = new Date(this.state.laneStartDate.getTime());
    var desiredCount = this.props.swimlaneDays === 14 || this.props.swimlaneDays === 30 ? this.props.swimlaneDays : 7;
    while (dates.length < desiredCount) {
      if (this.props.showWeekends !== false || (cursor.getDay() !== 0 && cursor.getDay() !== 6)) {
        dates.push(new Date(cursor.getTime()));
      }
      cursor.setDate(cursor.getDate() + 1);
    }
    return dates;
  }

  private eventOverlapsDate(event: any, date: Date): boolean {
    var start = new Date(event.start);
    if (isNaN(start.getTime())) {
      return false;
    }
    var end = event.end ? new Date(event.end) : new Date(start.getTime());
    var dayStart = new Date(date.getTime());
    var dayEnd = new Date(date.getTime());
    dayEnd.setDate(dayEnd.getDate() + 1);
    if (isNaN(end.getTime()) || end.getTime() <= start.getTime()) {
      end = new Date(start.getTime() + 1);
    }
    return start.getTime() < dayEnd.getTime() && end.getTime() > dayStart.getTime();
  }

  private formatLaneDate(date: Date): string {
    var dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    return dayNames[date.getDay()] + ' ' + String(date.getMonth() + 1) + '/' + String(date.getDate());
  }

  private formatEventTime(event: any): string {
    if (event.allDay) {
      return '';
    }
    var start = new Date(event.start);
    if (isNaN(start.getTime())) {
      return '';
    }
    var hours = start.getHours();
    var minutes = String(start.getMinutes());
    var suffix = hours >= 12 ? 'PM' : 'AM';
    var displayHours = hours % 12 || 12;
    return String(displayHours) + ':' + (minutes.length < 2 ? '0' + minutes : minutes) + ' ' + suffix;
  }

  private moveLaneRange(direction: number): void {
    var nextDate: Date;
    if (direction > 0) {
      var visibleDates = this.getLaneDates();
      nextDate = new Date(visibleDates[visibleDates.length - 1].getTime());
      nextDate.setDate(nextDate.getDate() + 1);
      while (this.props.showWeekends === false && (nextDate.getDay() === 0 || nextDate.getDay() === 6)) {
        nextDate.setDate(nextDate.getDate() + 1);
      }
    } else {
      nextDate = new Date(this.state.laneStartDate.getTime());
      var remainingDays = this.props.swimlaneDays === 14 || this.props.swimlaneDays === 30 ? this.props.swimlaneDays : 7;
      while (remainingDays > 0) {
        nextDate.setDate(nextDate.getDate() - 1);
        if (this.props.showWeekends !== false || (nextDate.getDay() !== 0 && nextDate.getDay() !== 6)) {
          remainingDays -= 1;
        }
      }
    }
    this.setState({ laneStartDate: nextDate }, () => this.loadEvents());
  }

  private moveLaneRangeToToday(): void {
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    this.setState({ laneStartDate: today }, () => this.loadEvents());
  }

  private renderSwimlanes(): React.ReactElement<any> {
    if (!this.props.swimlaneFieldName) {
      return <div className={styles.message}>{strings.SwimlaneSelectFieldMessage}</div>;
    }

    var dates = this.getLaneDates();
    var lanesByName: { [laneName: string]: any[] } = {};
    for (var i = 0; i < this.state.events.length; i += 1) {
      var event = this.state.events[i];
      var laneValues = this.getLaneValues(event.extendedProps.item);
      for (var laneIndex = 0; laneIndex < laneValues.length; laneIndex += 1) {
        var laneName = laneValues[laneIndex];
        lanesByName[laneName] = lanesByName[laneName] || [];
        lanesByName[laneName].push(event);
      }
    }

    var laneNames = Object.keys(lanesByName).sort();
    var gridStyle: any = { gridTemplateColumns: '180px repeat(' + String(dates.length) + ', minmax(120px, 1fr))' };
    var firstDate = dates.length > 0 ? this.formatLaneDate(dates[0]) : '';
    var lastDate = dates.length > 0 ? this.formatLaneDate(dates[dates.length - 1]) : '';

    return (
      <div className={styles.swimlanes}>
        <div className={styles.swimlaneToolbar}>
          <button type='button' className={styles.swimlaneIconButton} onClick={() => this.moveLaneRange(-1)} title={strings.SwimlanePreviousLabel} aria-label={strings.SwimlanePreviousLabel}>&lsaquo;</button>
          <button type='button' className={styles.swimlaneTodayButton} onClick={() => this.moveLaneRangeToToday()}>{strings.SwimlaneTodayLabel}</button>
          <button type='button' className={styles.swimlaneIconButton} onClick={() => this.moveLaneRange(1)} title={strings.SwimlaneNextLabel} aria-label={strings.SwimlaneNextLabel}>&rsaquo;</button>
          <span className={styles.swimlaneRange}>{firstDate + ' - ' + lastDate}</span>
        </div>
        <div className={styles.swimlaneScroller}>
          <div className={styles.swimlaneHeader} style={gridStyle}>
            <div className={styles.swimlaneCorner}>{strings.SwimlaneLaneLabel}</div>
            {dates.map((date: Date) => <div key={this.formatLocalDate(date)} className={styles.swimlaneDayHeader}>{this.formatLaneDate(date)}</div>)}
          </div>
          {laneNames.length === 0 && <div className={styles.swimlaneEmpty}>{strings.SwimlaneNoEventsMessage}</div>}
          {laneNames.map((currentLaneName: string) => (
            <div key={currentLaneName} className={styles.swimlaneRow} style={gridStyle}>
              <div className={styles.swimlaneName} title={currentLaneName}>{currentLaneName}</div>
              {dates.map((date: Date) => (
                <div key={this.formatLocalDate(date)} className={styles.swimlaneCell}>
                  {lanesByName[currentLaneName].filter((laneEvent: any) => this.eventOverlapsDate(laneEvent, date)).map((laneEvent: any) => {
                    var eventStyle = laneEvent.extendedProps.eventStyle || {};
                    var isSelected = Number(laneEvent.extendedProps.sourceItemId || 0) === this.state.selectedItemId
                      && (!this.state.selectedEventKey
                        || String(laneEvent.extendedProps.sourceItemKey || '') === this.state.selectedEventKey);
                    var laneEventStyle: any = {
                      backgroundColor: isSelected ? this.props.selectedEventBackgroundColor : eventStyle.backgroundColor,
                      borderColor: isSelected ? this.props.selectedEventBorderColor : eventStyle.borderColor,
                      color: isSelected ? this.props.selectedEventTextColor : eventStyle.textColor,
                      fontFamily: isSelected ? this.props.selectedEventFontFamily : (eventStyle.fontFamily || this.props.eventFontFamily),
                      fontSize: isSelected ? this.props.selectedEventFontSize : (eventStyle.fontSize || this.props.eventFontSize),
                      fontStyle: isSelected ? this.props.selectedEventFontStyle : (eventStyle.fontStyle || this.props.eventFontStyle),
                      fontWeight: isSelected ? (this.props.selectedEventFontBold !== false ? 'bold' : 'normal')
                        : (eventStyle.fontWeight || (this.props.eventFontBold ? 'bold' : 'normal')),
                      borderWidth: String(isSelected ? this.props.selectedEventBorderWidth : this.props.eventBorderWidth) + 'px',
                      borderRadius: String(isSelected ? this.props.selectedEventCornerRadius : this.props.eventCornerRadius) + 'px'
                    };
                    var eventTime = this.formatEventTime(laneEvent);
                    return (
                      <button key={laneEvent.id} type='button' className={styles.swimlaneEvent} style={laneEventStyle}
                        aria-pressed={isSelected}
                        title={laneEvent.title} onClick={() => this.handleEventActivated(parseInt(laneEvent.id, 10), laneEvent)}>
                        {eventTime && <span className={styles.swimlaneEventTime}>{eventTime}</span>}
                        <span>{laneEvent.title}</span>
                      </button>
                    );
                  })}
                </div>
              ))}
            </div>
          ))}
        </div>
      </div>
    );
  }

  private compareValues(leftValue: any, rightValue: any): number {
    var left = String(leftValue === undefined || leftValue === null ? '' : leftValue).trim();
    var right = String(rightValue === undefined || rightValue === null ? '' : rightValue).trim();
    var leftNumber = Number(left);
    var rightNumber = Number(right);
    if (left && right && !isNaN(leftNumber) && !isNaN(rightNumber)) {
      return leftNumber === rightNumber ? 0 : (leftNumber > rightNumber ? 1 : -1);
    }
    var leftDate = Date.parse(left);
    var rightDate = Date.parse(right);
    if (left && right && !isNaN(leftDate) && !isNaN(rightDate)) {
      return leftDate === rightDate ? 0 : (leftDate > rightDate ? 1 : -1);
    }
    left = left.toLowerCase();
    right = right.toLowerCase();
    return left === right ? 0 : (left > right ? 1 : -1);
  }

  private formatLocalDate(value: Date): string {
    var year = value.getFullYear();
    var month = String(value.getMonth() + 1);
    var day = String(value.getDate());
    return String(year) + '-' + (month.length < 2 ? '0' + month : month) + '-' + (day.length < 2 ? '0' + day : day);
  }

  private resolveConditionValue(condition: ICalendarCondition): { value: any; compareDateOnly: boolean } {
    if (String(condition.valueType || '').toLowerCase() !== 'expression') {
      return { value: condition.value, compareDateOnly: false };
    }

    var expression = String(condition.value || '').trim().toLowerCase();
    var dateMatch = /^date\((-?\d+)\)$/.exec(expression);
    if (expression === 'today' || dateMatch) {
      var dateValue = new Date();
      if (dateMatch) {
        dateValue.setDate(dateValue.getDate() + parseInt(dateMatch[1], 10));
      }
      return { value: this.formatLocalDate(dateValue), compareDateOnly: true };
    }

    if (expression === 'now') {
      return { value: new Date().toISOString(), compareDateOnly: false };
    }

    var user = this.props.spfxContext.pageContext.user;
    var legacyContext = (this.props.spfxContext.pageContext as any).legacyPageContext;
    if (expression === 'me.email') {
      return { value: user ? user.email : '', compareDateOnly: false };
    }
    if (expression === 'me.login') {
      return { value: user ? user.loginName : '', compareDateOnly: false };
    }
    if (expression === 'me.title') {
      return { value: user ? user.displayName : '', compareDateOnly: false };
    }
    if (expression === 'me.id') {
      return { value: legacyContext ? String(legacyContext.userId || '') : '', compareDateOnly: false };
    }
    if (expression === 'me') {
      return { value: user ? String(user.displayName || user.loginName || '') : '', compareDateOnly: false };
    }

    return { value: condition.value, compareDateOnly: false };
  }

  private conditionMatches(item: any, condition: ICalendarCondition): boolean {
    var candidate = String(this.getItemValue(item, condition.field) || '').trim();
    var resolvedValue = this.resolveConditionValue(condition);
    var expected = String(resolvedValue.value === undefined || resolvedValue.value === null ? '' : resolvedValue.value).trim();
    if (resolvedValue.compareDateOnly) {
      var candidateDate = new Date(candidate);
      if (!isNaN(candidateDate.getTime())) {
        candidate = this.formatLocalDate(candidateDate);
      }
    }
    var normalizedCandidate = candidate.toLowerCase();
    var normalizedExpected = expected.toLowerCase();
    var comparison = this.compareValues(candidate, expected);

    switch (String(condition.operator || 'eq').toLowerCase()) {
      case 'ne': return normalizedCandidate !== normalizedExpected;
      case 'contains': return normalizedCandidate.indexOf(normalizedExpected) >= 0;
      case 'notcontains': return normalizedCandidate.indexOf(normalizedExpected) < 0;
      case 'startswith': return normalizedCandidate.indexOf(normalizedExpected) === 0;
      case 'endswith': return normalizedCandidate.lastIndexOf(normalizedExpected) === normalizedCandidate.length - normalizedExpected.length;
      case 'gt': return comparison > 0;
      case 'ge': return comparison >= 0;
      case 'lt': return comparison < 0;
      case 'le': return comparison <= 0;
      default: return normalizedCandidate === normalizedExpected;
    }
  }

  private conditionsMatch(item: any, conditions: ICalendarCondition[]): boolean {
    if (!conditions || conditions.length === 0) {
      return true;
    }

    var result = this.conditionMatches(item, conditions[0]);
    for (var i = 1; i < conditions.length; i += 1) {
      var matches = this.conditionMatches(item, conditions[i]);
      result = String(conditions[i].logical || 'and').toLowerCase() === 'or' ? result || matches : result && matches;
    }
    return result;
  }

  private parseConditions(source: string): ICalendarCondition[] {
    if (!String(source || '').trim()) {
      return [];
    }
    try {
      var parsed = JSON.parse(source);
      return Array.isArray(parsed) ? parsed.filter((item: any) => !!item && !!item.field) : [];
    } catch (_error) {
      this.logDiagnostic('Filter JSON is invalid; no filter was applied.');
      return [];
    }
  }

  private parseStyleRules(sourceListName: string): ICalendarStyleRule[] {
    if (!String(this.props.conditionalStyleJson || '').trim()) {
      return [];
    }
    try {
      var parsed = JSON.parse(this.props.conditionalStyleJson);
      if (!Array.isArray(parsed)) {
        return [];
      }
      var groupedRules: { [ruleName: string]: ICalendarStyleRule } = {};
      var orderedNames: string[] = [];
      for (var index = 0; index < parsed.length; index += 1) {
        var sourceRule = parsed[index] || {};
        var ruleSourceList = String(sourceRule.sourceListName || '');
        if (ruleSourceList && ruleSourceList.toLowerCase() !== sourceListName.toLowerCase()) {
          continue;
        }
        var ruleName = String(sourceRule.rule || sourceRule.ruleName || 'rule-' + String(index));
        var groupedRuleName = ruleSourceList.toLowerCase() + '|' + ruleName;
        if (!groupedRules[groupedRuleName]) {
          groupedRules[groupedRuleName] = {
            enabled: sourceRule.enabled !== false,
            priority: typeof sourceRule.priority === 'number' ? sourceRule.priority : index,
            conditions: [],
            style: sourceRule.style || sourceRule
          };
          orderedNames.push(groupedRuleName);
        }
        groupedRules[groupedRuleName].enabled = sourceRule.enabled !== false;
        groupedRules[groupedRuleName].priority = typeof sourceRule.priority === 'number' ? sourceRule.priority : index;
        groupedRules[groupedRuleName].style = sourceRule.style || sourceRule;
        var sourceConditions = Array.isArray(sourceRule.conditions)
          ? sourceRule.conditions
          : (Array.isArray(sourceRule.condition) ? sourceRule.condition : []);
        if (sourceConditions.length === 0 && (sourceRule.conditionField || sourceRule.field)) {
          sourceConditions = [{
            field: sourceRule.conditionField || sourceRule.field,
            operator: sourceRule.operator,
            logical: sourceRule.logical,
            valueType: sourceRule.valueType,
            value: sourceRule.value
          }];
        }
        groupedRules[groupedRuleName].conditions = groupedRules[groupedRuleName].conditions.concat(sourceConditions);
      }
      return orderedNames.map((name: string) => groupedRules[name])
        .sort((left: ICalendarStyleRule, right: ICalendarStyleRule) => left.priority - right.priority);
    } catch (_error) {
      this.logDiagnostic('Conditional style JSON is invalid; no rules were applied.');
      return [];
    }
  }

  private normalizeHexColor(value: string, fallback: string): string {
    var color = String(value || '').trim();
    return /^#[0-9a-f]{3}$/i.test(color) || /^#[0-9a-f]{6}$/i.test(color) ? color : fallback;
  }

  private getEventStyle(item: any, rules: ICalendarStyleRule[], sourceBackgroundColor?: string): any {
    var style: any = {
      backgroundColor: this.normalizeHexColor(sourceBackgroundColor, this.props.eventBackgroundColor || '#3788d8'),
      borderColor: this.props.eventBorderColor || '#2c6faa',
      textColor: this.props.eventTextColor || '#ffffff'
    };
    for (var i = 0; i < rules.length; i += 1) {
      if (rules[i].enabled && this.conditionsMatch(item, rules[i].conditions)) {
        var override = rules[i].style || {};
        style.backgroundColor = override.backgroundColor || style.backgroundColor;
        style.borderColor = override.borderColor || style.borderColor;
        style.textColor = override.textColor || override.color || style.textColor;
        style.fontFamily = override.fontFamily || style.fontFamily;
        style.fontSize = override.fontSize || style.fontSize;
        style.fontStyle = override.fontStyle || style.fontStyle;
        style.fontWeight = override.fontWeight || style.fontWeight;
      }
    }
    return style;
  }

  private getVisibleEventRange(): { start: Date; end: Date } {
    if (this.props.enableSwimlanes === true) {
      var laneDates = this.getLaneDates();
      var laneStart = laneDates.length > 0 ? new Date(laneDates[0].getTime()) : new Date(this.state.laneStartDate.getTime());
      var laneEnd = laneDates.length > 0 ? new Date(laneDates[laneDates.length - 1].getTime()) : new Date(laneStart.getTime());
      laneStart.setHours(0, 0, 0, 0);
      laneEnd.setHours(0, 0, 0, 0);
      laneEnd.setDate(laneEnd.getDate() + 1);
      return { start: laneStart, end: laneEnd };
    }

    if (this._calendar && this._calendar.view && this._calendar.view.activeStart && this._calendar.view.activeEnd) {
      return {
        start: new Date(this._calendar.view.activeStart.getTime()),
        end: new Date(this._calendar.view.activeEnd.getTime())
      };
    }

    var fallbackStart = new Date();
    fallbackStart.setDate(1);
    fallbackStart.setHours(0, 0, 0, 0);
    var fallbackEnd = new Date(fallbackStart.getTime());
    fallbackEnd.setMonth(fallbackEnd.getMonth() + 1);
    return { start: fallbackStart, end: fallbackEnd };
  }

  private itemOverlapsRange(item: any, rangeStart: Date, rangeEnd: Date): boolean {
    var start = new Date(item.EventDate);
    if (isNaN(start.getTime())) {
      return false;
    }
    var end = item.EndDate ? new Date(item.EndDate) : new Date(start.getTime() + 1);
    if (isNaN(end.getTime()) || end.getTime() <= start.getTime()) {
      end = new Date(start.getTime() + 1);
    }
    return start.getTime() < rangeEnd.getTime() && end.getTime() >= rangeStart.getTime();
  }

  private getItemTitle(item: any): string {
    if (item && item.Title !== undefined && item.Title !== null) {
      return String(item.Title);
    }
    var textValues = item && item.FieldValuesAsText;
    return textValues && textValues.Title ? String(textValues.Title) : '';
  }

  private getXmlAttribute(element: Element, name: string, defaultValue?: string): string {
    if (!element) {
      return defaultValue || '';
    }
    var value = element.getAttribute(name);
    return value === null || value === undefined ? (defaultValue || '') : value;
  }

  private getFirstElementChild(element: Element): Element {
    if (!element) {
      return undefined;
    }
    for (var index = 0; index < element.childNodes.length; index += 1) {
      var node = element.childNodes[index];
      if (node.nodeType === 1) {
        return node as Element;
      }
    }
    return undefined;
  }

  private getRecurrenceWeekdays(element: Element): any[] {
    var weekdays: any[] = [];
    var weekdayMap: { [key: string]: any } = {
      su: RRuleEngine.SU,
      mo: RRuleEngine.MO,
      tu: RRuleEngine.TU,
      we: RRuleEngine.WE,
      th: RRuleEngine.TH,
      fr: RRuleEngine.FR,
      sa: RRuleEngine.SA
    };
    if (this.getXmlAttribute(element, 'day').toLowerCase() === 'true') {
      return [RRuleEngine.SU, RRuleEngine.MO, RRuleEngine.TU, RRuleEngine.WE,
        RRuleEngine.TH, RRuleEngine.FR, RRuleEngine.SA];
    }
    if (this.getXmlAttribute(element, 'weekday').toLowerCase() === 'true') {
      weekdays = weekdays.concat([RRuleEngine.MO, RRuleEngine.TU, RRuleEngine.WE, RRuleEngine.TH, RRuleEngine.FR]);
    }
    if (this.getXmlAttribute(element, 'weekend_day').toLowerCase() === 'true') {
      weekdays = weekdays.concat([RRuleEngine.SU, RRuleEngine.SA]);
    }
    Object.keys(weekdayMap).forEach((key: string) => {
      if (this.getXmlAttribute(element, key).toLowerCase() === 'true') {
        if (weekdays.indexOf(weekdayMap[key]) < 0) {
          weekdays.push(weekdayMap[key]);
        }
      }
    });
    return weekdays;
  }

  private getRecurrenceWeekday(value: string): any {
    var weekdayMap: { [key: string]: any } = {
      su: RRuleEngine.SU,
      mo: RRuleEngine.MO,
      tu: RRuleEngine.TU,
      we: RRuleEngine.WE,
      th: RRuleEngine.TH,
      fr: RRuleEngine.FR,
      sa: RRuleEngine.SA
    };
    return weekdayMap[String(value || '').toLowerCase()];
  }

  private getRecurrenceData(item: any): string {
    var recurrenceData = item && item.RecurrenceData;
    if ((recurrenceData === undefined || recurrenceData === null || recurrenceData === '')
      && item && item.FieldValuesAsText) {
      recurrenceData = item.FieldValuesAsText.RecurrenceData;
    }
    var value = String(recurrenceData || '').trim();
    if (value.indexOf('&lt;') >= 0) {
      value = value.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'").replace(/&amp;/g, '&');
    }
    return value;
  }

  private getWeekdayPosition(value: string): number {
    switch (String(value || '').toLowerCase()) {
      case 'first': return 1;
      case 'second': return 2;
      case 'third': return 3;
      case 'fourth': return 4;
      case 'last': return -1;
      default: return 1;
    }
  }

  private createRecurrenceRule(item: any): any {
    var recurrenceData = this.getRecurrenceData(item);
    var start = new Date(item && item.EventDate);
    if (!recurrenceData || isNaN(start.getTime()) || typeof DOMParser === 'undefined') {
      return undefined;
    }

    try {
      var documentValue = new DOMParser().parseFromString(recurrenceData, 'text/xml');
      if (documentValue.getElementsByTagName('parsererror').length > 0) {
        return undefined;
      }
      var repeatContainers = documentValue.getElementsByTagName('repeat');
      var pattern = repeatContainers.length > 0 ? this.getFirstElementChild(repeatContainers[0] as Element) : undefined;
      if (!pattern) {
        return undefined;
      }

      var options: any = { dtstart: start };
      var patternName = String(pattern.localName || pattern.nodeName || '').toLowerCase();
      var firstDayElements = documentValue.getElementsByTagName('firstDayOfWeek');
      if (firstDayElements.length > 0) {
        var weekStart = this.getRecurrenceWeekday(String(firstDayElements[0].textContent || ''));
        if (weekStart) {
          options.wkst = weekStart;
        }
      }
      var interval = parseInt(this.getXmlAttribute(pattern, 'dayFrequency',
        this.getXmlAttribute(pattern, 'weekFrequency',
          this.getXmlAttribute(pattern, 'monthFrequency',
            this.getXmlAttribute(pattern, 'yearFrequency', '1')))), 10);
      options.interval = !isNaN(interval) && interval > 0 ? interval : 1;

      if (patternName === 'daily') {
        options.freq = RRuleEngine.DAILY;
        if (this.getXmlAttribute(pattern, 'weekday').toLowerCase() === 'true') {
          options.byweekday = [RRuleEngine.MO, RRuleEngine.TU, RRuleEngine.WE, RRuleEngine.TH, RRuleEngine.FR];
        }
      } else if (patternName === 'weekly') {
        options.freq = RRuleEngine.WEEKLY;
        options.byweekday = this.getRecurrenceWeekdays(pattern);
        if (options.byweekday.length === 0) {
          options.byweekday = [[RRuleEngine.SU, RRuleEngine.MO, RRuleEngine.TU, RRuleEngine.WE,
            RRuleEngine.TH, RRuleEngine.FR, RRuleEngine.SA][start.getUTCDay()]];
        }
      } else if (patternName === 'monthly') {
        options.freq = RRuleEngine.MONTHLY;
        options.bymonthday = parseInt(this.getXmlAttribute(pattern, 'day', String(start.getUTCDate())), 10);
      } else if (patternName === 'monthlybyday') {
        options.freq = RRuleEngine.MONTHLY;
        var monthlyWeekdays = this.getRecurrenceWeekdays(pattern);
        if (monthlyWeekdays.length === 0) {
          monthlyWeekdays = [this.getRecurrenceWeekday(['su', 'mo', 'tu', 'we', 'th', 'fr', 'sa'][start.getUTCDay()])];
        }
        var monthlyPosition = this.getWeekdayPosition(this.getXmlAttribute(pattern, 'weekdayOfMonth'));
        options.byweekday = monthlyWeekdays;
        options.bysetpos = monthlyPosition;
      } else if (patternName === 'yearly') {
        options.freq = RRuleEngine.YEARLY;
        options.bymonth = parseInt(this.getXmlAttribute(pattern, 'month', String(start.getUTCMonth() + 1)), 10);
        options.bymonthday = parseInt(this.getXmlAttribute(pattern, 'day', String(start.getUTCDate())), 10);
      } else if (patternName === 'yearlybyday') {
        options.freq = RRuleEngine.YEARLY;
        options.bymonth = parseInt(this.getXmlAttribute(pattern, 'month', String(start.getUTCMonth() + 1)), 10);
        var yearlyWeekdays = this.getRecurrenceWeekdays(pattern);
        if (yearlyWeekdays.length === 0) {
          yearlyWeekdays = [this.getRecurrenceWeekday(['su', 'mo', 'tu', 'we', 'th', 'fr', 'sa'][start.getUTCDay()])];
        }
        var yearlyPosition = this.getWeekdayPosition(this.getXmlAttribute(pattern, 'weekdayOfMonth'));
        options.byweekday = yearlyWeekdays;
        options.bysetpos = yearlyPosition;
      } else {
        return undefined;
      }

      var repeatInstances = documentValue.getElementsByTagName('repeatInstances');
      if (repeatInstances.length > 0) {
        var count = parseInt(String(repeatInstances[0].textContent || ''), 10);
        if (!isNaN(count) && count > 0) {
          options.count = count;
        }
      }
      var windowEnds = documentValue.getElementsByTagName('windowEnd');
      if (windowEnds.length > 0) {
        var until = new Date(String(windowEnds[0].textContent || ''));
        if (!isNaN(until.getTime())) {
          options.until = until;
        }
      }

      return new RRuleEngine(options);
    } catch (error) {
      this.logDiagnostic('Failed to parse recurrence for item ' + String(item.Id) + ': '
        + String(error && error.message ? error.message : error));
      return undefined;
    }
  }

  private getRecurrenceExceptionKeys(items: any[]): { [key: string]: boolean } {
    var keys: { [key: string]: boolean } = {};
    items.forEach((item: any) => {
      var eventType = parseInt(String(item.EventType), 10);
      var masterId = item.MasterSeriesItemID || item.MasterSeriesItemId;
      var recurrenceId = item.RecurrenceID || item.RecurrenceId;
      var recurrenceDate = recurrenceId ? new Date(recurrenceId) : undefined;
      if ((eventType === 3 || eventType === 4) && masterId && recurrenceDate && !isNaN(recurrenceDate.getTime())) {
        keys[String(masterId) + '|' + String(recurrenceDate.getTime())] = true;
      }
    });
    return keys;
  }

  private getEventDuration(item: any, start: Date, end: Date, isRecurrenceMaster: boolean): number {
    if (isRecurrenceMaster) {
      var durationSeconds = parseFloat(String(item && item.Duration));
      if (!isNaN(durationSeconds) && durationSeconds > 0) {
        return durationSeconds * 1000;
      }
      var isAllDay = item && (item.fAllDayEvent === true || item.fAllDayEvent === 1 || item.fAllDayEvent === '1');
      return isAllDay ? 24 * 60 * 60 * 1000 : 60 * 60 * 1000;
    }

    return end && !isNaN(end.getTime()) && end.getTime() > start.getTime()
      ? end.getTime() - start.getTime()
      : 1;
  }

  private itemSpansMultipleDates(item: any): boolean {
    var start = new Date(item && item.EventDate);
    var end = new Date(item && item.EndDate);
    return !isNaN(start.getTime()) && !isNaN(end.getTime())
      && (start.getUTCFullYear() !== end.getUTCFullYear()
        || start.getUTCMonth() !== end.getUTCMonth()
        || start.getUTCDate() !== end.getUTCDate());
  }

  private isRecurrenceMaster(item: any, start: Date, end: Date, eventType: number): boolean {
    var hasRecurrenceFlag = item && (item.fRecurrence === true || item.fRecurrence === 1
      || item.fRecurrence === '1' || String(item.fRecurrence).toLowerCase() === 'true');
    var hasRecurrenceRule = !!this.getRecurrenceData(item);
    var spansMultipleDates = this.itemSpansMultipleDates(item);
    var isException = eventType === 3 || eventType === 4;

    return !isException && hasRecurrenceRule && (hasRecurrenceFlag || spansMultipleDates || eventType === 1);
  }

  private expandEventItems(items: any[], rangeStart: Date, rangeEnd: Date): ICalendarOccurrence[] {
    var occurrences: ICalendarOccurrence[] = [];
    var exceptionKeys = this.getRecurrenceExceptionKeys(items);
    items.forEach((item: any) => {
      var start = new Date(item.EventDate);
      if (isNaN(start.getTime())) {
        return;
      }
      var end = item.EndDate ? new Date(item.EndDate) : undefined;
      var eventType = parseInt(String(item.EventType), 10);
      var isRecurrenceMaster = this.isRecurrenceMaster(item, start, end, eventType);
      var duration = this.getEventDuration(item, start, end, isRecurrenceMaster);

      if (!isRecurrenceMaster) {
        if (this.itemOverlapsRange(item, rangeStart, rangeEnd) && eventType !== 4) {
          occurrences.push({ id: String(item.Id), item: item, start: start, end: end });
        }
        return;
      }

      var rule = this.createRecurrenceRule(item);
      if (!rule) {
        this.logDiagnostic('Recurrence rule could not be expanded for item ' + String(item.Id) + '; using the master date.');
        if (this.itemOverlapsRange(item, rangeStart, rangeEnd)) {
          occurrences.push({ id: String(item.Id), item: item, start: start, end: end });
        }
        return;
      }

      var occurrenceStarts = rule.between(rangeStart, rangeEnd, true);
      var addedCount = 0;
      occurrenceStarts.forEach((occurrenceStart: Date) => {
        var exceptionKey = String(item.Id) + '|' + String(occurrenceStart.getTime());
        if (!exceptionKeys[exceptionKey]) {
          occurrences.push({
            id: String(item.Id) + '-' + String(occurrenceStart.getTime()),
            item: item,
            start: occurrenceStart,
            end: new Date(occurrenceStart.getTime() + duration)
          });
          addedCount += 1;
        }
      });
      this.logDiagnostic('Expanded recurring item ' + String(item.Id) + ' into ' + String(addedCount)
        + ' occurrence(s) for the visible range.');
    });
    return occurrences;
  }

  private logRenderedEventDiagnostics(): void {
    if (this.props.enableDiagnostics === false || !this._calendar || !this._calendarEl) {
      return;
    }

    var apiEvents = this._calendar.getEvents ? this._calendar.getEvents() : [];
    var eventElements = this._calendarEl.querySelectorAll('.fc-event');
    console.log('[Calendar] FullCalendar render diagnostics.', {
      apiEventCount: apiEvents.length,
      domEventCount: eventElements.length,
      viewType: this._calendar.view ? this._calendar.view.type : '',
      activeStart: this._calendar.view && this._calendar.view.activeStart,
      activeEnd: this._calendar.view && this._calendar.view.activeEnd
    });

    for (var index = 0; index < eventElements.length; index += 1) {
      var element = eventElements[index] as HTMLElement;
      var computedStyle = typeof window !== 'undefined' ? window.getComputedStyle(element) : undefined;
      console.log('[Calendar] Rendered event element.', {
        text: String(element.textContent || '').trim(),
        display: computedStyle ? computedStyle.display : '',
        visibility: computedStyle ? computedStyle.visibility : '',
        opacity: computedStyle ? computedStyle.opacity : '',
        color: computedStyle ? computedStyle.color : '',
        backgroundColor: computedStyle ? computedStyle.backgroundColor : '',
        width: element.offsetWidth,
        height: element.offsetHeight
      });
    }
  }

  private getEffectiveDataSources(): ICalendarDataSource[] {
    var configuredSources = Array.isArray(this.props.dataSources) ? this.props.dataSources : [];
    var sources = configuredSources
      .filter((source: ICalendarDataSource) => !!String(source && source.listName || '').trim())
      .map((source: ICalendarDataSource) => ({
        listName: String(source.listName).trim(),
        targetPageUrl: String(source.targetPageUrl || '').trim() === '__defaultForm__'
          ? '' : String(source.targetPageUrl || '').trim(),
        startFieldName: String(source.startFieldName || '').trim(),
        endFieldName: String(source.endFieldName || '').trim(),
        backgroundColor: String(source.backgroundColor || '').trim(),
        baseTemplate: Number(source.baseTemplate || 0)
      }));
    if (sources.length === 0 && this.props.listName) {
      sources.push({
        listName: String(this.props.listName),
        targetPageUrl: String(this.props.linkTargetPageUrl || ''),
        startFieldName: 'EventDate',
        endFieldName: 'EndDate',
        backgroundColor: '',
        baseTemplate: 106
      });
    }
    return sources;
  }

  private async loadEvents(): Promise<void> {
    if (!this._calendar) {
      return;
    }

    var dataSources = this.getEffectiveDataSources();
    if (dataSources.length === 0) {
      this._eventLoadSequence += 1;
      this._calendar.removeAllEventSources();
      this.setState({ events: [] });
      return;
    }

    this.setState({ loading: true, errorMessage: '' });
    var loadSequence = ++this._eventLoadSequence;

    try {
      var visibleRange = this.getVisibleEventRange();
      var rangeStart = visibleRange.start.toISOString();
      var rangeEnd = visibleRange.end.toISOString();
      var combinedEvents: any[] = [];
      for (var sourceIndex = 0; sourceIndex < dataSources.length; sourceIndex += 1) {
      var source = dataSources[sourceIndex];
      var listName = source.listName;
      var sourceKey = 'source-' + String(sourceIndex) + '-' + listName.toLowerCase().replace(/[^a-z0-9]+/g, '-');
      try {
      var isEventsList = Number(source.baseTemplate || 0) === 106;
      var configuredStartField = String(source.startFieldName || '');
      var configuredEndField = String(source.endFieldName || '');
      var startFieldName = !configuredStartField || configuredStartField === '__eventsDefault__'
        ? 'EventDate' : configuredStartField;
      var endFieldName = !configuredEndField || configuredEndField === '__eventsDefault__'
        ? 'EndDate' : configuredEndField;
      if (!isEventsList && (configuredStartField === '__eventsDefault__' || configuredEndField === '__eventsDefault__'
        || !configuredStartField || !configuredEndField)) {
        throw new Error('Could not determine this modern calendar view\'s default date columns. Select Start date column and End date column.');
      }
      var supportsBuiltInRecurrence = isEventsList && startFieldName === 'EventDate' && endFieldName === 'EndDate';
      this.logDiagnostic('Loading events from list "' + listName + '" for ' + rangeStart + ' through ' + rangeEnd);
      const webUrl = this.props.spfxContext.pageContext.web.absoluteUrl.replace(/\/$/, '');
      const baseUrl = webUrl + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/items"
        + "?$select=*,FieldValuesAsText&$expand=FieldValuesAsText";
      const overlapUrl = baseUrl
        + "&$filter=" + startFieldName + " lt datetime'" + rangeEnd + "' and (" + endFieldName + " ge datetime'" + rangeStart
        + "' or (" + endFieldName + " eq null and " + startFieldName + " ge datetime'" + rangeStart + "'))"
        + '&$orderby=' + startFieldName + '&$top=5000';
      const startDateUrl = baseUrl
        + "&$filter=" + startFieldName + " ge datetime'" + rangeStart + "' and " + startFieldName + " lt datetime'" + rangeEnd + "'"
        + '&$orderby=' + startFieldName + '&$top=5000';
      const recurrenceMastersUrl = baseUrl + '&$filter=fRecurrence eq 1&$top=5000';

      let usedStartDateFallback = false;
      let response = await this.props.spfxContext.spHttpClient.get(overlapUrl, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata' }
      });
      if (!response.ok) {
        this.logDiagnostic('Overlap query was rejected; retrying with start-date range. HTTP ' + String(response.status));
        usedStartDateFallback = true;
        response = await this.props.spfxContext.spHttpClient.get(startDateUrl, SPHttpClient.configurations.v1, {
          headers: { Accept: 'application/json;odata=nometadata' }
        });
      }
      if (!response.ok) {
        throw new Error('Request failed. HTTP ' + String(response.status) + ' ' + response.statusText);
      }

      if (loadSequence !== this._eventLoadSequence) {
        return;
      }

      let data = await response.json();
      let items = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      if (items.length === 0 && !usedStartDateFallback) {
        this.logDiagnostic('Overlap query returned no items; retrying with start-date range.');
        usedStartDateFallback = true;
        response = await this.props.spfxContext.spHttpClient.get(startDateUrl, SPHttpClient.configurations.v1, {
          headers: { Accept: 'application/json;odata=nometadata' }
        });
        if (!response.ok) {
          throw new Error('Fallback request failed. HTTP ' + String(response.status) + ' ' + response.statusText);
        }
        if (loadSequence !== this._eventLoadSequence) {
          return;
        }
        data = await response.json();
        items = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      }
      this.logRestItems(usedStartDateFallback ? 'visible range (start-date fallback)' : 'visible range (overlap)',
        usedStartDateFallback ? startDateUrl : overlapUrl, items);
      if (supportsBuiltInRecurrence) {
        var recurrenceResponse = await this.props.spfxContext.spHttpClient.get(recurrenceMastersUrl, SPHttpClient.configurations.v1, {
          headers: { Accept: 'application/json;odata=nometadata' }
        });
        if (recurrenceResponse.ok) {
        var recurrenceData = await recurrenceResponse.json();
        var recurrenceItems = recurrenceData && recurrenceData.value ? recurrenceData.value
          : (recurrenceData && recurrenceData.d && recurrenceData.d.results ? recurrenceData.d.results : []);
        this.logRestItems('recurrence masters', recurrenceMastersUrl, recurrenceItems);
        var itemIndexesById: { [id: string]: number } = {};
        items.forEach((item: any, index: number) => { itemIndexesById[String(item.Id)] = index; });
        recurrenceItems.forEach((recurrenceItem: any) => {
          var existingIndex = itemIndexesById[String(recurrenceItem.Id)];
          if (existingIndex === undefined) {
            itemIndexesById[String(recurrenceItem.Id)] = items.length;
            items.push(recurrenceItem);
          } else {
            Object.keys(recurrenceItem).forEach((key: string) => {
              items[existingIndex][key] = recurrenceItem[key];
            });
          }
        });
        this.logDiagnostic('Loaded recurrence masters. Count=' + String(recurrenceItems.length));
        } else {
          this.logDiagnostic('Recurrence master query was rejected. HTTP ' + String(recurrenceResponse.status));
        }
      }

      for (var itemIndex = 0; itemIndex < items.length; itemIndex += 1) {
        var detailItem = items[itemIndex];
        if (!supportsBuiltInRecurrence || !detailItem || !detailItem.Id || !this.itemSpansMultipleDates(detailItem)) {
          continue;
        }

        var recurrenceDetailUrl = webUrl + "/_api/web/lists/getByTitle('"
          + escapeODataText(listName) + "')/items(" + String(detailItem.Id)
          + ')?$select=Title,EventDate,EndDate,RecurrenceData,UID';
        var recurrenceDetailResponse = await this.props.spfxContext.spHttpClient.get(recurrenceDetailUrl,
          SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=verbose' } });
        if (!recurrenceDetailResponse.ok) {
          this.logDiagnostic('Recurrence detail query was rejected for item ' + String(detailItem.Id)
            + '. HTTP ' + String(recurrenceDetailResponse.status));
          continue;
        }
        if (loadSequence !== this._eventLoadSequence) {
          return;
        }

        var recurrenceDetailText = await recurrenceDetailResponse.text();
        var recurrenceDetailItem: any;
        try {
          recurrenceDetailItem = this.parseRestItemResponse(recurrenceDetailText);
        } catch (parseError) {
          this.logDiagnostic('Recurrence detail response could not be parsed for item ' + String(detailItem.Id)
            + ': ' + String(parseError && parseError.message ? parseError.message : parseError));
          continue;
        }
        if (!recurrenceDetailItem) {
          this.logDiagnostic('Recurrence detail response was empty or invalid for item ' + String(detailItem.Id));
          continue;
        }
        this.logRestItems('recurrence detail for item ' + String(detailItem.Id), recurrenceDetailUrl,
          [recurrenceDetailItem]);
        Object.keys(recurrenceDetailItem).forEach((key: string) => {
          detailItem[key] = recurrenceDetailItem[key];
        });
      }
      items.forEach((sourceItem: any) => {
        sourceItem.EventDate = sourceItem[startFieldName];
        sourceItem.EndDate = sourceItem[endFieldName] || sourceItem[startFieldName];
      });
      const filterConditions = this.parseConditions(this.props.filterJson).filter((condition: ICalendarCondition) => {
        var conditionSourceList = String(condition.sourceListName || '');
        return !conditionSourceList || conditionSourceList.toLowerCase() === listName.toLowerCase();
      });
      const styleRules = this.parseStyleRules(listName);
      const filteredItems = items.filter((item: any) => this.conditionsMatch(item, filterConditions));
      const occurrences = this.expandEventItems(filteredItems, visibleRange.start, visibleRange.end);
      const events = occurrences.map((occurrence: ICalendarOccurrence) => {
        const eventStyle = this.getEventStyle(occurrence.item, styleRules, source.backgroundColor);
        const sourceItemKey = sourceKey + ':' + String(occurrence.item.Id);
        return {
          id: sourceKey + ':' + occurrence.id,
          title: this.getItemTitle(occurrence.item),
          start: occurrence.start,
          end: occurrence.end,
          allDay: occurrence.item.fAllDayEvent === true || occurrence.item.fAllDayEvent === 1 || occurrence.item.fAllDayEvent === '1',
          backgroundColor: eventStyle.backgroundColor,
          borderColor: eventStyle.borderColor,
          textColor: eventStyle.textColor,
          extendedProps: {
            location: occurrence.item.Location || '',
            item: occurrence.item,
            sourceItemId: occurrence.item.Id,
            sourceItemKey: sourceItemKey,
            sourceListName: listName,
            sourceTargetPageUrl: source.targetPageUrl || '',
            eventStyle: eventStyle
          }
        };
      });

      if (loadSequence !== this._eventLoadSequence) {
        return;
      }

      this.logDiagnostic('Loaded events successfully. SharePoint=' + String(items.length)
        + ', after filters=' + String(filteredItems.length) + ', occurrences=' + String(occurrences.length)
        + ', mapped=' + String(events.length));
      if (this.props.enableDiagnostics !== false) {
        console.log('[Calendar] Mapped event payload.', events.map((event: any) => {
          return {
            id: event.id,
            title: event.title,
            start: event.start,
            end: event.end,
            allDay: event.allDay,
            backgroundColor: event.backgroundColor,
            borderColor: event.borderColor,
            textColor: event.textColor
          };
        }));
      }
      combinedEvents = combinedEvents.concat(events);
      } catch (sourceError) {
        console.error('[Calendar] Failed to load source "' + listName + '": '
          + String(sourceError && sourceError.message ? sourceError.message : sourceError));
      }
      }
      this.setState({ loading: false, errorMessage: '', events: combinedEvents }, () => {
        if (loadSequence !== this._eventLoadSequence || !this._calendar) {
          return;
        }
        this._calendar.removeAllEventSources();
        this._calendar.addEventSource(combinedEvents);
        if (typeof window !== 'undefined') {
          window.setTimeout(() => this.logRenderedEventDiagnostics(), 0);
        }
      });
    } catch (error) {
      if (loadSequence !== this._eventLoadSequence) {
        return;
      }
      const message = error && error.message ? error.message : String(error);
      console.error('[Calendar] Failed to load combined calendar sources: ' + message);
      this.setState({ loading: false, errorMessage: strings.LoadErrorMessage + ' ' + message });
    }
  }

  private renderDataSourceLegend(): React.ReactElement<any> {
    var sources = this.getEffectiveDataSources();
    if (sources.length < 2) {
      return undefined;
    }
    var panelClassName = styles.dataSourceLegendPanel
      + (this.state.legendOpen ? ' ' + styles.dataSourceLegendPanelOpen : '');
    var legendId = 'calendar-data-source-legend-' + String(this.props.spfxContext.instanceId || '').replace(/[^a-z0-9_-]/gi, '');
    return (
      <div className={styles.dataSourceLegend}>
        <button type='button' className={styles.dataSourceLegendButton}
          title={strings.DataSourceLegendButtonLabel}
          aria-label={strings.DataSourceLegendButtonLabel}
          aria-expanded={this.state.legendOpen}
          aria-controls={legendId}
          onClick={() => this.setState({ legendOpen: !this.state.legendOpen })}>
          <span className='ms-Icon ms-Icon--Info' aria-hidden='true' />
        </button>
        <div id={legendId} className={panelClassName} role='group'
          aria-label={strings.DataSourceLegendTitle}>
          <div className={styles.dataSourceLegendTitle}>{strings.DataSourceLegendTitle}</div>
          {sources.map((source: ICalendarDataSource, index: number) => {
            var backgroundColor = this.normalizeHexColor(source.backgroundColor,
              this.props.eventBackgroundColor || '#3788d8');
            var textColor = this.props.eventTextColor || '#ffffff';
            var sourceName = String(source.listName || '');
            return (
              <div className={styles.dataSourceLegendItem} key={sourceName + '-' + String(index)} title={sourceName}>
                <span className={styles.dataSourceLegendSwatch}
                  style={{ backgroundColor: backgroundColor, color: textColor }}>Aa</span>
                <span className={styles.dataSourceLegendName}>{sourceName}</span>
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  public render(): React.ReactElement<ICalendarProps> {
    this.logDiagnostic('render() called. listName=' + (this.props.listName || '(none)'));
    const calendarStyle: any = {
      backgroundColor: this.props.calendarBackgroundColor || '#ffffff',
      borderColor: this.props.webPartBorderColor || '#d2d0ce',
      borderStyle: 'solid',
      borderWidth: String(typeof this.props.webPartBorderWidth === 'number' ? this.props.webPartBorderWidth : 0) + 'px',
      borderRadius: String(typeof this.props.webPartCornerRadius === 'number' ? this.props.webPartCornerRadius : 0) + 'px',
      '--calendar-web-part-title-color': this.props.webPartTitleTextColor || '#323130',
      '--calendar-web-part-title-font-family': this.props.webPartTitleFontFamily || 'inherit',
      '--calendar-web-part-title-font-size': this.props.webPartTitleFontSize || '20px',
      '--calendar-web-part-title-font-style': this.props.webPartTitleFontStyle || 'normal',
      '--calendar-web-part-title-font-weight': this.props.webPartTitleFontBold !== false ? 'bold' : 'normal',
      '--calendar-web-part-title-alignment': this.props.webPartTitleAlignment || 'left',
      '--calendar-web-part-title-background': this.props.webPartTitleBackgroundColor || '#ffffff',
      '--calendar-web-part-title-min-height': String(typeof this.props.webPartTitleMinHeight === 'number'
        ? this.props.webPartTitleMinHeight : 0) + 'px',
      '--calendar-web-part-description-color': this.props.webPartDescriptionTextColor || '#605e5c',
      '--calendar-web-part-description-font-family': this.props.webPartDescriptionFontFamily || 'inherit',
      '--calendar-web-part-description-font-size': this.props.webPartDescriptionFontSize || '14px',
      '--calendar-web-part-description-font-style': this.props.webPartDescriptionFontStyle || 'normal',
      '--calendar-web-part-description-font-weight': this.props.webPartDescriptionFontBold === true ? 'bold' : 'normal',
      '--calendar-toolbar-title-color': this.props.calendarTitleTextColor || '#323130',
      '--calendar-toolbar-title-font-family': this.props.calendarTitleFontFamily || 'inherit',
      '--calendar-toolbar-title-font-size': this.props.calendarTitleFontSize || '28px',
      '--calendar-toolbar-title-font-style': this.props.calendarTitleFontStyle || 'normal',
      '--calendar-toolbar-title-font-weight': this.props.calendarTitleFontBold !== false ? 'bold' : 'normal',
      '--calendar-toolbar-button-background': this.props.toolbarButtonBackgroundColor || '#2f4358',
      '--calendar-toolbar-button-color': this.props.toolbarButtonTextColor || '#ffffff',
      '--calendar-toolbar-button-border': this.props.toolbarButtonBorderColor || '#2f4358',
      '--calendar-toolbar-button-active-background': this.props.toolbarButtonActiveBackgroundColor || '#172536',
      '--calendar-toolbar-button-active-color': this.props.toolbarButtonActiveTextColor || '#ffffff',
      '--calendar-toolbar-button-font-family': this.props.toolbarButtonFontFamily || 'inherit',
      '--calendar-toolbar-button-font-size': this.props.toolbarButtonFontSize || '16px',
      '--calendar-toolbar-button-font-style': this.props.toolbarButtonFontStyle || 'normal',
      '--calendar-toolbar-button-font-weight': this.props.toolbarButtonFontBold === true ? 'bold' : 'normal',
      '--calendar-toolbar-button-border-width': String(typeof this.props.toolbarButtonBorderWidth === 'number' ? this.props.toolbarButtonBorderWidth : 1) + 'px',
      '--calendar-toolbar-button-corner-radius': String(typeof this.props.toolbarButtonCornerRadius === 'number' ? this.props.toolbarButtonCornerRadius : 4) + 'px',
      '--calendar-event-font-family': this.props.eventFontFamily || 'inherit',
      '--calendar-event-font-size': this.props.eventFontSize || 'inherit',
      '--calendar-event-font-style': this.props.eventFontStyle || 'normal',
      '--calendar-event-font-weight': this.props.eventFontBold === true ? 'bold' : 'normal',
      '--calendar-event-border-width': String(typeof this.props.eventBorderWidth === 'number' ? this.props.eventBorderWidth : 1) + 'px',
      '--calendar-event-corner-radius': String(typeof this.props.eventCornerRadius === 'number' ? this.props.eventCornerRadius : 3) + 'px',
      '--details-overlay-background': this.getColorWithOpacity(this.props.detailsOverlayColor || '#000000',
        typeof this.props.detailsOverlayOpacity === 'number' ? this.props.detailsOverlayOpacity : 28),
      '--details-panel-background': this.props.detailsPanelBackgroundColor || '#ffffff',
      '--details-panel-border-color': this.props.detailsPanelBorderColor || '#8a8886',
      '--details-panel-border-width': String(typeof this.props.detailsPanelBorderWidth === 'number' ? this.props.detailsPanelBorderWidth : 1) + 'px',
      '--details-panel-corner-radius': String(typeof this.props.detailsPanelCornerRadius === 'number' ? this.props.detailsPanelCornerRadius : 6) + 'px',
      '--details-panel-max-width': String(typeof this.props.detailsPanelMaxWidth === 'number' ? this.props.detailsPanelMaxWidth : 680) + 'px',
      '--details-panel-padding': String(typeof this.props.detailsPanelPadding === 'number' ? this.props.detailsPanelPadding : 30) + 'px',
      '--details-panel-shadow': this.props.detailsPanelShadow !== false ? '0 12px 36px rgba(0, 0, 0, .24)' : 'none',
      '--details-title-color': this.props.detailsTitleTextColor || '#201f1e',
      '--details-title-font-family': this.props.detailsTitleFontFamily || 'inherit',
      '--details-title-font-size': this.props.detailsTitleFontSize || '24px',
      '--details-title-font-style': this.props.detailsTitleFontStyle || 'normal',
      '--details-title-font-weight': this.props.detailsTitleFontBold !== false ? 'bold' : 'normal',
      '--details-label-color': this.props.detailsLabelTextColor || '#323130',
      '--details-label-font-family': this.props.detailsLabelFontFamily || 'inherit',
      '--details-label-font-size': this.props.detailsLabelFontSize || '14px',
      '--details-label-font-style': this.props.detailsLabelFontStyle || 'normal',
      '--details-label-font-weight': this.props.detailsLabelFontBold !== false ? 'bold' : 'normal',
      '--details-value-color': this.props.detailsValueTextColor || '#201f1e',
      '--details-value-font-family': this.props.detailsValueFontFamily || 'inherit',
      '--details-value-font-size': this.props.detailsValueFontSize || '14px',
      '--details-value-font-style': this.props.detailsValueFontStyle || 'normal',
      '--details-value-font-weight': this.props.detailsValueFontBold === true ? 'bold' : 'normal',
      '--details-divider-color': this.props.detailsDividerColor || '#edebe9',
      '--details-divider-width': String(typeof this.props.detailsDividerWidth === 'number' ? this.props.detailsDividerWidth : 1) + 'px',
      '--details-close-color': this.props.detailsCloseButtonColor || '#323130',
      '--details-close-background': this.props.detailsCloseButtonBackgroundColor || 'transparent',
      '--details-close-hover-background': this.props.detailsCloseButtonHoverBackgroundColor || '#edebe9',
      '--details-close-corner-radius': String(typeof this.props.detailsCloseButtonCornerRadius === 'number'
        ? this.props.detailsCloseButtonCornerRadius : 2) + 'px'
    };
    var dataSources = this.getEffectiveDataSources();
    var headerClassName = styles.webPartHeader
      + (dataSources.length > 1 ? ' ' + styles.webPartHeaderWithLegend : '')
      + (String(this.props.description || '').trim() ? ' ' + styles.webPartHeaderWithDescription : '');
    return (
      <div className={styles.calendar} style={calendarStyle}>
        {this.renderEventDetails()}
        <div className={headerClassName}>
          <WebPartTitle key={'calendar-web-part-title-' + String(this.props.title || '')}
            displayMode={this.props.displayMode}
            className={styles.webPartTitle}
            title={this.props.title}
            updateProperty={this.props.fUpdateProperty} />
          {String(this.props.description || '').trim()
            && <div className={styles.webPartDescription}>{this.props.description}</div>}
          {this.renderDataSourceLegend()}
        </div>

        {
          dataSources.length > 0 ? (
            <div>
              {this.state.errorMessage && <div className={styles.error}>{this.state.errorMessage}</div>}
              {this.state.loading && <div className={styles.message}>{strings.LoadingMessage}</div>}
              {this.props.enableSwimlanes === true && this.renderSwimlanes()}
              <div className={styles.calendarContainer}
                style={{
                  display: this.props.enableSwimlanes === true ? 'none' : 'block',
                  minHeight: String(this.props.calendarHeight || 600) + 'px'
                }}
                ref={this._setCalendarRef} />
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
