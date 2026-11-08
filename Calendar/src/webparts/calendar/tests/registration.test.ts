/// <reference types="mocha" />
import { assert } from 'chai';
import Calendar from '../components/Calendar';
import CalendarWebPart from '../CalendarWebPart';

describe('Calendar registration integration', (): void => {
  var calendar: any;
  var webPart: any;

  beforeEach((): void => {
    calendar = Object.create(Calendar.prototype);
    calendar.props = {
      enableRegistration: true,
      registrationPageUrl: '/sites/events/SitePages/Register.aspx',
      includeRegistrationReturnUrl: false,
      enableEventDetails: false,
      spfxContext: { pageContext: { web: { absoluteUrl: 'https://example.test/sites/events' } } }
    };
    calendar.setState = (): void => { return; };
    calendar.rerenderSelectedEvents = (): void => { return; };
    calendar.logDiagnostic = (): void => { return; };
    webPart = Object.create(CalendarWebPart.prototype);
    webPart.properties = { enableDiagnostics: false };
    webPart.notifyDynamicData = (): void => { return; };
  });

  it('links a session to its parent registration event, not to the session ID', (): void => {
    var url = calendar.getRegistrationUrl(42, { extendedProps: { item: { EREventId: 7 } } });
    assert.equal(url, 'https://example.test/sites/events/SitePages/Register.aspx?sessionid=42&eventid=7');
  });

  it('infers the Sessions date defaults from the provisioned classic calendar query', (): void => {
    var query = '<Where><DateRangesOverlap><FieldRef Name="ERStart" /><FieldRef Name="EREnd" />'
      + '<Value Type="DateTime"><Month /></Value></DateRangesOverlap></Where>';
    assert.deepEqual(webPart.getCalendarDateFields(query), { startFieldName: 'ERStart', endFieldName: 'EREnd' });
  });

  it('preserves web-relative destinations, custom parameter names, queries and fragments', (): void => {
    calendar.props.registrationPageUrl = 'SitePages/Register.aspx?mode=register#booking';
    calendar.props.registrationSessionIdParam = 'session';
    calendar.props.registrationEventIdParam = 'event';
    var url = calendar.getRegistrationUrl(42, { EREventId: 7 });
    assert.equal(url, 'https://example.test/sites/events/SitePages/Register.aspx?mode=register&session=42&event=7#booking');
  });

  it('reads registration lookup IDs at each supported event level', (): void => {
    assert.equal(calendar.getRegistrationParentEventId({ EREventId: 7 }), 7);
    assert.equal(calendar.getRegistrationParentEventId({ extendedProps: { EREventId: 8 } }), 8);
    assert.equal(calendar.getRegistrationParentEventId({ extendedProps: { item: { EREventId: 9 } } }), 9);
    assert.equal(calendar.getRegistrationParentEventId({ EREventId: 7, EventId: 42 }), 7);
  });

  it('preserves legacy lookup fields and omits unavailable parent IDs', (): void => {
    assert.equal(calendar.getRegistrationParentEventId({ EventId: 8 }), 8);
    assert.equal(calendar.getRegistrationParentEventId({ extendedProps: { item: { EventLookupId: { Id: 9 } } } }), 9);
    assert.equal(calendar.getRegistrationParentEventId(undefined), 0);
    assert.notInclude(calendar.getRegistrationUrl(42), 'eventid=');
    calendar.props.enableRegistration = false;
    assert.equal(calendar.getRegistrationUrl(42, { EREventId: 7 }), '');
  });

  it('publishes the distinct parent event and session while preserving the original item identity', (): void => {
    var notifications: string[] = [];
    webPart.notifyDynamicData = (property: string): void => { notifications.push(property); };
    calendar.props.onEventSelectionChange = webPart.handleEventSelectionChange.bind(webPart);
    calendar.handleEventActivated(42, { extendedProps: { sourceItemKey: 'sessions:42', item: { EREventId: 7 } } });
    assert.equal(webPart.getPropertyValue('selectedItemId'), 42);
    assert.equal(webPart.getPropertyValue('selectedSessionId'), 42);
    assert.equal(webPart.getPropertyValue('selectedEventId'), 7);
    assert.equal(webPart.getPropertyValue('selectedEventKey'), 'sessions:42');
    assert.include(notifications, 'selectedEventId');
    assert.isTrue(webPart.getPropertyDefinitions().some((definition: any): boolean => {
      return definition.id === 'selectedEventId';
    }));
    webPart.handleEventSelectionChange(0, 'view');
    assert.equal(webPart.getPropertyValue('selectedEventId'), 0);
    assert.equal(webPart.getPropertyValue('selectedSessionId'), 0);
    webPart.handleEventSelectionChange(42, 'view');
    assert.equal(webPart.getPropertyValue('selectedEventId'), 0);
    assert.equal(webPart.getPropertyValue('selectedItemId'), 42);
  });
});
