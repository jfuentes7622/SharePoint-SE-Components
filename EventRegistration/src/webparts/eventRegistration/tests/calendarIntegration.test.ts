import { assert } from 'chai';
import EventRegistrationService from '../services/EventRegistrationService';
import EventRegistrationWebPart from '../EventRegistrationWebPart';

describe('Event Registration Sessions calendar provisioning', function(): void {
  var service: any;
  var posts: Array<{ url: string; body: any; headers: any }>;
  var existing: any;
  var failure: any;
  var query = '<Where><DateRangesOverlap><FieldRef Name="ERStart" /><FieldRef Name="EREnd" />'
    + '<Value Type="DateTime"><Month /></Value></DateRangesOverlap></Where>';

  beforeEach(function(): void {
    service = new EventRegistrationService('https://example.test', undefined, "Organizer's Events", false);
    posts = [];
    existing = undefined;
    failure = undefined;
    service.rest = {
      escape: function(value: string): string { return value.replace(/'/g, "''"); },
      item: function(data: any): any { return data.d || data; },
      rows: function(data: any): any[] { return data.value; },
      get: function(url: string): Promise<any> {
        assert.include(url, '/views?$select=');
        assert.include(decodeURIComponent(url), "$filter=Title eq 'Registration Sessions Calendar'");
        return failure ? Promise.reject(failure) : Promise.resolve({ data: { value: existing ? [existing] : [] } });
      },
      post: function(url: string, body: any, headers: any): Promise<any> {
        posts.push({ url: url, body: body, headers: headers });
        if (url.indexOf('/add') >= 0) {
          existing = { ViewType: 'CALENDAR', ViewQuery: body.parameters.Query, ViewData: '' };
        } else {
          existing.ViewData = body.ViewData;
        }
        return Promise.resolve({ data: existing });
      }
    };
  });

  it('creates a public classic calendar with the Sessions dates and title without changing the default', function(): Promise<void> {
    return service.ensureSessionsCalendarView().then(function(): Promise<void> {
      assert.include(posts[0].url, "getByTitle('Organizer''s Events Sessions')/views/add");
      var parameters = posts[0].body.parameters;
      assert.equal(parameters.__metadata.type, 'SP.ViewCreationInformation');
      assert.equal(parameters.ViewTypeKind, 524288);
      assert.deepEqual(parameters.ViewFields.results, ['ERStart', 'EREnd', 'Title']);
      assert.equal(parameters.Query, query);
      assert.isFalse(parameters.SetAsDefaultView);
      assert.isFalse(parameters.PersonalView);
      assert.equal(posts[1].body.__metadata.type, 'SP.View');
      assert.include(posts[1].body.ViewData, 'CalendarMonthTitle');
      assert.include(posts[1].body.ViewData, 'CalendarWeekTitle');
      assert.include(posts[1].body.ViewData, 'CalendarDayTitle');
      assert.equal(posts[1].headers['X-HTTP-Method'], 'MERGE');
      return service.ensureSessionsCalendarView();
    }).then(function(): void {
      assert.lengthOf(posts, 2);
    });
  });

  it('leaves an existing compatible calendar and its custom display mappings untouched', function(): Promise<void> {
    existing = { ViewType: 'CALENDAR', ViewQuery: query, ViewData: 'Custom mappings' };
    return service.ensureSessionsCalendarView().then(function(): void { assert.lengthOf(posts, 0); });
  });

  it('completes metadata when a previous setup was interrupted', function(): Promise<void> {
    existing = { ViewType: 'CALENDAR', ViewQuery: query, ViewData: '' };
    return service.ensureSessionsCalendarView().then(function(): void {
      assert.lengthOf(posts, 1);
      assert.include(posts[0].body.ViewData, 'CalendarMonthTitle');
    });
  });

  it('rejects a conflicting view instead of replacing it', function(): Promise<void> {
    existing = { ViewType: 'HTML', ViewQuery: '' };
    return service.ensureSessionsCalendarView().then(function(): void {
      assert.fail('Expected incompatible view failure');
    }, function(error: Error): void {
      assert.include(error.message, 'incompatible');
      assert.lengthOf(posts, 0);
    });
  });

  it('propagates permission errors without attempting creation', function(): Promise<void> {
    failure = { status: 403, message: 'Access denied' };
    return service.ensureSessionsCalendarView().then(function(): void {
      assert.fail('Expected discovery failure');
    }, function(error: any): void {
      assert.strictEqual(error, failure);
      assert.lengthOf(posts, 0);
    });
  });

  it('propagates calendar metadata failures and retries without duplicating the view', function(): Promise<void> {
    var post = service.rest.post;
    service.rest.post = function(url: string, body: any, headers: any): Promise<any> {
      return headers ? Promise.reject(new Error('View update failed')) : post(url, body, headers);
    };
    return service.ensureSessionsCalendarView().then(function(): void {
      assert.fail('Expected metadata failure');
    }, function(error: Error): Promise<void> {
      assert.equal(error.message, 'View update failed');
      service.rest.post = post;
      return service.ensureSessionsCalendarView();
    }).then(function(): void {
      assert.lengthOf(posts, 2);
    });
  });

  it('runs calendar setup only after lists and fields are ready and surfaces its errors', function(): Promise<void> {
    var calls: string[] = [];
    service.getCurrentUser = function(): Promise<any> { return Promise.resolve({ isSiteAdmin: true }); };
    service.isOrganizer = function(): Promise<boolean> { return Promise.resolve(true); };
    service.provisionLists = function(): Promise<void> { calls.push('lists'); return Promise.resolve(); };
    service.provisionFields = function(): Promise<void> { calls.push('fields'); return Promise.resolve(); };
    service.ensureSessionsCalendarView = function(): Promise<void> {
      calls.push('calendar');
      return Promise.reject(new Error('Calendar setup failed'));
    };
    return service.provision(0).then(function(): void {
      assert.fail('Expected provisioning failure');
    }, function(error: Error): void {
      assert.equal(error.message, 'Calendar setup failed');
      assert.deepEqual(calls, ['lists', 'fields', 'calendar']);
    });
  });
});

describe('Event Registration Calendar selection', function(): void {
  var webPart: any;

  beforeEach(function(): void {
    webPart = Object.create(EventRegistrationWebPart.prototype);
    webPart.properties = { enableDiagnostics: false };
  });

  it('uses the parent event ID instead of the selected session item ID', function(): void {
    webPart.readCalendarSelection({
      getPropertyDefinitions: function(): any[] { return [{ id: 'selectedEventId' }]; },
      getPropertyValue: function(id: string): number { return id === 'selectedEventId' ? 7 : 42; }
    });
    assert.equal(webPart.dynamicEventId, 7);
    assert.equal(webPart.dynamicSessionId, 42);
  });

  it('does not reinterpret a session ID as an event when the parent is unavailable', function(): void {
    webPart.readCalendarSelection({
      getPropertyDefinitions: function(): any[] { return [{ id: 'selectedEventId' }]; },
      getPropertyValue: function(id: string): number { return id === 'selectedEventId' ? 0 : 42; }
    });
    assert.equal(webPart.dynamicEventId, 0);
  });

  it('preserves selectedItemId compatibility for older Calendar sources', function(): void {
    webPart.readCalendarSelection({
      getPropertyDefinitions: function(): any[] { return [{ id: 'selectedItemId' }]; },
      getPropertyValue: function(id: string): number { return id === 'selectedItemId' ? 7 : 42; }
    });
    assert.equal(webPart.dynamicEventId, 7);
    assert.equal(webPart.dynamicSessionId, 42);
  });

  it('consumes Calendar query parameters, including its return URL', function(): void {
    var originalUrl = window.location.href;
    try {
      window.history.replaceState({}, '', '?eventid=7&sessionid=42&return=%2Fsites%2Fevents%2FSitePages%2FCalendar.aspx');
      assert.deepEqual(webPart.readQuerySelection(), {
        eventId: 7, sessionId: 42, sourceUrl: '/sites/events/SitePages/Calendar.aspx'
      });
    } finally {
      window.history.replaceState({}, '', originalUrl);
    }
  });
});
