import { assert } from 'chai';
import { renderToStaticMarkup } from 'react-dom/server';
import EventRegistration from '../components/EventRegistration';
import { IEventRegistrationProps } from '../components/IEventRegistrationProps';
import { getDisplayMode, getRegistrationPageUrl } from '../helpers/displayMode';
import { IEventItem, ISessionItem } from '../models';

describe('Event Registration display modes', function(): void {
  var component: any;
  var event: IEventItem;
  var session: ISessionItem;
  var commands: { [label: string]: () => void };

  beforeEach(function(): void {
    var props: IEventRegistrationProps = {
      title: 'Events', description: '', webUrl: 'https://example.test', spHttpClient: undefined,
      listPrefix: 'Test Events', organizerGroupId: 0, buttonDisplayMode: 'text', enableDiagnostics: false,
      initialEventId: 0, initialSessionId: 0, sourceUrl: '', onRegistrationSelected: function(): void { return; }
    };
    component = new EventRegistration(props);
    commands = {};
    var renderCommand = component.commandButton;
    component.commandButton = function(label: string, icon: string, onClick: () => void,
      disabled?: boolean, secondary?: boolean, danger?: boolean): React.ReactNode {
      commands[label] = onClick;
      return renderCommand.call(component, label, icon, onClick, disabled, secondary, danger);
    };
    component.setState = function(update: any): void {
      Object.keys(update).forEach(function(key: string): void { component.state[key] = update[key]; });
    };
    event = {
      Id: 10, Title: 'Design workshop', ERDescription: 'Hands-on session', ERLocation: 'Room A',
      ERStatus: 'open', ERRegistrationMode: 'perOccurrence', ERRecurrenceType: 'none',
      ERRecurrenceInterval: 1, ERMaxSeatsPerRegistration: 2
    };
    session = {
      Id: 20, Title: 'Workshop session', EREventId: 10,
      ERStart: new Date(Date.now() + 86400000).toISOString(),
      EREnd: new Date(Date.now() + 90000000).toISOString(),
      ERCapacity: 20, ERWaitlistCapacity: 5, ERStatus: 'open', EROccurrenceKey: 'workshop'
    };
    component.setState({
      initialized: true, loading: false, events: [event], sessions: [session],
      user: { id: 1, title: 'Participant' },
      availability: { '20': { sessionId: 20, available: 12, waitlistCapacity: 5, waitlisted: 2 } }
    });
  });

  function markup(): string {
    return renderToStaticMarkup(component.render());
  }

  it('defaults existing and invalid configurations to Combined', function(): void {
    assert.equal(getDisplayMode(undefined), 'combined');
    assert.equal(getDisplayMode('invalid'), 'combined');
    component.state.organizer = true;
    var html = markup();
    assert.include(html, 'Participant registration');
    assert.include(html, 'Organizer administration');
    assert.include(html, 'Register for this session');
  });

  it('shows only management to authorized organizers and starts on event creation', function(): void {
    var previousProps: IEventRegistrationProps = Object.assign({}, component.props);
    component.props.displayMode = 'organizer';
    component.initialize = function(): void { return; };
    component.componentDidUpdate(previousProps);
    assert.equal(component.state.organizerTab, 'events');
    component.setState({ initialized: true, loading: false, organizer: true });
    var html = markup();
    assert.include(html, 'Organizer administration');
    assert.include(html, 'Create event');
    assert.notInclude(html, 'Participant registration');
  });

  it('never grants organizer access based on display mode', function(): void {
    component.props.displayMode = 'organizer';
    var html = markup();
    assert.include(html, 'You do not have permission to manage events.');
    assert.notInclude(html, 'Organizer administration');
    assert.notInclude(html, 'Create event');
    assert.notInclude(html, 'Participant registration');
  });

  it('shows searchable upcoming published events without organizer controls', function(): void {
    component.props.displayMode = 'catalog';
    component.state.organizer = true;
    var html = markup();
    assert.include(html, 'Design workshop');
    assert.include(html, 'Search events');
    assert.include(html, '12 seats; 3 waitlist seats available');
    assert.include(html, 'View / Register');
    assert.notInclude(html, 'Organizer administration');
    assert.notInclude(html, 'Named attendees');
    component.state.search = 'room a';
    assert.include(markup(), 'Design workshop');
    component.state.search = 'missing';
    assert.include(markup(), 'No upcoming events match your search.');
  });

  it('excludes drafts, closed events and events without upcoming sessions from the catalog', function(): void {
    component.props.displayMode = 'catalog';
    event.ERStatus = 'draft';
    assert.notInclude(markup(), 'Design workshop');
    event.ERStatus = 'closed';
    assert.notInclude(markup(), 'Design workshop');
    event.ERStatus = 'open';
    session.EREnd = new Date(Date.now() - 86400000).toISOString();
    assert.notInclude(markup(), 'Design workshop');
  });

  it('allows browsing published events before registration opens but prevents registration', function(): void {
    component.props.displayMode = 'catalog';
    event.ERRegistrationOpen = new Date(Date.now() + 3600000).toISOString();
    assert.include(markup(), 'Design workshop');
    assert.include(markup(), 'Registration not currently open');
    component.openEvent(10);
    var html = markup();
    assert.include(html, 'Registration opens:');
    assert.notInclude(html, 'Register for this session');
    assert.notInclude(html, 'Named attendees');
  });

  it('opens focused inline details and returns to the filtered catalog', function(): void {
    component.props.displayMode = 'catalog';
    component.state.search = 'room a';
    markup();
    commands['View / Register']();
    var html = markup();
    assert.include(html, 'Back to events');
    assert.include(html, 'Register for this session');
    assert.notInclude(html, 'Search events');
    commands['Register for this session']();
    assert.include(markup(), 'Named attendees');
    component.state.selectedAttendeeIds = [1];
    commands['Back to events']();
    assert.equal(component.state.selectedEventId, 0);
    assert.equal(component.state.selectedSessionId, 0);
    assert.deepEqual(component.state.selectedAttendeeIds, []);
    assert.equal(component.state.search, 'room a');
    assert.include(markup(), 'Search events');
  });

  it('renders only the selected event on a registration page and keeps My registrations', function(): void {
    component.props.displayMode = 'registration';
    component.state.selectedEventId = 10;
    component.state.selectedSessionId = 20;
    var html = markup();
    assert.include(html, 'Design workshop');
    assert.include(html, 'Named attendees');
    assert.include(html, 'My registrations');
    assert.notInclude(html, 'Search events');
    assert.notInclude(html, 'Back to events');
    commands['Close']();
    assert.equal(component.state.selectedEventId, 10);
    assert.equal(component.state.selectedSessionId, 0);
    component.state.participantTab = 'mine';
    assert.include(markup(), 'No registrations found.');
    assert.notInclude(markup(), 'Named attendees');
  });

  it('reports missing selections and unavailable events without exposing drafts', function(): void {
    component.props.displayMode = 'registration';
    assert.include(markup(), 'configure a default event ID');
    component.state.selectedEventId = 999;
    assert.include(markup(), 'This event is unavailable');
    component.state.selectedEventId = 10;
    event.ERStatus = 'draft';
    assert.notInclude(markup(), 'Design workshop');
  });

  it('does not open a registration form for a session belonging to another event', function(): void {
    component.props.displayMode = 'registration';
    component.state.selectedEventId = 10;
    component.state.selectedSessionId = 999;
    assert.notInclude(markup(), 'Named attendees');
  });

  it('supports series selection and hides registration after the window closes', function(): void {
    component.props.displayMode = 'registration';
    component.state.selectedEventId = 10;
    event.ERRegistrationMode = 'series';
    assert.include(markup(), 'Select series');
    assert.notInclude(markup(), 'Register for this session');
    commands['Select series']();
    assert.equal(component.state.selectedSessionId, 20);
    assert.include(markup(), 'Register for series');
    event.ERRegistrationClose = new Date(Date.now() - 3600000).toISOString();
    assert.notInclude(markup(), 'Select series');
    assert.notInclude(markup(), 'Named attendees');
  });

  it('clears stale attendee/session choices when external selections are removed', function(): void {
    var previousProps: IEventRegistrationProps = Object.assign({}, component.props, { initialEventId: 10, initialSessionId: 20 });
    component.setState({ selectedEventId: 10, selectedSessionId: 20, selectedAttendeeIds: [1] });
    component.componentDidUpdate(previousProps);
    assert.equal(component.state.selectedEventId, 0);
    assert.equal(component.state.selectedSessionId, 0);
    assert.deepEqual(component.state.selectedAttendeeIds, []);
  });

  it('loads organizer data only in Combined or management mode', function(): Promise<void> {
    var receivedOrganizer: boolean;
    var receivedRegistrant: number;
    var attendeeRequests = 0;
    component.state.organizer = true;
    component.service = {
      getEvents: function(organizer: boolean): Promise<any[]> { receivedOrganizer = organizer; return Promise.resolve([]); },
      getSessions: function(): Promise<any[]> { return Promise.resolve([]); },
      getSiteUsers: function(): Promise<any[]> { return Promise.resolve([]); },
      getRegistrations: function(userId: number): Promise<any[]> { receivedRegistrant = userId; return Promise.resolve([]); },
      getAttendees: function(): Promise<any[]> { attendeeRequests += 1; return Promise.resolve([]); },
      getAvailability: function(): Promise<any[]> { return Promise.resolve([]); }
    };
    component.props.displayMode = 'catalog';
    return component.reload().then(function(): Promise<void> {
      assert.isFalse(receivedOrganizer);
      assert.equal(receivedRegistrant, 1);
      assert.equal(attendeeRequests, 0);
      component.props.displayMode = 'registration';
      return component.reload();
    }).then(function(): Promise<void> {
      assert.isFalse(receivedOrganizer);
      assert.equal(attendeeRequests, 0);
      component.props.displayMode = 'organizer';
      return component.reload();
    }).then(function(): void {
      assert.isTrue(receivedOrganizer);
      assert.isUndefined(receivedRegistrant);
      assert.equal(attendeeRequests, 1);
    });
  });

  it('does not expose provisioning in participant modes, including for organizers', function(): void {
    component.setState({ initialized: false, organizer: true });
    component.props.displayMode = 'catalog';
    assert.notInclude(markup(), 'Provision lists and fields');
    component.props.displayMode = 'registration';
    assert.notInclude(markup(), 'Provision lists and fields');
    component.props.displayMode = 'organizer';
    assert.include(markup(), 'Provision lists and fields');
  });

  it('preserves the Combined registration close behavior', function(): void {
    component.setState({ selectedEventId: 10, selectedSessionId: 20, selectedAttendeeIds: [1] });
    markup();
    commands['Close']();
    assert.equal(component.state.selectedEventId, 0);
    assert.equal(component.state.selectedSessionId, 0);
    assert.deepEqual(component.state.selectedAttendeeIds, []);
  });

  it('retains the focused event after a successful registration', function(): Promise<void> {
    component.props.displayMode = 'registration';
    component.setState({ selectedEventId: 10, selectedSessionId: 20 });
    component.service = {
      register: function(): Promise<any> {
        return Promise.resolve({ registration: { Id: 100, ERConfirmationNumber: 'CONF-100' }, status: 'confirmed' });
      }
    };
    component.reload = function(): Promise<void> { return Promise.resolve(); };
    component.submitRegistration(event);
    return new Promise<void>(function(resolve): void {
      setTimeout(function(): void {
        assert.equal(component.state.selectedEventId, 10);
        assert.equal(component.state.selectedSessionId, 0);
        assert.include(component.state.message, 'CONF-100');
        resolve();
      }, 0);
    });
  });

  it('builds same-origin registration links preserving page parameters but replacing selections', function(): void {
    var target = new URL(getRegistrationPageUrl('/sites/events/Register.aspx?EventId=1&SessionId=2&view=details#form',
      'https://example.test/sites/events/Events.aspx', 10));
    assert.equal(target.pathname, '/sites/events/Register.aspx');
    assert.equal(target.searchParams.get('eventid'), '10');
    assert.isFalse(target.searchParams.has('SessionId'));
    assert.isFalse(target.searchParams.has('EventId'));
    assert.equal(target.searchParams.get('view'), 'details');
    assert.equal(target.hash, '#form');
  });

  it('rejects external or executable registration page URLs', function(): void {
    ['https://other.test/Register.aspx', '//other.test/Register.aspx', 'javascript:alert(1)'].forEach(function(url: string): void {
      assert.throws(function(): void { getRegistrationPageUrl(url, 'https://example.test/Events.aspx', 10); });
    });
  });
});
