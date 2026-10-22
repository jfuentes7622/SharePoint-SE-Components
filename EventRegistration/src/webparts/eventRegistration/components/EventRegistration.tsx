import * as React from 'react';
import styles from './EventRegistration.module.scss';
import { IEventRegistrationProps } from './IEventRegistrationProps';
import {
  IAttendeeItem,
  IAvailability,
  IEventDraft,
  IEventItem,
  IRegistrationItem,
  IRegistrationResult,
  ISessionItem,
  IUserInfo
} from '../models';
import EventRegistrationService from '../services/EventRegistrationService';

export interface IEventRegistrationState {
  loading: boolean;
  initialized: boolean;
  organizer: boolean;
  user?: IUserInfo;
  users: IUserInfo[];
  events: IEventItem[];
  sessions: ISessionItem[];
  availability: { [sessionId: string]: IAvailability };
  registrations: IRegistrationItem[];
  attendees: IAttendeeItem[];
  participantTab: string;
  organizerTab: string;
  selectedEventId: number;
  selectedSessionId: number;
  selectedAttendeeIds: number[];
  message: string;
  error: string;
  lastResult?: IRegistrationResult;
  draft: IEventDraft;
  capacityEdits: { [sessionId: string]: string };
  waitlistEdits: { [sessionId: string]: string };
}

function emptyDraft(userId?: number): IEventDraft {
  var now = new Date();
  var later = new Date(now.getTime() + 3600000);
  var month = new Date(now.getTime() + 30 * 86400000);
  return {
    title: '',
    description: '',
    location: '',
    organizerId: userId || 0,
    status: 'draft',
    registrationMode: 'perOccurrence',
    recurrenceType: 'none',
    recurrenceInterval: 1,
    recurrenceDays: [],
    recurrenceUntil: toInputValue(month),
    registrationOpen: toInputValue(now),
    registrationClose: toInputValue(month),
    maxSeatsPerRegistration: 1,
    firstStart: toInputValue(now),
    firstEnd: toInputValue(later),
    capacity: 20,
    waitlistCapacity: 5
  };
}

function toInputValue(date: Date): string {
  var adjusted = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return adjusted.toISOString().slice(0, 16);
}

export default class EventRegistration extends React.Component<IEventRegistrationProps, IEventRegistrationState> {
  private service: EventRegistrationService;

  constructor(props: IEventRegistrationProps) {
    super(props);
    this.service = new EventRegistrationService(props.webUrl, props.spHttpClient, props.listPrefix,
      props.enableDiagnostics !== false);
    this.state = {
      loading: true,
      initialized: false,
      organizer: false,
      users: [],
      events: [],
      sessions: [],
      availability: {},
      registrations: [],
      attendees: [],
      participantTab: 'browse',
      organizerTab: 'setup',
      selectedEventId: props.initialEventId || 0,
      selectedSessionId: props.initialSessionId || 0,
      selectedAttendeeIds: [],
      message: '',
      error: '',
      draft: emptyDraft(),
      capacityEdits: {},
      waitlistEdits: {}
    };
  }

  public componentDidMount(): void {
    this.initialize();
  }

  public componentWillReceiveProps(nextProps: IEventRegistrationProps): void {
    if (nextProps.listPrefix !== this.props.listPrefix || nextProps.enableDiagnostics !== this.props.enableDiagnostics) {
      this.service = new EventRegistrationService(nextProps.webUrl, nextProps.spHttpClient, nextProps.listPrefix,
        nextProps.enableDiagnostics !== false);
      this.initialize();
    }
    if (nextProps.initialEventId !== this.props.initialEventId || nextProps.initialSessionId !== this.props.initialSessionId) {
      this.setState({
        selectedEventId: nextProps.initialEventId || this.state.selectedEventId,
        selectedSessionId: nextProps.initialSessionId || this.state.selectedSessionId
      });
    }
  }

  public render(): React.ReactElement<IEventRegistrationProps> {
    return (
      <div className={styles.eventRegistration}>
        <div className={styles.header}>
          <div>
            <h2>{this.props.title}</h2>
            {this.props.description && <p className={styles.description}>{this.props.description}</p>}
          </div>
          <span>{this.state.user ? 'Signed in as ' + this.state.user.title : ''}</span>
        </div>
        {this.renderMessages()}
        {this.state.loading && <div className={styles.message} role="status">Loading…</div>}
        {!this.state.initialized && !this.state.loading && this.renderNotSetup()}
        {this.state.initialized && this.renderParticipant()}
        {this.state.initialized && this.state.organizer && this.renderOrganizer()}
      </div>
    );
  }

  private renderMessages(): React.ReactNode {
    return (
      <div>
        {this.state.error && <div className={styles.message + ' ' + styles.error} role="alert">{this.state.error}</div>}
        {this.state.message && <div className={styles.message + ' ' + styles.success} role="status">{this.state.message}</div>}
      </div>
    );
  }

  private renderNotSetup(): React.ReactNode {
    return (
      <div className={styles.panel}>
        <h3>Setup required</h3>
        <p>The configured Event Registration lists are unavailable. A Site Collection Administrator
          {this.props.organizerGroupId > 0 ? ' or configured organizer' : ''} must run setup.</p>
        {this.state.organizer && this.commandButton('Provision lists and fields', 'Settings', this.runSetup.bind(this))}
      </div>
    );
  }

  private renderParticipant(): React.ReactNode {
    return (
      <section aria-label="Participant registration">
        <div className={styles.tabs} role="tablist" aria-label="Participant views">
          {this.tabButton('browse', 'Browse events', false)}
          {this.tabButton('mine', 'My registrations', false)}
        </div>
        {this.state.participantTab === 'browse' ? this.renderBrowse() : this.renderMyRegistrations()}
      </section>
    );
  }

  private renderBrowse(): React.ReactNode {
    var currentTime = new Date().getTime();
    var openEvents = this.state.events.filter(function(event: IEventItem): boolean {
      return event.ERStatus === 'open'
        && (!event.ERRegistrationOpen || new Date(event.ERRegistrationOpen).getTime() <= currentTime)
        && (!event.ERRegistrationClose || new Date(event.ERRegistrationClose).getTime() >= currentTime);
    });
    if (openEvents.length === 0) {
      return <div className={styles.panel}>There are no open events.</div>;
    }
    return (
      <div>
        {openEvents.map((event: IEventItem): React.ReactNode => {
          var sessions = this.sessionsForEvent(event.Id);
          return (
            <article className={styles.card} key={event.Id}>
              <h3>{event.Title}</h3>
              <p>{event.ERDescription || ''}</p>
              <div className={styles.meta}>{event.ERLocation || 'Location not specified'} · Registration: {
                event.ERRegistrationMode === 'series' ? 'entire series' : 'per occurrence'
              }</div>
              {sessions.map((session: ISessionItem): React.ReactNode => this.renderSessionChoice(event, session))}
              {event.ERRegistrationMode === 'series' && sessions.length > 0 && (
                <div className={styles.actions}>
                  {this.commandButton('Select series', 'Calendar', (): void => {
                    this.setState({ selectedEventId: event.Id, selectedSessionId: sessions[0].Id, selectedAttendeeIds: [] });
                  })}
                </div>
              )}
              {this.state.selectedEventId === event.Id && this.renderRegistrationForm(event)}
            </article>
          );
        })}
      </div>
    );
  }

  private renderSessionChoice(event: IEventItem, session: ISessionItem): React.ReactNode {
    var availability = this.state.availability[String(session.Id)];
    return (
      <div className={styles.card} key={session.Id}>
        <strong>{new Date(session.ERStart).toLocaleString()}</strong>
        <span> – {new Date(session.EREnd).toLocaleString()}</span>
        <div className={styles.meta}>
          {availability
            ? String(availability.available) + ' confirmed seats available; '
              + String(Math.max(0, availability.waitlistCapacity - availability.waitlisted)) + ' waitlist seats available'
            : 'Availability loading'}
        </div>
        {event.ERRegistrationMode !== 'series' && this.commandButton('Register for this session', 'AddFriend', (): void => {
          this.setState({ selectedEventId: event.Id, selectedSessionId: session.Id, selectedAttendeeIds: [] });
        })}
      </div>
    );
  }

  private renderRegistrationForm(event: IEventItem): React.ReactNode {
    var max = Number(event.ERMaxSeatsPerRegistration || 1);
    return (
      <div className={styles.panel}>
        <h4>Named attendees ({String(this.state.selectedAttendeeIds.length)} of {String(max)})</h4>
        <div className={styles.userList} role="group" aria-label="Select named attendees">
          {this.state.users.map((user: IUserInfo): React.ReactNode => {
            var selected = this.state.selectedAttendeeIds.indexOf(user.id) >= 0;
            return (
              <label className={styles.userOption} key={user.id}>
                <input type="checkbox" checked={selected} onChange={(): void => this.toggleAttendee(user.id, max)} />
                {' ' + user.title + (user.email ? ' (' + user.email + ')' : '')}
              </label>
            );
          })}
        </div>
        <div className={styles.actions}>
          {this.commandButton(event.ERRegistrationMode === 'series' ? 'Register for series' : 'Register',
            'Accept', (): void => this.submitRegistration(event), this.state.selectedAttendeeIds.length === 0)}
          {this.commandButton('Close', 'Cancel', (): void => this.setState({ selectedEventId: 0 }), false, true)}
        </div>
      </div>
    );
  }

  private renderMyRegistrations(): React.ReactNode {
    var mine = this.state.registrations.filter((registration: IRegistrationItem): boolean => {
      return !!this.state.user && Number(registration.ERRegistrantId) === Number(this.state.user.id);
    });
    return (
      <div className={styles.panel}>
        <h3>My Registrations</h3>
        {mine.length === 0 && <p>No registrations found.</p>}
        {mine.map((registration: IRegistrationItem): React.ReactNode => (
          <div className={styles.card} key={registration.Id}>
            <strong>{registration.ERConfirmationNumber}</strong>
            <div className={styles.meta}>{registration.ERScope} · {registration.ERStatus} · {
              new Date(registration.ERCreatedAt).toLocaleString()
            }</div>
            {registration.ERStatus !== 'cancelled' && this.commandButton('Cancel registration', 'Cancel',
              (): void => this.cancelRegistration(registration), false, false, true)}
          </div>
        ))}
      </div>
    );
  }

  private renderOrganizer(): React.ReactNode {
    return (
      <section aria-label="Organizer administration">
        <h2>Organizer</h2>
        <div className={styles.tabs} role="tablist" aria-label="Organizer views">
          {this.tabButton('setup', 'Setup', true)}
          {this.tabButton('events', 'Events', true)}
          {this.tabButton('sessions', 'Sessions', true)}
          {this.tabButton('registrations', 'Registrations / Attendees', true)}
          {this.tabButton('waitlist', 'Waitlist', true)}
          {this.tabButton('export', 'CSV Export', true)}
        </div>
        {this.state.organizerTab === 'setup' && this.renderSetup()}
        {this.state.organizerTab === 'events' && this.renderEventsAdmin()}
        {this.state.organizerTab === 'sessions' && this.renderSessionsAdmin()}
        {this.state.organizerTab === 'registrations' && this.renderRegistrationsAdmin(false)}
        {this.state.organizerTab === 'waitlist' && this.renderRegistrationsAdmin(true)}
        {this.state.organizerTab === 'export' && this.renderExport()}
      </section>
    );
  }

  private renderSetup(): React.ReactNode {
    return (
      <div className={styles.panel}>
        <h3>Schema setup and validation</h3>
        <p>Creates only missing lists and fields. Existing field types are validated; incompatible schemas are reported
          and never replaced.</p>
        {this.commandButton('Validate / provision', 'Settings', this.runSetup.bind(this))}
      </div>
    );
  }

  private renderEventsAdmin(): React.ReactNode {
    var draft = this.state.draft;
    return (
      <div className={styles.panel}>
        <h3>{draft.id ? 'Edit event' : 'Create event'}</h3>
        <div className={styles.grid}>
          {this.textField('Title', draft.title, (value: string): void => this.updateDraft('title', value))}
          {this.textField('Location', draft.location, (value: string): void => this.updateDraft('location', value))}
          {this.selectField('Organizer', String(draft.organizerId), this.state.users.map(function(user: IUserInfo): any {
            return { value: String(user.id), text: user.title };
          }), (value: string): void => this.updateDraft('organizerId', Number(value)))}
          {this.selectField('Status', draft.status, [
            { value: 'draft', text: 'Draft' }, { value: 'open', text: 'Open' },
            { value: 'closed', text: 'Closed' }, { value: 'cancelled', text: 'Cancelled' }
          ], (value: string): void => this.updateDraft('status', value))}
          {this.selectField('Registration mode', draft.registrationMode, [
            { value: 'perOccurrence', text: 'Per occurrence' }, { value: 'series', text: 'Entire series' }
          ], (value: string): void => this.updateDraft('registrationMode', value))}
          {this.selectField('Recurrence', draft.recurrenceType, [
            { value: 'none', text: 'None' }, { value: 'daily', text: 'Daily' }, { value: 'weekly', text: 'Weekly' }
          ], (value: string): void => this.updateDraft('recurrenceType', value))}
          {this.numberField('Recurrence interval', draft.recurrenceInterval,
            (value: number): void => this.updateDraft('recurrenceInterval', value), 1)}
          {this.textField('Weekly days (0=Sun…6=Sat)', draft.recurrenceDays.join(','),
            (value: string): void => this.updateDraft('recurrenceDays', value.split(',').map(function(day: string): string {
              return day.trim();
            }).filter(function(day: string): boolean { return day !== ''; })))}
          {this.dateField('Recurrence until', draft.recurrenceUntil,
            (value: string): void => this.updateDraft('recurrenceUntil', value))}
          {this.dateField('Registration opens', draft.registrationOpen,
            (value: string): void => this.updateDraft('registrationOpen', value))}
          {this.dateField('Registration closes', draft.registrationClose,
            (value: string): void => this.updateDraft('registrationClose', value))}
          {this.numberField('Maximum seats per registration', draft.maxSeatsPerRegistration,
            (value: number): void => this.updateDraft('maxSeatsPerRegistration', value), 1)}
          {this.dateField('First session starts', draft.firstStart,
            (value: string): void => this.updateDraft('firstStart', value))}
          {this.dateField('First session ends', draft.firstEnd,
            (value: string): void => this.updateDraft('firstEnd', value))}
          {this.numberField('Session capacity', draft.capacity,
            (value: number): void => this.updateDraft('capacity', value), 0)}
          {this.numberField('Waitlist capacity', draft.waitlistCapacity,
            (value: number): void => this.updateDraft('waitlistCapacity', value), 0)}
        </div>
        <div className={styles.field}>
          <label>Description</label>
          <textarea aria-label="Event description" value={draft.description}
            onChange={(event: React.FormEvent<HTMLTextAreaElement>): void =>
              this.updateDraft('description', (event.target as HTMLTextAreaElement).value)} />
        </div>
        <div className={styles.actions}>
          {this.commandButton('Save event and generate sessions', 'Save', this.saveEvent.bind(this))}
          {draft.id && this.commandButton('Clear form', 'Clear', (): void => this.setState({
            draft: emptyDraft(this.state.user ? this.state.user.id : 0)
          }), false, true)}
        </div>
        <h3>Existing events</h3>
        {this.state.events.map((event: IEventItem): React.ReactNode => (
          <div className={styles.card} key={event.Id}>
            <strong>{event.Title}</strong> <span className={styles.meta}>{event.ERStatus} · {event.ERRegistrationMode}</span>
            {this.commandButton('Edit', 'Edit', (): void => this.editEvent(event), false, true)}
          </div>
        ))}
      </div>
    );
  }

  private renderSessionsAdmin(): React.ReactNode {
    return (
      <div className={styles.panel}>
        <h3>Sessions and capacity</h3>
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead><tr><th>Session</th><th>Start</th><th>Capacity</th><th>Waitlist</th><th>Action</th></tr></thead>
            <tbody>
              {this.state.sessions.map((session: ISessionItem): React.ReactNode => (
                <tr key={session.Id}>
                  <td>{session.Title}</td>
                  <td>{new Date(session.ERStart).toLocaleString()}</td>
                  <td><input aria-label={'Capacity for ' + session.Title} type="number" min="0"
                    value={this.state.capacityEdits[String(session.Id)] !== undefined
                      ? this.state.capacityEdits[String(session.Id)] : String(session.ERCapacity)}
                    onChange={(event: React.FormEvent<HTMLInputElement>): void =>
                      this.setEdit('capacityEdits', session.Id, (event.target as HTMLInputElement).value)} /></td>
                  <td><input aria-label={'Waitlist capacity for ' + session.Title} type="number" min="0"
                    value={this.state.waitlistEdits[String(session.Id)] !== undefined
                      ? this.state.waitlistEdits[String(session.Id)] : String(session.ERWaitlistCapacity)}
                    onChange={(event: React.FormEvent<HTMLInputElement>): void =>
                      this.setEdit('waitlistEdits', session.Id, (event.target as HTMLInputElement).value)} /></td>
                  <td>{this.commandButton('Save capacity', 'Save', (): void => this.saveCapacity(session), false, true)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
  }

  private renderRegistrationsAdmin(waitlistOnly: boolean): React.ReactNode {
    var registrations = this.state.registrations.filter(function(registration: IRegistrationItem): boolean {
      return !waitlistOnly || registration.ERStatus === 'waitlisted';
    });
    return (
      <div className={styles.panel}>
        <h3>{waitlistOnly ? 'Waitlist (FIFO by created timestamp)' : 'Registrations and named attendees'}</h3>
        {registrations.map((registration: IRegistrationItem): React.ReactNode => (
          <div className={styles.card} key={registration.Id}>
            <strong>{registration.ERConfirmationNumber}</strong>
            <div className={styles.meta}>{registration.ERStatus} · {registration.ERScope} · seats {registration.ERSeatCount}
              {' · ' + new Date(registration.ERCreatedAt).toLocaleString()}</div>
            <ul>
              {this.state.attendees.filter(function(attendee: IAttendeeItem): boolean {
                return Number(attendee.ERRegistrationId) === Number(registration.Id);
              }).map(function(attendee: IAttendeeItem): React.ReactNode {
                return <li key={attendee.Id}>{attendee.ERAttendee ? attendee.ERAttendee.title : attendee.ERAttendeeKey}
                  {' · ' + attendee.ERStatus + ' · session ' + String(attendee.ERSessionId || '')}</li>;
              })}
            </ul>
            {registration.ERStatus !== 'cancelled' && this.commandButton('Cancel', 'Cancel',
              (): void => this.cancelRegistration(registration), false, false, true)}
          </div>
        ))}
      </div>
    );
  }

  private renderExport(): React.ReactNode {
    return (
      <div className={styles.panel}>
        <h3>CSV Export</h3>
        <p>Exports registration and named-attendee data currently loaded in this organizer view.</p>
        {this.commandButton('Download CSV', 'ExcelDocument', this.downloadCsv.bind(this))}
      </div>
    );
  }

  private initialize(): void {
    this.setState({ loading: true, initialized: false, error: '', message: '' });
    Promise.all([this.service.getCurrentUser(), this.service.isOrganizer(this.props.organizerGroupId)])
      .then((parts: any[]): Promise<void> => {
        var user = parts[0];
        this.setState({
          user: user,
          organizer: parts[1],
          draft: emptyDraft(user.id)
        });
        return this.service.hasRequiredLists().then((configured: boolean): Promise<void> | void => {
          if (!configured) {
            this.setState({ loading: false });
            return;
          }
          return this.reload(parts[1], user.id);
        });
      }).catch((error: any): void => {
        this.setState({
          loading: false,
          initialized: false,
          error: this.errorMessage(error)
        });
      });
  }

  private reload(organizerOverride?: boolean, userIdOverride?: number): Promise<void> {
    var organizer = organizerOverride === undefined ? this.state.organizer : organizerOverride;
    var userId = userIdOverride === undefined ? (this.state.user ? this.state.user.id : 0) : userIdOverride;
    return Promise.all([
      this.service.getEvents(organizer),
      this.service.getSessions(undefined, organizer),
      this.service.getSiteUsers(),
      this.service.getRegistrations(organizer ? undefined : userId),
      organizer ? this.service.getAttendees() : Promise.resolve([])
    ]).then((parts: any[]): Promise<void> => {
      var sessions: ISessionItem[] = parts[1];
      return this.service.getAvailability(sessions).then((availability: IAvailability[]): void => {
        var map: { [sessionId: string]: IAvailability } = {};
        availability.forEach(function(item: IAvailability): void { map[String(item.sessionId)] = item; });
        this.setState({
          loading: false,
          initialized: true,
          events: parts[0],
          sessions: sessions,
          users: parts[2],
          registrations: parts[3],
          attendees: parts[4],
          availability: map
        });
      });
    });
  }

  private runSetup(): void {
    this.setState({ loading: true, error: '', message: '' });
    this.service.provision(this.props.organizerGroupId).then((result: any): Promise<void> => {
      this.setState({
        message: 'Setup complete. Created ' + String(result.createdLists.length) + ' lists and '
          + String(result.createdFields.length) + ' fields; validated ' + String(result.validatedLists.length) + ' lists.'
      });
      return this.reload();
    }).catch((error: any): void => {
      this.setState({ loading: false, error: this.errorMessage(error) });
    });
  }

  private submitRegistration(event: IEventItem): void {
    var session = this.state.sessions.filter((item: ISessionItem): boolean => item.Id === this.state.selectedSessionId)[0];
    var selectedUsers = this.state.users.filter((user: IUserInfo): boolean => this.state.selectedAttendeeIds.indexOf(user.id) >= 0);
    if (!this.state.user) {
      this.setState({ error: 'Current user information is unavailable.' });
      return;
    }
    this.setState({ loading: true, error: '', message: '' });
    this.service.register({
      event: event,
      session: session,
      sessions: this.sessionsForEvent(event.Id),
      registrant: this.state.user,
      attendees: selectedUsers,
      scope: event.ERRegistrationMode
    }).then((result: IRegistrationResult): Promise<void> => {
      this.props.onRegistrationSelected({
        id: result.registration.Id,
        status: result.status,
        confirmation: result.registration.ERConfirmationNumber
      });
      this.setState({
        lastResult: result,
        selectedEventId: 0,
        selectedAttendeeIds: [],
        message: 'Registration ' + result.registration.ERConfirmationNumber + ' is ' + result.status + '.'
      });
      this.redirectIfRequested(result);
      return this.reload();
    }).catch((error: any): void => {
      this.setState({ loading: false, error: this.errorMessage(error) });
    });
  }

  private cancelRegistration(registration: IRegistrationItem): void {
    if (!this.state.user) {
      return;
    }
    this.setState({ loading: true, error: '', message: '' });
    this.service.cancelRegistration(registration, this.state.user, this.state.organizer).then((): Promise<void> => {
      this.props.onRegistrationSelected({
        id: registration.Id,
        status: 'cancelled',
        confirmation: registration.ERConfirmationNumber
      });
      this.setState({ message: 'Registration ' + registration.ERConfirmationNumber + ' was cancelled.' });
      return this.reload();
    }).catch((error: any): void => this.setState({ loading: false, error: this.errorMessage(error) }));
  }

  private saveEvent(): void {
    this.setState({ loading: true, error: '', message: '' });
    this.service.saveEvent(this.state.draft).then((event: IEventItem): Promise<any> => {
      return this.service.generateSessionsForEvent(event, this.state.draft).then((summary: any): any => {
        return { event: event, summary: summary };
      });
    }).then((result: any): Promise<void> => {
      this.setState({
        draft: emptyDraft(this.state.user ? this.state.user.id : 0),
        message: 'Saved ' + result.event.Title + '. Created ' + String(result.summary.created)
          + ' sessions; ' + String(result.summary.existing) + ' deterministic occurrences already existed.'
      });
      return this.reload();
    }).catch((error: any): void => this.setState({ loading: false, error: this.errorMessage(error) }));
  }

  private saveCapacity(session: ISessionItem): void {
    var capacity = Number(this.state.capacityEdits[String(session.Id)] !== undefined
      ? this.state.capacityEdits[String(session.Id)] : session.ERCapacity);
    var waitlist = Number(this.state.waitlistEdits[String(session.Id)] !== undefined
      ? this.state.waitlistEdits[String(session.Id)] : session.ERWaitlistCapacity);
    this.setState({ loading: true, error: '', message: '' });
    this.service.updateSessionCapacity(session, capacity, waitlist).then((): Promise<void> => {
      this.setState({ message: 'Capacity updated for ' + session.Title + '.' });
      return this.reload();
    }).catch((error: any): void => this.setState({ loading: false, error: this.errorMessage(error) }));
  }

  private editEvent(event: IEventItem): void {
    var session = this.sessionsForEvent(event.Id)[0];
    this.setState({
      draft: {
        id: event.Id,
        title: event.Title,
        description: event.ERDescription || '',
        location: event.ERLocation || '',
        organizerId: event.EROrganizerId || (this.state.user ? this.state.user.id : 0),
        status: event.ERStatus,
        registrationMode: event.ERRegistrationMode,
        recurrenceType: event.ERRecurrenceType,
        recurrenceInterval: event.ERRecurrenceInterval || 1,
        recurrenceDays: String(event.ERRecurrenceDays || '').split(',').filter(function(value: string): boolean { return !!value; }),
        recurrenceUntil: event.ERRecurrenceUntil ? toInputValue(new Date(event.ERRecurrenceUntil)) : '',
        registrationOpen: event.ERRegistrationOpen ? toInputValue(new Date(event.ERRegistrationOpen)) : '',
        registrationClose: event.ERRegistrationClose ? toInputValue(new Date(event.ERRegistrationClose)) : '',
        maxSeatsPerRegistration: event.ERMaxSeatsPerRegistration || 1,
        firstStart: session ? toInputValue(new Date(session.ERStart)) : toInputValue(new Date()),
        firstEnd: session ? toInputValue(new Date(session.EREnd)) : toInputValue(new Date(Date.now() + 3600000)),
        capacity: session ? session.ERCapacity : 20,
        waitlistCapacity: session ? session.ERWaitlistCapacity : 5
      }
    });
  }

  private toggleAttendee(id: number, max: number): void {
    var selected = this.state.selectedAttendeeIds.slice();
    var index = selected.indexOf(id);
    if (index >= 0) {
      selected.splice(index, 1);
    } else if (selected.length < max) {
      selected.push(id);
    } else {
      this.setState({ error: 'This event permits at most ' + String(max) + ' named attendees.' });
      return;
    }
    this.setState({ selectedAttendeeIds: selected, error: '' });
  }

  private sessionsForEvent(eventId: number): ISessionItem[] {
    return this.state.sessions.filter(function(session: ISessionItem): boolean {
      return Number(session.EREventId) === Number(eventId) && session.ERStatus === 'open'
        && new Date(session.EREnd).getTime() >= new Date().getTime();
    });
  }

  private updateDraft(key: string, value: any): void {
    var draft: any = {};
    Object.keys(this.state.draft).forEach((name: string): void => { draft[name] = (this.state.draft as any)[name]; });
    draft[key] = value;
    this.setState({ draft: draft as IEventDraft });
  }

  private setEdit(field: string, sessionId: number, value: string): void {
    var map: any = {};
    var current: any = (this.state as any)[field];
    Object.keys(current).forEach(function(key: string): void { map[key] = current[key]; });
    map[String(sessionId)] = value;
    var update: any = {};
    update[field] = map;
    this.setState(update);
  }

  private tabButton(key: string, label: string, organizer: boolean): React.ReactNode {
    var selected = organizer ? this.state.organizerTab === key : this.state.participantTab === key;
    return (
      <button type="button" className={styles.tab} role="tab" aria-selected={selected}
        aria-label={label} title={label} onClick={(): void => {
          if (organizer) {
            this.setState({ organizerTab: key });
          } else {
            this.setState({ participantTab: key });
          }
        }}>{label}</button>
    );
  }

  private commandButton(
    label: string,
    iconName: string,
    onClick: () => void,
    disabled?: boolean,
    secondary?: boolean,
    danger?: boolean
  ): React.ReactNode {
    var className = styles.button;
    if (secondary) {
      className += ' ' + styles.secondary;
    }
    if (danger) {
      className += ' ' + styles.danger;
    }
    if (this.props.buttonDisplayMode === 'icon') {
      className += ' ' + styles.commandButtonIconOnly;
    }
    return (
      <button type="button" className={className} title={label} aria-label={label}
        disabled={disabled === true} onClick={onClick}>
        {this.renderCommandContent(iconName, label)}
      </button>
    );
  }

  private renderCommandContent(iconName: string, label: string): React.ReactNode {
    if (this.props.buttonDisplayMode === 'text') {
      return label;
    }
    return (
      <span className={styles.commandContent}>
        <i className={'ms-Icon ms-Icon--' + iconName} aria-hidden="true"></i>
        <span className={this.props.buttonDisplayMode === 'icon' ? styles.visuallyHidden : ''}>{label}</span>
      </span>
    );
  }

  private textField(label: string, value: string, onChange: (value: string) => void): React.ReactNode {
    return <div className={styles.field}><label>{label}</label><input aria-label={label} type="text" value={value || ''}
      onChange={(event: React.FormEvent<HTMLInputElement>): void => onChange((event.target as HTMLInputElement).value)} /></div>;
  }

  private numberField(label: string, value: number, onChange: (value: number) => void, min: number): React.ReactNode {
    return <div className={styles.field}><label>{label}</label><input aria-label={label} type="number" min={min}
      value={String(value)} onChange={(event: React.FormEvent<HTMLInputElement>): void =>
        onChange(Number((event.target as HTMLInputElement).value))} /></div>;
  }

  private dateField(label: string, value: string, onChange: (value: string) => void): React.ReactNode {
    return <div className={styles.field}><label>{label}</label><input aria-label={label} type="datetime-local"
      value={value || ''} onChange={(event: React.FormEvent<HTMLInputElement>): void =>
        onChange((event.target as HTMLInputElement).value)} /></div>;
  }

  private selectField(
    label: string,
    value: string,
    options: Array<{ value: string; text: string }>,
    onChange: (value: string) => void
  ): React.ReactNode {
    return (
      <div className={styles.field}><label>{label}</label>
        <select aria-label={label} value={value} onChange={(event: React.FormEvent<HTMLSelectElement>): void =>
          onChange((event.target as HTMLSelectElement).value)}>
          {options.map(function(option: any): React.ReactNode {
            return <option key={option.value} value={option.value}>{option.text}</option>;
          })}
        </select>
      </div>
    );
  }

  private downloadCsv(): void {
    var header = ['Registration ID', 'Confirmation', 'Registration Status', 'Scope', 'Created', 'Registrant',
      'Attendee', 'Attendee Email', 'Attendee Status', 'Event ID', 'Session ID', 'Slot ID'];
    var rows: string[][] = [header];
    this.state.registrations.forEach((registration: IRegistrationItem): void => {
      var attendees = this.state.attendees.filter(function(attendee: IAttendeeItem): boolean {
        return Number(attendee.ERRegistrationId) === Number(registration.Id);
      });
      if (attendees.length === 0) {
        attendees = [{} as IAttendeeItem];
      }
      attendees.forEach(function(attendee: IAttendeeItem): void {
        rows.push([
          String(registration.Id), registration.ERConfirmationNumber, registration.ERStatus, registration.ERScope,
          registration.ERCreatedAt, registration.ERRegistrant ? registration.ERRegistrant.title : '',
          attendee.ERAttendee ? attendee.ERAttendee.title : attendee.ERAttendeeKey || '',
          attendee.ERAttendee ? attendee.ERAttendee.email : '', attendee.ERStatus || '',
          String(registration.EREventId || ''), String(attendee.ERSessionId || ''), String(attendee.ERSlotId || '')
        ]);
      });
    });
    var csv = rows.map(function(row: string[]): string {
      return row.map(function(value: string): string { return '"' + String(value || '').replace(/"/g, '""') + '"'; }).join(',');
    }).join('\r\n');
    var blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'event-registration-' + new Date().toISOString().slice(0, 10) + '.csv';
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
  }

  private redirectIfRequested(result: IRegistrationResult): void {
    if (!this.props.sourceUrl || typeof window === 'undefined') {
      return;
    }
    try {
      var target = new URL(this.props.sourceUrl, window.location.href);
      if (target.origin !== window.location.origin) {
        this.setState({ error: 'The source/return URL was ignored because it is not same-origin.' });
        return;
      }
      target.searchParams.set('registrationid', String(result.registration.Id));
      target.searchParams.set('status', result.status);
      window.location.href = target.href;
    } catch (error) {
      this.setState({ error: 'The source/return URL is invalid and was not used.' });
    }
  }

  private errorMessage(error: any): string {
    var message = error && (error.message || error.statusText) ? String(error.message || error.statusText) : String(error);
    console.error('[EventRegistration] ' + message);
    return message;
  }
}
