import { SPHttpClient } from '@microsoft/sp-http';
import {
  IAttendeeItem,
  IAvailability,
  IEventDraft,
  IEventItem,
  IRegistrationItem,
  IRegistrationRequest,
  IRegistrationResult,
  ISessionItem,
  ISetupResult,
  ISlotItem,
  IUserInfo
} from '../models';
import { generateOccurrences } from '../helpers/recurrence';
import { validateCapacityDecrease } from '../helpers/capacity';
import SharePointRestClient, { IRestResult } from './SharePointRestClient';
import { IFieldSchema, IListSchema, listSchemas, listTitle } from './schema';

interface IClaimRecord {
  slot: ISlotItem;
  attendee: IAttendeeItem;
}

export default class EventRegistrationService {
  private rest: SharePointRestClient;
  private prefix: string;
  private entityTypes: { [key: string]: string } = {};
  private listIds: { [key: string]: string } = {};

  constructor(webUrl: string, client: SPHttpClient, prefix: string, enableDiagnostics: boolean) {
    this.rest = new SharePointRestClient(webUrl, client, enableDiagnostics);
    this.prefix = String(prefix || 'SPS Event Registration').trim() || 'SPS Event Registration';
  }

  public setDiagnostics(enabled: boolean): void {
    this.rest.setDiagnostics(enabled);
  }

  public hasRequiredLists(): Promise<boolean> {
    var titles = listSchemas.map((schema: IListSchema): string => listTitle(this.prefix, schema.key));
    var filter = titles.map((title: string): string => "Title eq '" + this.rest.escape(title) + "'").join(' or ');
    return this.rest.get('web/lists?$select=Title&$filter=' + encodeURIComponent(filter)
      + '&$top=' + String(titles.length)).then((result: IRestResult): boolean => {
        var lists = this.rest.rows(result.data);
        return titles.every(function(title: string): boolean {
          return lists.some(function(list: any): boolean { return list.Title === title; });
        });
      });
  }

  public getCurrentUser(): Promise<IUserInfo & { isSiteAdmin: boolean }> {
    return this.rest.get('web/currentuser?$select=Id,Title,Email,LoginName,IsSiteAdmin').then((result: IRestResult): any => {
      var item = this.rest.item(result.data);
      return {
        id: Number(item.Id),
        title: String(item.Title || item.LoginName || ''),
        email: String(item.Email || ''),
        loginName: String(item.LoginName || ''),
        isSiteAdmin: item.IsSiteAdmin === true
      };
    });
  }

  public isOrganizer(groupId: number): Promise<boolean> {
    return this.getCurrentUser().then((user: any): Promise<boolean> | boolean => {
      if (user.isSiteAdmin) {
        return true;
      }
      if (!groupId || groupId <= 0) {
        return false;
      }
      return this.rest.get("web/sitegroups/getById(" + String(groupId) + ")/users?$select=Id")
        .then((result: IRestResult): boolean => {
          return this.rest.rows(result.data).some(function(member: any): boolean {
            return Number(member.Id) === Number(user.id);
          });
        });
    });
  }

  public getSiteGroups(): Promise<Array<{ key: number; text: string }>> {
    return this.rest.get('web/sitegroups?$select=Id,Title&$orderby=Title').then((result: IRestResult): any[] => {
      return this.rest.rows(result.data).map(function(group: any): any {
        return { key: Number(group.Id), text: String(group.Title || '') };
      });
    });
  }

  public getSiteUsers(): Promise<IUserInfo[]> {
    return this.rest.get("web/siteusers?$select=Id,Title,Email,LoginName&$filter=PrincipalType eq 1&$orderby=Title")
      .then((result: IRestResult): IUserInfo[] => {
        return this.rest.rows(result.data).map(function(user: any): IUserInfo {
          return {
            id: Number(user.Id),
            title: String(user.Title || user.LoginName || ''),
            email: String(user.Email || ''),
            loginName: String(user.LoginName || '')
          };
        }).filter(function(user: IUserInfo): boolean { return user.id > 0; });
      });
  }

  public provision(groupId: number): Promise<ISetupResult> {
    return this.getCurrentUser().then((user: any): Promise<ISetupResult> => {
      if (!user.isSiteAdmin && (!groupId || groupId <= 0)) {
        throw new Error('Only a Site Collection Administrator can run setup until an organizer group is configured.');
      }
      return this.isOrganizer(groupId).then((allowed: boolean): Promise<ISetupResult> => {
        if (!allowed) {
          throw new Error('Setup requires Site Collection Administrator or configured organizer group membership.');
        }
        var result: ISetupResult = { createdLists: [], createdFields: [], validatedLists: [], errors: [] };
        return this.provisionLists(0, result).then((): Promise<ISetupResult> => {
          return this.provisionFields(0, result);
        }).then((): ISetupResult => {
          if (result.errors.length > 0) {
            throw new Error('Setup did not complete: ' + result.errors.join(' | '));
          }
          return result;
        });
      });
    });
  }

  public getEvents(includeAll?: boolean): Promise<IEventItem[]> {
    var filter = includeAll ? '' : "&$filter=ERStatus eq 'open'";
    return this.getRows('events', '?$select=Id,Title,ERDescription,ERLocation,EROrganizerId,EROrganizer/Id,EROrganizer/Title,'
      + 'EROrganizer/Email,ERStatus,ERRegistrationMode,ERRecurrenceType,ERRecurrenceInterval,ERRecurrenceDays,'
      + 'ERRecurrenceUntil,ERRegistrationOpen,ERRegistrationClose,ERMaxSeatsPerRegistration&$expand=EROrganizer'
      + filter + '&$orderby=Title').then((rows: any[]): IEventItem[] => {
        return rows.map((row: any): IEventItem => {
          if (row.EROrganizer) {
            row.EROrganizer = this.normalizePerson(row.EROrganizer);
          }
          return row as IEventItem;
        });
      });
  }

  public getSessions(eventId?: number, includeClosed?: boolean): Promise<ISessionItem[]> {
    var filters: string[] = [];
    if (eventId && eventId > 0) {
      filters.push('EREventId eq ' + String(eventId));
    }
    if (!includeClosed) {
      filters.push("ERStatus eq 'open'");
      filters.push("EREnd ge datetime'" + new Date().toISOString() + "'");
    }
    var filter = filters.length > 0 ? '&$filter=' + filters.join(' and ') : '';
    return this.getRows('sessions', '?$select=Id,Title,EREventId,ERStart,EREnd,ERCapacity,ERWaitlistCapacity,ERStatus,'
      + 'EROccurrenceKey' + filter + '&$orderby=ERStart').then((rows: any[]): ISessionItem[] => rows as ISessionItem[]);
  }

  public getAvailability(sessions: ISessionItem[]): Promise<IAvailability[]> {
    var work = sessions.map((session: ISessionItem): Promise<IAvailability> => {
      return Promise.all([
        this.getRows('slots', "?$select=Id&$filter=ERSessionId eq " + String(session.Id) + " and ERStatus eq 'claimed'"),
        this.getRows('attendees', "?$select=Id&$filter=ERSessionId eq " + String(session.Id) + " and ERStatus eq 'waitlisted'")
      ]).then((parts: any[][]): IAvailability => {
        return {
          sessionId: session.Id,
          capacity: Number(session.ERCapacity || 0),
          claimed: parts[0].length,
          available: Math.max(0, Number(session.ERCapacity || 0) - parts[0].length),
          waitlistCapacity: Number(session.ERWaitlistCapacity || 0),
          waitlisted: parts[1].length
        };
      });
    });
    return Promise.all(work);
  }

  public saveEvent(draft: IEventDraft): Promise<IEventItem> {
    if (!draft.title.trim()) {
      return Promise.reject(new Error('Event title is required.'));
    }
    if (draft.maxSeatsPerRegistration < 1) {
      return Promise.reject(new Error('Maximum seats per registration must be at least 1.'));
    }
    var payload: any = {
      Title: draft.title.trim(),
      ERDescription: draft.description || '',
      ERLocation: draft.location || '',
      EROrganizerId: draft.organizerId,
      ERStatus: draft.status || 'draft',
      ERRegistrationMode: draft.registrationMode || 'perOccurrence',
      ERRecurrenceType: draft.recurrenceType || 'none',
      ERRecurrenceInterval: Math.max(1, Math.floor(Number(draft.recurrenceInterval || 1))),
      ERRecurrenceDays: (draft.recurrenceDays || []).join(','),
      ERRecurrenceUntil: draft.recurrenceUntil ? new Date(draft.recurrenceUntil).toISOString() : undefined,
      ERRegistrationOpen: draft.registrationOpen ? new Date(draft.registrationOpen).toISOString() : undefined,
      ERRegistrationClose: draft.registrationClose ? new Date(draft.registrationClose).toISOString() : undefined,
      ERMaxSeatsPerRegistration: Math.floor(draft.maxSeatsPerRegistration)
    };
    if (draft.id && draft.id > 0) {
      return this.updateItem('events', draft.id, payload).then((): Promise<IEventItem> => this.getItem<IEventItem>('events', draft.id));
    }
    return this.createItem<IEventItem>('events', payload);
  }

  public generateSessionsForEvent(event: IEventItem, draft: IEventDraft): Promise<{ created: number; existing: number }> {
    var firstStart = new Date(draft.firstStart);
    var firstEnd = new Date(draft.firstEnd);
    var until = draft.recurrenceUntil ? new Date(draft.recurrenceUntil) : firstStart;
    if (event.ERRecurrenceType !== 'none' && !draft.recurrenceUntil) {
      return Promise.reject(new Error('A recurrence-until date is required for repeating events.'));
    }
    var weekdayNumbers = (draft.recurrenceDays || []).map(function(day: string): number {
      return Number(day);
    });
    var occurrences = generateOccurrences(event.Id, firstStart, firstEnd, {
      type: event.ERRecurrenceType,
      interval: event.ERRecurrenceInterval || 1,
      weekdays: weekdayNumbers,
      until: until
    });
    return this.getSessions(event.Id, true).then((existing: ISessionItem[]): Promise<{ created: number; existing: number }> => {
      var keys: { [key: string]: boolean } = {};
      existing.forEach(function(session: ISessionItem): void { keys[session.EROccurrenceKey] = true; });
      var createRows = occurrences.filter(function(occurrence: any): boolean { return !keys[occurrence.key]; });
      var created = 0;
      return createRows.reduce((chain: Promise<void>, occurrence: any): Promise<void> => {
        return chain.then((): Promise<void> => {
          return this.createItem<ISessionItem>('sessions', {
            Title: event.Title + ' - ' + occurrence.start.toLocaleString(),
            EREventId: event.Id,
            ERStart: occurrence.start.toISOString(),
            EREnd: occurrence.end.toISOString(),
            ERCapacity: Math.max(0, Math.floor(Number(draft.capacity || 0))),
            ERWaitlistCapacity: Math.max(0, Math.floor(Number(draft.waitlistCapacity || 0))),
            ERStatus: 'open',
            EROccurrenceKey: occurrence.key
          }).then((session: ISessionItem): Promise<void> => {
            created += 1;
            return this.ensureSlots(session, Number(session.ERCapacity || draft.capacity || 0));
          });
        });
      }, Promise.resolve()).then((): { created: number; existing: number } => {
        return { created: created, existing: occurrences.length - created };
      });
    });
  }

  public updateSessionCapacity(session: ISessionItem, capacity: number, waitlistCapacity: number): Promise<void> {
    return this.getRows('slots', "?$select=Id,ERSlotNumber,ERStatus&$filter=ERSessionId eq " + String(session.Id))
      .then((slots: ISlotItem[]): Promise<void> => {
        var claimed = slots.filter(function(slot: ISlotItem): boolean { return slot.ERStatus === 'claimed'; });
        validateCapacityDecrease(capacity, claimed.length);
        return this.updateItem('sessions', session.Id, {
          ERCapacity: capacity,
          ERWaitlistCapacity: Math.max(0, Math.floor(waitlistCapacity))
        }).then((): Promise<void> => {
          if (slots.length > capacity) {
            return this.trimAvailableSlots(slots, slots.length - capacity);
          }
          return this.ensureSlots(session, capacity);
        });
      });
  }

  public register(request: IRegistrationRequest): Promise<IRegistrationResult> {
    var sessions = request.scope === 'series'
      ? request.sessions.filter(function(session: ISessionItem): boolean {
        return session.ERStatus === 'open' && new Date(session.EREnd).getTime() >= new Date().getTime();
      })
      : (request.session ? [request.session] : []);
    if (sessions.length === 0) {
      return Promise.reject(new Error('No applicable open session was selected.'));
    }
    var now = new Date().getTime();
    if (request.event.ERStatus !== 'open'
      || (request.event.ERRegistrationOpen && new Date(request.event.ERRegistrationOpen).getTime() > now)
      || (request.event.ERRegistrationClose && new Date(request.event.ERRegistrationClose).getTime() < now)) {
      return Promise.reject(new Error('Registration is not currently open for this event.'));
    }
    if (!request.attendees || request.attendees.length === 0) {
      return Promise.reject(new Error('Select at least one named attendee.'));
    }
    if (request.attendees.length > Number(request.event.ERMaxSeatsPerRegistration || 1)) {
      return Promise.reject(new Error('This event allows at most ' + String(request.event.ERMaxSeatsPerRegistration) + ' seats per registration.'));
    }
    return this.assertNoDuplicates(request.event.Id, sessions, request.attendees).then((): Promise<IRegistrationResult> => {
      var confirmation = this.confirmationNumber();
      return this.createItem<IRegistrationItem>('registrations', {
        Title: confirmation,
        EREventId: request.event.Id,
        ERSessionId: request.scope === 'perOccurrence' && request.session ? request.session.Id : undefined,
        ERRegistrantId: request.registrant.id,
        ERScope: request.scope,
        ERSeatCount: request.attendees.length,
        ERStatus: 'pending',
        ERConfirmationNumber: confirmation,
        ERCreatedAt: new Date().toISOString()
      }).then((registration: IRegistrationItem): Promise<IRegistrationResult> => {
        return this.createPendingAttendees(registration, request.event.Id, sessions, request.attendees)
          .then((attendees: IAttendeeItem[]): Promise<IRegistrationResult> => {
            return this.claimAll(attendees, request.registrant.id).then(
              (claims: IClaimRecord[]): Promise<IRegistrationResult> => {
                return this.finalizeConfirmed(registration, attendees, claims, request.registrant.email);
              },
              (claimError: any): Promise<IRegistrationResult> => {
                return this.rollbackClaimsForRegistration(registration.Id).then((rollbackErrors: string[]): Promise<IRegistrationResult> => {
                  if (rollbackErrors.length > 0) {
                    throw new Error('Seat claim failed and rollback was incomplete: ' + rollbackErrors.join(' | ')
                      + '. Original error: ' + this.rest.errorMessage(claimError));
                  }
                  return this.canWaitlist(attendees).then((allowed: boolean): Promise<IRegistrationResult> => {
                    if (!allowed) {
                      return this.deleteRegistrationArtifacts(registration.Id).then((): Promise<IRegistrationResult> => {
                        throw new Error('Capacity is no longer available and the waitlist is full. No registration was created.');
                      });
                    }
                    return this.setAttendeeStatuses(attendees, 'waitlisted').then((): Promise<IRegistrationResult> => {
                      return this.updateItem('registrations', registration.Id, { ERStatus: 'waitlisted' }).then((): Promise<IRegistrationResult> => {
                        registration.ERStatus = 'waitlisted';
                        attendees.forEach(function(attendee: IAttendeeItem): void { attendee.ERStatus = 'waitlisted'; });
                        return this.createNotification('registrationWaitlisted', request.event.Id,
                          request.session ? request.session.Id : undefined, registration, request.registrant.email)
                          .then((): IRegistrationResult => ({ registration: registration, attendees: attendees, status: 'waitlisted' }));
                      });
                    });
                  });
                });
              });
          }, (error: any): Promise<IRegistrationResult> => {
            return this.deleteRegistrationArtifacts(registration.Id).then((): Promise<IRegistrationResult> => {
              throw error;
            });
          });
      });
    });
  }

  public getRegistrations(userId?: number): Promise<IRegistrationItem[]> {
    var filter = userId && userId > 0 ? '&$filter=ERRegistrantId eq ' + String(userId) : '';
    return this.getRows('registrations', '?$select=Id,Title,EREventId,ERSessionId,ERRegistrantId,ERRegistrant/Id,'
      + 'ERRegistrant/Title,ERRegistrant/Email,ERScope,ERSeatCount,ERStatus,ERConfirmationNumber,ERCreatedAt'
      + '&$expand=ERRegistrant' + filter + '&$orderby=ERCreatedAt desc')
      .then((rows: any[]): IRegistrationItem[] => {
        return rows.map((row: any): IRegistrationItem => {
          if (row.ERRegistrant) {
            row.ERRegistrant = this.normalizePerson(row.ERRegistrant);
          }
          return row as IRegistrationItem;
        });
      });
  }

  public getAttendees(registrationId?: number): Promise<IAttendeeItem[]> {
    var filter = registrationId && registrationId > 0 ? '&$filter=ERRegistrationId eq ' + String(registrationId) : '';
    return this.getRows('attendees', '?$select=Id,ERRegistrationId,EREventId,ERSessionId,ERAttendeeId,ERAttendee/Id,'
      + 'ERAttendee/Title,ERAttendee/Email,ERAttendeeKey,ERStatus,ERSlotId&$expand=ERAttendee'
      + filter + '&$orderby=Id').then((rows: any[]): IAttendeeItem[] => {
        return rows.map((row: any): IAttendeeItem => {
          if (row.ERAttendee) {
            row.ERAttendee = this.normalizePerson(row.ERAttendee);
          }
          return row as IAttendeeItem;
        });
      });
  }

  public cancelRegistration(registration: IRegistrationItem, requestingUser: IUserInfo, organizer: boolean): Promise<void> {
    if (!organizer && Number(registration.ERRegistrantId) !== Number(requestingUser.id)) {
      return Promise.reject(new Error('Participants can cancel only their own registrations.'));
    }
    if (registration.ERStatus === 'cancelled') {
      return Promise.resolve();
    }
    return this.getAttendees(registration.Id).then((attendees: IAttendeeItem[]): Promise<void> => {
      var sessionIds: { [key: string]: boolean } = {};
      attendees.forEach(function(attendee: IAttendeeItem): void {
        if (attendee.ERSessionId) {
          sessionIds[String(attendee.ERSessionId)] = true;
        }
      });
      return this.releaseAttendeeSlots(attendees).then((errors: string[]): Promise<void> => {
        if (errors.length > 0) {
          throw new Error('Cancellation stopped because one or more slots could not be released safely: ' + errors.join(' | '));
        }
        return this.setAttendeeStatuses(attendees, 'cancelled').then((): Promise<void> => {
          return this.updateItem('registrations', registration.Id, { ERStatus: 'cancelled' }).then((): Promise<void> => {
            return this.createNotification('registrationCancelled', registration.EREventId, registration.ERSessionId,
              registration, registration.ERRegistrant ? registration.ERRegistrant.email : requestingUser.email)
              .then((): Promise<void> => {
                var ids = Object.keys(sessionIds).map(function(value: string): number { return Number(value); });
                return ids.reduce((chain: Promise<void>, sessionId: number): Promise<void> => {
                  return chain.then((): Promise<void> => this.promoteWaitlist(sessionId));
                }, Promise.resolve());
              });
          });
        });
      });
    });
  }

  public promoteWaitlist(sessionId: number): Promise<void> {
    return this.getRows('registrations', "?$select=Id&$filter=ERStatus eq 'waitlisted'&$orderby=ERCreatedAt asc")
      .then((rows: any[]): Promise<void> => {
        var ids = rows.map(function(row: any): number { return Number(row.Id); });
        return this.filterRegistrationsForSession(ids, sessionId, 0, []).then((eligible: number[]): Promise<void> => {
          return this.tryPromoteNext(eligible, 0);
        });
      });
  }

  private filterRegistrationsForSession(ids: number[], sessionId: number, index: number, eligible: number[]): Promise<number[]> {
    if (index >= ids.length) {
      return Promise.resolve(eligible);
    }
    return this.getAttendees(ids[index]).then((attendees: IAttendeeItem[]): Promise<number[]> => {
      if (attendees.some(function(attendee: IAttendeeItem): boolean {
        return attendee.ERStatus === 'waitlisted' && Number(attendee.ERSessionId) === Number(sessionId);
      })) {
        eligible.push(ids[index]);
      }
      return this.filterRegistrationsForSession(ids, sessionId, index + 1, eligible);
    });
  }

  private tryPromoteNext(ids: number[], index: number): Promise<void> {
    if (index >= ids.length) {
      return Promise.resolve();
    }
    return Promise.all([
      this.getItem<IRegistrationItem>('registrations', ids[index]),
      this.getAttendees(ids[index]),
      this.getSessions(undefined, true)
    ]).then((parts: any[]): Promise<void> => {
      var registration: IRegistrationItem = parts[0];
      var sessionEndById: { [key: string]: number } = {};
      (parts[2] as ISessionItem[]).forEach(function(session: ISessionItem): void {
        sessionEndById[String(session.Id)] = new Date(session.EREnd).getTime();
      });
      var currentTime = new Date().getTime();
      var attendees: IAttendeeItem[] = parts[1].filter(function(attendee: IAttendeeItem): boolean {
        return attendee.ERStatus === 'waitlisted' && !!attendee.ERSessionId
          && sessionEndById[String(attendee.ERSessionId)] >= currentTime;
      });
      if (attendees.length === 0) {
        return this.tryPromoteNext(ids, index + 1);
      }
      return this.claimAll(attendees, registration.ERRegistrantId).then(
        (claims: IClaimRecord[]): Promise<void> => {
          return this.finalizeConfirmed(registration, attendees, claims,
            registration.ERRegistrant ? registration.ERRegistrant.email : '', 'waitlistPromoted').then((): void => undefined);
        },
        (): Promise<void> => {
          return this.rollbackClaimsForRegistration(registration.Id).then((rollbackErrors: string[]): Promise<void> => {
            if (rollbackErrors.length > 0) {
              throw new Error('Waitlist promotion rollback failed: ' + rollbackErrors.join(' | '));
            }
            return this.tryPromoteNext(ids, index + 1);
          });
        });
    });
  }

  private provisionLists(index: number, result: ISetupResult): Promise<ISetupResult> {
    if (index >= listSchemas.length) {
      return Promise.resolve(result);
    }
    var schema = listSchemas[index];
    return this.ensureList(schema, result).then((): Promise<ISetupResult> => {
      return this.provisionLists(index + 1, result);
    }).catch((error: any): Promise<ISetupResult> => {
      result.errors.push(listTitle(this.prefix, schema.key) + ': ' + this.rest.errorMessage(error));
      return this.provisionLists(index + 1, result);
    });
  }

  private provisionFields(index: number, result: ISetupResult): Promise<ISetupResult> {
    if (index >= listSchemas.length) {
      return Promise.resolve(result);
    }
    var schema = listSchemas[index];
    return this.ensureFields(schema, 0, result).then((): Promise<ISetupResult> => {
      return this.provisionFields(index + 1, result);
    }).catch((error: any): Promise<ISetupResult> => {
      result.errors.push(listTitle(this.prefix, schema.key) + ' fields: ' + this.rest.errorMessage(error));
      return this.provisionFields(index + 1, result);
    });
  }

  private ensureList(schema: IListSchema, result: ISetupResult): Promise<void> {
    var title = listTitle(this.prefix, schema.key);
    return this.rest.get("web/lists/getByTitle('" + this.rest.escape(title) + "')?$select=Id,ListItemEntityTypeFullName")
      .then((found: IRestResult): void => {
        var list = this.rest.item(found.data);
        this.listIds[schema.key] = String(list.Id);
        this.entityTypes[schema.key] = String(list.ListItemEntityTypeFullName || '');
        result.validatedLists.push(title);
      }).catch((error: any): Promise<void> => {
        if (Number(error.status) !== 404) {
          throw error;
        }
        return this.rest.post('web/lists', {
          __metadata: { type: 'SP.List' },
          BaseTemplate: 100,
          Title: title,
          Description: schema.description,
          ContentTypesEnabled: false
        }).then((created: IRestResult): void => {
          var list = this.rest.item(created.data);
          this.listIds[schema.key] = String(list.Id);
          this.entityTypes[schema.key] = String(list.ListItemEntityTypeFullName || '');
          result.createdLists.push(title);
        });
      });
  }

  private ensureFields(schema: IListSchema, index: number, result: ISetupResult): Promise<void> {
    if (index >= schema.fields.length) {
      return Promise.resolve();
    }
    var field = schema.fields[index];
    var title = listTitle(this.prefix, schema.key);
    return this.rest.get("web/lists/getByTitle('" + this.rest.escape(title) + "')/fields/getByInternalNameOrTitle('"
      + this.rest.escape(field.internalName) + "')?$select=InternalName,TypeAsString,Required")
      .then((found: IRestResult): void => {
        var existing = this.rest.item(found.data);
        if (String(existing.TypeAsString || '').toLowerCase() !== field.type.toLowerCase()) {
          throw new Error('Field ' + field.internalName + ' exists as ' + existing.TypeAsString
            + ', expected ' + field.type + '. Non-destructive setup will not replace it.');
        }
        if (field.required === true && existing.Required !== true) {
          throw new Error('Field ' + field.internalName + ' exists but is not required. Update it manually or use a new prefix.');
        }
      }).catch((error: any): Promise<void> => {
        if (Number(error.status) !== 404) {
          throw error;
        }
        return this.createField(schema, field).then((): void => {
          result.createdFields.push(title + '.' + field.internalName);
        });
      }).then((): Promise<void> => this.ensureFields(schema, index + 1, result));
  }

  private createField(schema: IListSchema, field: IFieldSchema): Promise<void> {
    var xml = field.schemaXml;
    if (field.lookupListKey) {
      var targetId = this.listIds[field.lookupListKey];
      if (!targetId) {
        return Promise.reject(new Error('Lookup target list "' + field.lookupListKey + '" has not been provisioned.'));
      }
      xml = xml.replace('{LIST_ID}', '{' + targetId.replace(/[{}]/g, '') + '}');
    }
    var title = listTitle(this.prefix, schema.key);
    return this.rest.post("web/lists/getByTitle('" + this.rest.escape(title) + "')/fields/CreateFieldAsXml", {
      parameters: {
        __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
        SchemaXml: xml,
        Options: 0
      }
    }).then((): void => undefined);
  }

  private ensureSlots(session: ISessionItem, desiredCapacity: number): Promise<void> {
    return this.getRows('slots', '?$select=Id,ERSlotNumber,ERStatus&$filter=ERSessionId eq ' + String(session.Id)
      + '&$orderby=ERSlotNumber').then((slots: ISlotItem[]): Promise<void> => {
      var highest = 0;
      slots.forEach(function(slot: ISlotItem): void { highest = Math.max(highest, Number(slot.ERSlotNumber || 0)); });
      var additions: number[] = [];
      for (var number = highest + 1; number <= desiredCapacity; number += 1) {
        additions.push(number);
      }
      return additions.reduce((chain: Promise<void>, slotNumber: number): Promise<void> => {
        return chain.then((): Promise<void> => this.createItem<ISlotItem>('slots', {
          Title: session.Title + ' / Seat ' + String(slotNumber),
          EREventId: session.EREventId,
          ERSessionId: session.Id,
          ERSlotNumber: slotNumber,
          ERStatus: 'available'
        }).then((): void => undefined));
      }, Promise.resolve());
    });
  }

  private trimAvailableSlots(slots: ISlotItem[], removeCount: number): Promise<void> {
    var available = slots.filter(function(slot: ISlotItem): boolean {
      return slot.ERStatus === 'available';
    }).sort(function(left: ISlotItem, right: ISlotItem): number {
      return Number(right.ERSlotNumber || 0) - Number(left.ERSlotNumber || 0);
    }).slice(0, removeCount);
    if (available.length < removeCount) {
      return Promise.reject(new Error('Capacity cannot be reduced because too few slots are safely available to remove.'));
    }
    return available.reduce((chain: Promise<void>, slot: ISlotItem): Promise<void> => {
      return chain.then((): Promise<void> => {
        return this.getRawItem('slots', slot.Id).then((fresh: ISlotItem): Promise<void> => {
          if (fresh.ERStatus !== 'available') {
            throw new Error('Slot ' + String(fresh.Id) + ' was claimed while capacity was being changed. Retry the update.');
          }
          return this.deleteItem('slots', fresh.Id, fresh.etag);
        });
      });
    }, Promise.resolve());
  }

  private assertNoDuplicates(eventId: number, sessions: ISessionItem[], attendees: IUserInfo[]): Promise<void> {
    var keys = attendees.map((user: IUserInfo): string => this.attendeeKey(user));
    var duplicates: string[] = [];
    return sessions.reduce((chain: Promise<void>, session: ISessionItem): Promise<void> => {
      return chain.then((): Promise<void> => {
        return this.getRows('attendees', "?$select=ERAttendeeKey,ERStatus&$filter=EREventId eq " + String(eventId)
          + " and ERSessionId eq " + String(session.Id) + " and ERStatus ne 'cancelled'").then((rows: any[]): void => {
          rows.forEach(function(row: any): void {
            var key = String(row.ERAttendeeKey || '').toLowerCase();
            if (keys.indexOf(key) >= 0 && duplicates.indexOf(key) < 0) {
              duplicates.push(key);
            }
          });
        });
      });
    }, Promise.resolve()).then((): void => {
      if (duplicates.length > 0) {
        throw new Error('An active registration already exists for attendee(s): ' + duplicates.join(', ') + '.');
      }
    });
  }

  private createPendingAttendees(
    registration: IRegistrationItem,
    eventId: number,
    sessions: ISessionItem[],
    users: IUserInfo[]
  ): Promise<IAttendeeItem[]> {
    var rows: IAttendeeItem[] = [];
    return sessions.reduce((sessionChain: Promise<void>, session: ISessionItem): Promise<void> => {
      return sessionChain.then((): Promise<void> => {
        return users.reduce((userChain: Promise<void>, user: IUserInfo): Promise<void> => {
          return userChain.then((): Promise<void> => {
            return this.createItem<IAttendeeItem>('attendees', {
              Title: user.title,
              ERRegistrationId: registration.Id,
              EREventId: eventId,
              ERSessionId: session.Id,
              ERAttendeeId: user.id,
              ERAttendeeKey: this.attendeeKey(user),
              ERStatus: 'pending'
            }).then((created: IAttendeeItem): void => { rows.push(created); });
          });
        }, Promise.resolve());
      });
    }, Promise.resolve()).then((): IAttendeeItem[] => rows);
  }

  private claimAll(attendees: IAttendeeItem[], registrantId: number): Promise<IClaimRecord[]> {
    var claims: IClaimRecord[] = [];
    return attendees.reduce((chain: Promise<void>, attendee: IAttendeeItem): Promise<void> => {
      return chain.then((): Promise<void> => this.claimOne(attendee, registrantId).then((claim: IClaimRecord): void => {
        claims.push(claim);
      }));
    }, Promise.resolve()).then((): IClaimRecord[] => claims);
  }

  private claimOne(attendee: IAttendeeItem, registrantId: number): Promise<IClaimRecord> {
    return this.getRows('slots', "?$select=Id,EREventId,ERSessionId,ERSlotNumber,ERStatus&$filter=ERSessionId eq "
      + String(attendee.ERSessionId) + " and ERStatus eq 'available'&$orderby=ERSlotNumber&$top=30")
      .then((candidates: ISlotItem[]): Promise<IClaimRecord> => this.tryClaimCandidate(candidates, 0, attendee, registrantId));
  }

  private tryClaimCandidate(candidates: ISlotItem[], index: number, attendee: IAttendeeItem, registrantId: number): Promise<IClaimRecord> {
    if (index >= candidates.length) {
      return Promise.reject(new Error('No unclaimed capacity slot remains for session ' + String(attendee.ERSessionId) + '.'));
    }
    var candidate = candidates[index];
    return this.getRawItem('slots', candidate.Id).then((fresh: ISlotItem): Promise<IClaimRecord> => {
      if (fresh.ERStatus !== 'available' || !fresh.etag) {
        return this.tryClaimCandidate(candidates, index + 1, attendee, registrantId);
      }
      return this.updateItem('slots', fresh.Id, {
        ERStatus: 'claimed',
        ERClaimedById: attendee.ERAttendeeId || registrantId,
        ERRegistrationId: attendee.ERRegistrationId,
        ERAttendeeId: attendee.Id,
        ERAttendeeKey: attendee.ERAttendeeKey
      }, fresh.etag).then((): Promise<IClaimRecord> => {
        return this.updateItem('attendees', attendee.Id, { ERSlotId: fresh.Id }).then((): IClaimRecord => {
          attendee.ERSlotId = fresh.Id;
          return { slot: fresh, attendee: attendee };
        });
      }).catch((error: any): Promise<IClaimRecord> => {
        if (Number(error.status) === 412) {
          return this.tryClaimCandidate(candidates, index + 1, attendee, registrantId);
        }
        throw error;
      });
    });
  }

  private finalizeConfirmed(
    registration: IRegistrationItem,
    attendees: IAttendeeItem[],
    claims: IClaimRecord[],
    email: string,
    notificationType?: string
  ): Promise<IRegistrationResult> {
    return this.setAttendeeStatuses(attendees, 'confirmed').then((): Promise<IRegistrationResult> => {
      return this.updateItem('registrations', registration.Id, { ERStatus: 'confirmed' }).then((): Promise<IRegistrationResult> => {
        registration.ERStatus = 'confirmed';
        attendees.forEach(function(attendee: IAttendeeItem): void { attendee.ERStatus = 'confirmed'; });
        return this.createNotification(notificationType || 'registrationConfirmed', registration.EREventId,
          registration.ERSessionId, registration, email).then((): IRegistrationResult => {
            return { registration: registration, attendees: attendees, status: 'confirmed' };
          });
      });
    });
  }

  private canWaitlist(attendees: IAttendeeItem[]): Promise<boolean> {
    var bySession: { [key: string]: number } = {};
    attendees.forEach(function(attendee: IAttendeeItem): void {
      var key = String(attendee.ERSessionId);
      bySession[key] = (bySession[key] || 0) + 1;
    });
    var sessionIds = Object.keys(bySession);
    return Promise.all(sessionIds.map((id: string): Promise<boolean> => {
      return Promise.all([
        this.getItem<ISessionItem>('sessions', Number(id)),
        this.getRows('attendees', "?$select=Id&$filter=ERSessionId eq " + id + " and ERStatus eq 'waitlisted'")
      ]).then((parts: any[]): boolean => {
        return Number(parts[0].ERWaitlistCapacity || 0) - parts[1].length >= bySession[id];
      });
    })).then((values: boolean[]): boolean => values.every(function(value: boolean): boolean { return value; }));
  }

  private rollbackClaimsForRegistration(registrationId: number): Promise<string[]> {
    return this.getRows('slots', "?$select=Id&$filter=ERRegistrationId eq " + String(registrationId)
      + " and ERStatus eq 'claimed'").then((slots: ISlotItem[]): Promise<string[]> => {
      var errors: string[] = [];
      return slots.reduce((chain: Promise<void>, slot: ISlotItem): Promise<void> => {
        return chain.then((): Promise<void> => {
          return this.getRawItem('slots', slot.Id).then((fresh: ISlotItem): Promise<void> => {
            return this.updateItem('slots', slot.Id, {
              ERStatus: 'available',
              ERClaimedById: null,
              ERRegistrationId: null,
              ERAttendeeId: null,
              ERAttendeeKey: ''
            }, fresh.etag);
          }).catch((error: any): void => {
            errors.push('slot ' + String(slot.Id) + ': ' + this.rest.errorMessage(error));
          });
        });
      }, Promise.resolve()).then((): string[] => errors);
    });
  }

  private releaseAttendeeSlots(attendees: IAttendeeItem[]): Promise<string[]> {
    var slotIds = attendees.filter(function(attendee: IAttendeeItem): boolean {
      return !!attendee.ERSlotId;
    }).map(function(attendee: IAttendeeItem): number { return Number(attendee.ERSlotId); });
    var errors: string[] = [];
    return slotIds.reduce((chain: Promise<void>, slotId: number): Promise<void> => {
      return chain.then((): Promise<void> => {
        return this.getRawItem('slots', slotId).then((slot: ISlotItem): Promise<void> => {
          return this.updateItem('slots', slotId, {
            ERStatus: 'available',
            ERClaimedById: null,
            ERRegistrationId: null,
            ERAttendeeId: null,
            ERAttendeeKey: ''
          }, slot.etag).then((): Promise<void> => this.updateItem('attendees',
            attendees.filter(function(attendee: IAttendeeItem): boolean {
              return Number(attendee.ERSlotId) === Number(slotId);
            })[0].Id, { ERSlotId: null }));
        }).catch((error: any): void => {
          errors.push('slot ' + String(slotId) + ': ' + this.rest.errorMessage(error));
        });
      });
    }, Promise.resolve()).then((): string[] => errors);
  }

  private setAttendeeStatuses(attendees: IAttendeeItem[], status: string): Promise<void> {
    return attendees.reduce((chain: Promise<void>, attendee: IAttendeeItem): Promise<void> => {
      return chain.then((): Promise<void> => this.updateItem('attendees', attendee.Id, { ERStatus: status }));
    }, Promise.resolve());
  }

  private createNotification(
    type: string,
    eventId: number,
    sessionId: number,
    registration: IRegistrationItem,
    recipientEmail: string
  ): Promise<void> {
    var payload = {
      registrationId: registration.Id,
      confirmationNumber: registration.ERConfirmationNumber,
      eventId: eventId,
      sessionId: sessionId,
      status: registration.ERStatus,
      occurredAt: new Date().toISOString()
    };
    return this.createItem<any>('notifications', {
      Title: type + ' - ' + registration.ERConfirmationNumber,
      EREventId: eventId || undefined,
      ERSessionId: sessionId || undefined,
      ERRegistrationId: registration.Id,
      ERNotificationType: type,
      ERRecipientEmails: recipientEmail || '',
      ERPayloadJson: JSON.stringify(payload),
      ERProcessed: false,
      ERErrorText: ''
    }).then((): void => undefined);
  }

  private deleteRegistrationArtifacts(registrationId: number): Promise<void> {
    return this.getAttendees(registrationId).then((attendees: IAttendeeItem[]): Promise<void> => {
      return attendees.reduce((chain: Promise<void>, attendee: IAttendeeItem): Promise<void> => {
        return chain.then((): Promise<void> => this.deleteItem('attendees', attendee.Id));
      }, Promise.resolve()).then((): Promise<void> => this.deleteItem('registrations', registrationId));
    });
  }

  private attendeeKey(user: IUserInfo): string {
    return String(user.email || user.loginName || user.id).trim().toLowerCase();
  }

  private normalizePerson(person: any): IUserInfo {
    return {
      id: Number(person.Id || person.id || 0),
      title: String(person.Title || person.title || ''),
      email: String(person.Email || person.email || ''),
      loginName: String(person.LoginName || person.loginName || '')
    };
  }

  private confirmationNumber(): string {
    var now = new Date();
    var random = Math.floor(Math.random() * 9000) + 1000;
    return 'ER-' + String(now.getUTCFullYear())
      + this.pad(now.getUTCMonth() + 1) + this.pad(now.getUTCDate())
      + '-' + String(now.getTime()).slice(-6) + '-' + String(random);
  }

  private pad(value: number): string {
    return value < 10 ? '0' + String(value) : String(value);
  }

  private getRows(key: string, query: string): Promise<any[]> {
    return this.rest.get("web/lists/getByTitle('" + this.rest.escape(listTitle(this.prefix, key)) + "')/items" + query)
      .then((result: IRestResult): any[] => this.rest.rows(result.data));
  }

  private getItem<T>(key: string, id: number): Promise<T> {
    return this.rest.get("web/lists/getByTitle('" + this.rest.escape(listTitle(this.prefix, key)) + "')/items(" + String(id) + ")")
      .then((result: IRestResult): T => this.rest.item(result.data) as T);
  }

  private getRawItem(key: string, id: number): Promise<ISlotItem> {
    return this.rest.get("web/lists/getByTitle('" + this.rest.escape(listTitle(this.prefix, key)) + "')/items(" + String(id)
      + ")?$select=Id,EREventId,ERSessionId,ERSlotNumber,ERStatus,ERRegistrationId,ERAttendeeId,ERAttendeeKey")
      .then((result: IRestResult): ISlotItem => {
        var item = this.rest.item(result.data) as ISlotItem;
        item.etag = result.response.headers.get('ETag') || result.response.headers.get('etag') || '';
        if (!item.etag) {
          throw new Error('SharePoint did not return an ETag for slot ' + String(id) + '; the slot was not modified.');
        }
        return item;
      });
  }

  private createItem<T>(key: string, payload: any): Promise<T> {
    return this.getEntityType(key).then((entityType: string): Promise<T> => {
      var body: any = this.cleanPayload(payload);
      body.__metadata = { type: entityType };
      return this.rest.post("web/lists/getByTitle('" + this.rest.escape(listTitle(this.prefix, key)) + "')/items", body)
        .then((result: IRestResult): T => this.rest.item(result.data) as T);
    });
  }

  private updateItem(key: string, id: number, payload: any, etag?: string): Promise<void> {
    return this.getEntityType(key).then((entityType: string): Promise<void> => {
      var body: any = this.cleanPayload(payload);
      body.__metadata = { type: entityType };
      return this.rest.post("web/lists/getByTitle('" + this.rest.escape(listTitle(this.prefix, key)) + "')/items(" + String(id) + ")",
        body, { 'IF-MATCH': etag || '*', 'X-HTTP-Method': 'MERGE' }).then((): void => undefined);
    });
  }

  private deleteItem(key: string, id: number, etag?: string): Promise<void> {
    return this.rest.post("web/lists/getByTitle('" + this.rest.escape(listTitle(this.prefix, key)) + "')/items(" + String(id) + ")",
      undefined, { 'IF-MATCH': etag || '*', 'X-HTTP-Method': 'DELETE' }).then((): void => undefined);
  }

  private getEntityType(key: string): Promise<string> {
    if (this.entityTypes[key]) {
      return Promise.resolve(this.entityTypes[key]);
    }
    return this.rest.get("web/lists/getByTitle('" + this.rest.escape(listTitle(this.prefix, key))
      + "')?$select=Id,ListItemEntityTypeFullName").then((result: IRestResult): string => {
        var list = this.rest.item(result.data);
        this.listIds[key] = String(list.Id);
        this.entityTypes[key] = String(list.ListItemEntityTypeFullName);
        return this.entityTypes[key];
      });
  }

  private cleanPayload(payload: any): any {
    var result: any = {};
    Object.keys(payload || {}).forEach(function(key: string): void {
      if (payload[key] !== undefined) {
        result[key] = payload[key];
      }
    });
    return result;
  }
}
