export interface IUserInfo {
  id: number;
  title: string;
  email: string;
  loginName?: string;
}

export interface IEventItem {
  Id: number;
  Title: string;
  ERDescription?: string;
  ERLocation?: string;
  EROrganizerId?: number;
  EROrganizer?: IUserInfo;
  ERStatus: string;
  ERRegistrationMode: string;
  ERRecurrenceType: string;
  ERRecurrenceInterval: number;
  ERRecurrenceDays?: string;
  ERRecurrenceUntil?: string;
  ERRegistrationOpen?: string;
  ERRegistrationClose?: string;
  ERMaxSeatsPerRegistration: number;
}

export interface ISessionItem {
  Id: number;
  Title: string;
  EREventId: number;
  ERStart: string;
  EREnd: string;
  ERCapacity: number;
  ERWaitlistCapacity: number;
  ERStatus: string;
  EROccurrenceKey: string;
}

export interface ISlotItem {
  Id: number;
  EREventId: number;
  ERSessionId: number;
  ERSlotNumber: number;
  ERStatus: string;
  ERClaimedById?: number;
  ERRegistrationId?: number;
  ERAttendeeId?: number;
  ERAttendeeKey?: string;
  etag?: string;
}

export interface IRegistrationItem {
  Id: number;
  EREventId: number;
  ERSessionId?: number;
  ERRegistrantId: number;
  ERRegistrant?: IUserInfo;
  ERScope: string;
  ERSeatCount: number;
  ERStatus: string;
  ERConfirmationNumber: string;
  ERCreatedAt: string;
}

export interface IAttendeeItem {
  Id: number;
  ERRegistrationId: number;
  EREventId: number;
  ERSessionId?: number;
  ERAttendeeId: number;
  ERAttendee?: IUserInfo;
  ERAttendeeKey: string;
  ERStatus: string;
  ERSlotId?: number;
}

export interface IRegistrationRequest {
  event: IEventItem;
  session?: ISessionItem;
  sessions: ISessionItem[];
  registrant: IUserInfo;
  attendees: IUserInfo[];
  scope: string;
}

export interface IAvailability {
  sessionId: number;
  capacity: number;
  claimed: number;
  available: number;
  waitlistCapacity: number;
  waitlisted: number;
}

export interface ISetupResult {
  createdLists: string[];
  createdFields: string[];
  validatedLists: string[];
  errors: string[];
}

export interface IEventDraft {
  id?: number;
  title: string;
  description: string;
  location: string;
  organizerId: number;
  status: string;
  registrationMode: string;
  recurrenceType: string;
  recurrenceInterval: number;
  recurrenceDays: string[];
  recurrenceUntil: string;
  registrationOpen: string;
  registrationClose: string;
  maxSeatsPerRegistration: number;
  firstStart: string;
  firstEnd: string;
  capacity: number;
  waitlistCapacity: number;
}

export interface IRegistrationResult {
  registration: IRegistrationItem;
  attendees: IAttendeeItem[];
  status: string;
}

export interface IRegistrationSelection {
  id: number;
  status: string;
  confirmation: string;
}
