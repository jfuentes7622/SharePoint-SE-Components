export interface IFieldSchema {
  internalName: string;
  displayName: string;
  type: string;
  required?: boolean;
  schemaXml: string;
  lookupListKey?: string;
}

export interface IListSchema {
  key: string;
  suffix: string;
  description: string;
  fields: IFieldSchema[];
}

function text(name: string, displayName: string, required?: boolean, maxLength?: number): IFieldSchema {
  return {
    internalName: name,
    displayName: displayName,
    type: 'Text',
    required: required,
    schemaXml: '<Field Type="Text" Name="' + name + '" StaticName="' + name + '" DisplayName="' + displayName
      + '" Required="' + (required ? 'TRUE' : 'FALSE') + '" MaxLength="' + String(maxLength || 255) + '" />'
  };
}

function note(name: string, displayName: string): IFieldSchema {
  return {
    internalName: name,
    displayName: displayName,
    type: 'Note',
    schemaXml: '<Field Type="Note" Name="' + name + '" StaticName="' + name + '" DisplayName="' + displayName
      + '" NumLines="6" RichText="FALSE" AppendOnly="FALSE" />'
  };
}

function numberField(name: string, displayName: string, required?: boolean): IFieldSchema {
  return {
    internalName: name,
    displayName: displayName,
    type: 'Number',
    required: required,
    schemaXml: '<Field Type="Number" Name="' + name + '" StaticName="' + name + '" DisplayName="' + displayName
      + '" Required="' + (required ? 'TRUE' : 'FALSE') + '" Min="0" Decimals="0" />'
  };
}

function dateField(name: string, displayName: string, required?: boolean): IFieldSchema {
  return {
    internalName: name,
    displayName: displayName,
    type: 'DateTime',
    required: required,
    schemaXml: '<Field Type="DateTime" Name="' + name + '" StaticName="' + name + '" DisplayName="' + displayName
      + '" Required="' + (required ? 'TRUE' : 'FALSE') + '" Format="DateTime" FriendlyDisplayFormat="Disabled" />'
  };
}

function booleanField(name: string, displayName: string): IFieldSchema {
  return {
    internalName: name,
    displayName: displayName,
    type: 'Boolean',
    schemaXml: '<Field Type="Boolean" Name="' + name + '" StaticName="' + name + '" DisplayName="' + displayName
      + '"><Default>0</Default></Field>'
  };
}

function person(name: string, displayName: string, required?: boolean): IFieldSchema {
  return {
    internalName: name,
    displayName: displayName,
    type: 'User',
    required: required,
    schemaXml: '<Field Type="User" Name="' + name + '" StaticName="' + name + '" DisplayName="' + displayName
      + '" Required="' + (required ? 'TRUE' : 'FALSE')
      + '" List="UserInfo" ShowField="ImnName" UserSelectionMode="PeopleOnly" UserSelectionScope="0" />'
  };
}

function lookup(name: string, displayName: string, listKey: string, required?: boolean): IFieldSchema {
  return {
    internalName: name,
    displayName: displayName,
    type: 'Lookup',
    required: required,
    lookupListKey: listKey,
    schemaXml: '<Field Type="Lookup" Name="' + name + '" StaticName="' + name + '" DisplayName="' + displayName
      + '" Required="' + (required ? 'TRUE' : 'FALSE') + '" List="{LIST_ID}" ShowField="Title" RelationshipDeleteBehavior="Restrict" />'
  };
}

function choice(name: string, displayName: string, choices: string[], defaultValue: string, required?: boolean): IFieldSchema {
  var values = choices.map(function(value: string): string { return '<CHOICE>' + value + '</CHOICE>'; }).join('');
  return {
    internalName: name,
    displayName: displayName,
    type: 'Choice',
    required: required,
    schemaXml: '<Field Type="Choice" Name="' + name + '" StaticName="' + name + '" DisplayName="' + displayName
      + '" Required="' + (required ? 'TRUE' : 'FALSE') + '" Format="Dropdown"><CHOICES>' + values
      + '</CHOICES><Default>' + defaultValue + '</Default></Field>'
  };
}

export var listSchemas: IListSchema[] = [
  {
    key: 'events',
    suffix: 'Events',
    description: 'Event definitions and recurrence settings.',
    fields: [
      note('ERDescription', 'Description'),
      text('ERLocation', 'Location'),
      person('EROrganizer', 'Organizer', true),
      choice('ERStatus', 'Status', ['draft', 'open', 'closed', 'cancelled'], 'draft', true),
      choice('ERRegistrationMode', 'Registration Mode', ['perOccurrence', 'series'], 'perOccurrence', true),
      choice('ERRecurrenceType', 'Recurrence Type', ['none', 'daily', 'weekly'], 'none', true),
      numberField('ERRecurrenceInterval', 'Recurrence Interval', true),
      text('ERRecurrenceDays', 'Recurrence Days'),
      dateField('ERRecurrenceUntil', 'Recurrence Until'),
      dateField('ERRegistrationOpen', 'Registration Open'),
      dateField('ERRegistrationClose', 'Registration Close'),
      numberField('ERMaxSeatsPerRegistration', 'Max Seats Per Registration', true)
    ]
  },
  {
    key: 'sessions',
    suffix: 'Sessions',
    description: 'Generated event occurrences.',
    fields: [
      lookup('EREvent', 'Event', 'events', true),
      dateField('ERStart', 'Start', true),
      dateField('EREnd', 'End', true),
      numberField('ERCapacity', 'Capacity', true),
      numberField('ERWaitlistCapacity', 'Waitlist Capacity', true),
      choice('ERStatus', 'Status', ['open', 'closed', 'cancelled'], 'open', true),
      text('EROccurrenceKey', 'Occurrence Key', true)
    ]
  },
  {
    key: 'slots',
    suffix: 'Registration Slots',
    description: 'Pre-created numbered capacity slots claimed with ETags.',
    fields: [
      lookup('EREvent', 'Event', 'events', true),
      lookup('ERSession', 'Session', 'sessions', true),
      numberField('ERSlotNumber', 'Slot Number', true),
      choice('ERStatus', 'Status', ['available', 'claimed'], 'available', true),
      person('ERClaimedBy', 'Claimed By'),
      lookup('ERRegistration', 'Registration', 'registrations'),
      lookup('ERAttendee', 'Attendee', 'attendees'),
      text('ERAttendeeKey', 'Attendee Key')
    ]
  },
  {
    key: 'registrations',
    suffix: 'Registrations',
    description: 'Registration headers for occurrence and series requests.',
    fields: [
      lookup('EREvent', 'Event', 'events', true),
      lookup('ERSession', 'Session', 'sessions'),
      person('ERRegistrant', 'Registrant', true),
      choice('ERScope', 'Scope', ['perOccurrence', 'series'], 'perOccurrence', true),
      numberField('ERSeatCount', 'Seat Count', true),
      choice('ERStatus', 'Status', ['pending', 'confirmed', 'waitlisted', 'cancelled'], 'pending', true),
      text('ERConfirmationNumber', 'Confirmation Number', true),
      dateField('ERCreatedAt', 'Created Timestamp', true)
    ]
  },
  {
    key: 'attendees',
    suffix: 'Registration Attendees',
    description: 'Named attendees and their session/slot assignment.',
    fields: [
      lookup('ERRegistration', 'Registration', 'registrations', true),
      lookup('EREvent', 'Event', 'events', true),
      lookup('ERSession', 'Session', 'sessions'),
      person('ERAttendee', 'Attendee', true),
      text('ERAttendeeKey', 'Attendee Key', true),
      choice('ERStatus', 'Status', ['pending', 'confirmed', 'waitlisted', 'cancelled'], 'pending', true),
      lookup('ERSlot', 'Slot', 'slots')
    ]
  },
  {
    key: 'notifications',
    suffix: 'Notification Events',
    description: 'Power Automate contract; email is never sent by the web part.',
    fields: [
      lookup('EREvent', 'Event', 'events'),
      lookup('ERSession', 'Session', 'sessions'),
      lookup('ERRegistration', 'Registration', 'registrations'),
      choice('ERNotificationType', 'Notification Type',
        ['registrationConfirmed', 'registrationWaitlisted', 'registrationCancelled', 'waitlistPromoted'], 'registrationConfirmed', true),
      note('ERRecipientEmails', 'Recipient Emails'),
      note('ERPayloadJson', 'Payload JSON'),
      booleanField('ERProcessed', 'Processed'),
      note('ERErrorText', 'Error Text')
    ]
  }
];

export function listTitle(prefix: string, key: string): string {
  var normalized = String(prefix || 'SPS Event Registration').trim() || 'SPS Event Registration';
  for (var i = 0; i < listSchemas.length; i += 1) {
    if (listSchemas[i].key === key) {
      return normalized + ' ' + listSchemas[i].suffix;
    }
  }
  throw new Error('Unknown Event Registration list key: ' + key);
}
