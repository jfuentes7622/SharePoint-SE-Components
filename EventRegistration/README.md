# Event Registration

**Event Registration** is a SharePoint Framework 1.5.1 web part for SharePoint
Subscription Edition. It provides participant registration, named multi-seat
attendees, organizer administration, deterministic recurrence, ETag-protected
capacity, waitlists, CSV export, and a Power Automate notification contract.

## Versions and package

| Artifact | Version |
|---|---|
| npm package | `1.0.0` |
| SPFx solution | `1.0.0.1` |
| SPFx feature | `1.0.0.1` |
| React | `15.6.2` |
| TypeScript | `2.4.2` |
| `@microsoft/sp-*` | `1.5.1` |

The gulp pre-build task derives the solution and feature version from the npm
version and keeps both SPFx versions identical. The deployable package is
`sharepoint/solution/sps-event-registration.sppkg`.

## Configuration and permissions

The property pane provides:

- **List name prefix** (default `SPS Event Registration`).
- **Organizer SharePoint group**, loaded from site groups.
- **Preferred Calendar instance**, discovered from available Dynamic Data
  sources whose alias is `CalendarWebPart`.
- **Button display** (`text`, `iconText`, or `icon`).
- **Enable diagnostics logging**.
- **Title** and optional multiline **Description**, displayed above the registration views.
- **Appearance mode**: Manual styling, Site theme only, or Site theme with overrides.
- Inline PnP color pickers matching Search Control for background, text, accent,
  button text, active button background, and borders; sliders for border width,
  corner radius, and font size; and a font-family text field.
- **Exclude this web part from SPS Tabs**.

The solution version remains at the top of the property pane. Diagnostics are
enabled unless explicitly disabled; genuine failures remain logged regardless.
Existing instances default to Manual styling, retain their button mode, and
use the original title when no title property has been saved. Site theme only
disables manual color/font-family controls without deleting their saved values.
Error, success, and destructive-action colors retain their semantic meaning.
The initial site theme is cached without rendering until SharePoint has
initialized the web part; subsequent theme changes refresh the rendered view
and property pane.

The web part reads `IsSiteAdmin` for the current user. A user is an organizer
when they are a Site Collection Administrator or a member of the configured
organizer group. Before a group is configured, only a Site Collection
Administrator may run setup. Participants can browse open events, register
named users, view their own registrations, and cancel only registrations they
created. Organizer views load and manage all registrations.

List permissions remain normal SharePoint permissions. For participant
self-service, authenticated participants require read/add/update rights on the
six configured lists. Use SharePoint list permissions appropriate to the site;
the web part's ownership checks are not a substitute for list security.

## Setup

Until all six configured lists exist, the web part shows **Setup required**
and, for authorized organizers, **Provision lists and fields**, without a raw
missing-list REST error. It checks list availability before requesting event
data. Permission, connection, and schema failures still surface as errors.

An authorized organizer opens **Organizer > Setup** and selects **Validate /
provision**. Setup:

1. Creates all six missing custom lists.
2. Creates fields with SharePoint SE-compatible
   `fields/CreateFieldAsXml` calls.
3. Resolves lookup-list GUIDs before creating lookup fields.
4. Validates existing field internal names, types, and required flags.
5. Reports list and field failures. It never reports success when validation
   contains errors.

Setup is non-destructive. It never deletes or replaces an incompatible field.
Correct the reported schema manually or configure a new prefix.

REST reads and writes explicitly use `OData-Version: 3.0` to match SharePoint's
JSON formats and the typed `__metadata` payloads used for list, field, and item
creation. This overrides SPHttpClient v1's OData 4.0 default, which can cause
list creation to fail with a JSON Light parsing error.

## List schema

All custom internal names begin with `ER`. Each list also uses SharePoint's
standard `Title`, `Id`, `Created`, and `Modified` fields.

### `<prefix> Events`

| Internal name | Type | Purpose |
|---|---|---|
| `ERDescription` | Note | Description |
| `ERLocation` | Text | Location |
| `EROrganizer` | Person | Organizer |
| `ERStatus` | Choice | `draft`, `open`, `closed`, `cancelled` |
| `ERRegistrationMode` | Choice | `perOccurrence`, `series` |
| `ERRecurrenceType` | Choice | `none`, `daily`, `weekly` |
| `ERRecurrenceInterval` | Number | Daily/weekly interval |
| `ERRecurrenceDays` | Text | Comma-separated weekday numbers, Sunday `0` |
| `ERRecurrenceUntil` | Date/time | Inclusive recurrence end |
| `ERRegistrationOpen` | Date/time | Registration window start |
| `ERRegistrationClose` | Date/time | Registration window end |
| `ERMaxSeatsPerRegistration` | Number | Named attendees allowed per request |

### `<prefix> Sessions`

`EREvent` (Events lookup), `ERStart`, `EREnd`, `ERCapacity`,
`ERWaitlistCapacity`, `ERStatus` (`open`, `closed`, `cancelled`), and unique
deterministic `EROccurrenceKey`.

### `<prefix> Registration Slots`

`EREvent`, `ERSession`, `ERSlotNumber`, `ERStatus` (`available`, `claimed`),
`ERClaimedBy`, `ERRegistration`, `ERAttendee`, and `ERAttendeeKey`. Numbered
items are created before registration and are the concurrency primitive.

### `<prefix> Registrations`

`EREvent`, optional `ERSession`, `ERRegistrant`, `ERScope`,
`ERSeatCount`, `ERStatus` (`pending`, `confirmed`, `waitlisted`, `cancelled`),
`ERConfirmationNumber`, and `ERCreatedAt`.

### `<prefix> Registration Attendees`

`ERRegistration`, `EREvent`, optional `ERSession`, `ERAttendee`,
`ERAttendeeKey`, `ERStatus`, and optional `ERSlot`. Series registrations have
one attendee row per named attendee per applicable session.

### `<prefix> Notification Events`

`EREvent`, `ERSession`, `ERRegistration`, `ERNotificationType`,
`ERRecipientEmails`, `ERPayloadJson`, `ERProcessed`, and `ERErrorText`.
Notification types are `registrationConfirmed`, `registrationWaitlisted`,
`registrationCancelled`, and `waitlistPromoted`.

## Recurrence semantics

The organizer specifies the first session start/end and chooses no recurrence,
daily, or weekly recurrence. Intervals must be positive. Weekly recurrence
accepts selected weekday numbers (`0` Sunday through `6` Saturday), defaulting
to the first session's weekday when none is supplied. The until date is
inclusive. Generation is capped at 366 occurrences per operation.

Occurrence keys use:

```text
<event-id>-<UTC YYYYMMDD>T<UTC HHMMSS>Z
```

Session generation compares this key with existing sessions and creates only
missing occurrences, so rerunning generation is idempotent.

`perOccurrence` creates one registration for one session. `series` creates one
registration header and attendee/seat assignments for every open future
session in the event.

## Capacity, concurrency, and rollback

Each session pre-creates one numbered slot per confirmed seat. A claim:

1. Reads an available slot.
2. Reads that individual slot again and captures its current ETag.
3. Sends `MERGE` with the **exact** `IF-MATCH` ETag.
4. On HTTP 412, tries another available slot.

Each named attendee claims exactly one slot. Active attendee keys are checked
per applicable session before artifacts are created. A series claim succeeds
only after every named attendee has a slot in every applicable session. If any
claim fails, the service explicitly re-reads and releases every slot already
claimed by the registration. Incomplete rollback is reported as an error and
is never hidden.

Cancellation re-reads each slot, releases it with the exact ETag, clears lookup
assignments, and then cancels attendee and registration records. Capacity
increases create additional slots. A decrease below the number of claimed
slots is rejected; safe decreases remove only currently available slots using
exact ETags.

## Waitlist and promotion

When confirmed capacity is exhausted, the request becomes waitlisted only if
every applicable session has enough waitlist capacity. Waitlisted attendees
have no slot. Cancellation attempts promotion in FIFO registration-created
order:

- Per-occurrence candidates need capacity in the affected future session.
- Series candidates must claim every named attendee seat in every future
  applicable session.
- A conflict releases all seats claimed by that promotion attempt before the
  next FIFO candidate is considered.

Successful promotion updates registration/attendee statuses and creates a
`waitlistPromoted` Notification Event.

## Power Automate contract

The web part never sends email. A flow should trigger when an item is created
in `<prefix> Notification Events`, branch on `ERNotificationType`, parse
`ERPayloadJson`, send the desired message to `ERRecipientEmails`, and set
`ERProcessed` to true. If processing fails, the flow should retain false and
write an actionable message to `ERErrorText`.

## Calendar and URL integration

The preferred Calendar is subscribed with the SPFx 1.5.1 Dynamic Data provider
compatibility API. The web part consumes:

- `selectedSessionId` as a session selection when exposed.
- `selectedItemId` as an event selection.

URL query parameters `sessionid` and `eventid` take precedence over Dynamic
Data. `source` or `returnurl` can request a post-registration redirect. The
web part resolves and validates the URL and redirects only when its origin
matches the current page origin.

The web part publishes `selectedRegistrationId`,
`selectedRegistrationStatus`, and `selectedRegistrationConfirmation` as
Dynamic Data properties for future integrations.

## Organizer user interface

Organizer tabs provide Setup, Events, Sessions, Registrations / Attendees,
Waitlist, and CSV Export. Event editing includes recurrence and session
generation. Session editing enforces capacity rules. The registration views
show named attendees and allow organizer cancellation. CSV contains
registration and attendee rows. Participant and organizer operations expose
loading, success, and actionable error states; controls have labels and
responsive keyboard-accessible HTML.

## Build, test, and deploy

The serve profiles match Search Control: `npx gulp serve` uses the SPO workbench
(`default`), and `npx gulp serve --config=navinLab` uses the SPS lab workbench.
Both use HTTPS port 4321 and the workbench API on port 5432.

From this directory:

```powershell
npm install
npx gulp test
npx gulp bundle --ship
npx gulp package-solution --ship
```

The helper tests run without SharePoint and cover deterministic recurrence,
daily/weekly generation, cross-session capacity decisions, waitlist capacity,
capacity-decrease validation, manual/theme appearance, theme overrides, and
restoring saved manual settings. Mocked REST client tests cover OData request
headers, typed payloads, exact ETag preservation, and error propagation.
End-to-end REST transaction behavior requires a
SharePoint SE integration site and is not simulated by the SPFx 1.5.1 test
runner.

Upload `sharepoint/solution/sps-event-registration.sppkg` to the SharePoint app
catalog, deploy it, add the app to the target site if required, place the web
part on a modern page, configure the organizer group, and run setup.
