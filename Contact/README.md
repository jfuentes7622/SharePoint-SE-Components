# SPS Contacts

SPFx 1.5.1 personnel directory web part for SharePoint Server Subscription Edition.

## Features

- Load contacts from a SharePoint list or maintain inline contact entries in the property pane.
- Filter contacts through directorate and division hierarchy.
- Display job title, name, email, phone, VOIP, biography link, and optional photo.
- Render photos as square, rounded, or circular images.
- Customize card headers, titles, information areas, fonts, colors, alignment, and corners.
- Load an optional override stylesheet.
- Enable diagnostic console logging for list, record, and image services.

## Data Source

Enable **Use a List** to use SharePoint data, then choose **Use an Existing List** or **Create a New List**. After a list is selected or created, the **Columns to Use** group loads its columns for every contact field mapping.

To create a compatible list on demand, enter a name under **New Contact List Name** and select **Create Contact List**. The web part creates the required columns in the current SharePoint web, selects the new list, and configures all field mappings. The editing user must have permission to manage lists.

Disable **Use Existing List** to enter contact records directly in the property-pane collection editor.

## Appearance

| Property | Description |
|---|---|
| `imageShape` | Contact image shape: `square`, `rounded`, or `circle`. |
| `imageWidth` | Contact image size in pixels. |
| `PersonnelPanelHeaderBackColor` / `PersonnelPanelHeaderTextColor` | Card header colors. |
| `headerFontFamily` / `headerFontStyle` / `headerFontBold` | Header typography. |
| `headerAlignment` / `headerTopCorners` | Header alignment and corner style. |
| `tileTitleColor` | Job-title text color. |
| `titleFontFamily` / `titleFontStyle` / `titleFontBold` | Job-title typography. |
| `tileInfoBackgroundColor` | Background behind contact details. |
| `overridecss` | Optional external stylesheet URL. |
| `enableDiagnostics` | Writes contact loading and rendering details to the browser console. |

## Build And Package

```powershell
npm install
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.

## Usage

1. Add **SPS Contacts** to a page and choose an existing SharePoint list or inline contact entries.
2. For a list source, map the organization, identity, contact, photo, and biography columns used by the cards.
3. Select any directorate/division filters and configure image, header, and tile appearance.
4. Verify cards with missing photos or optional contact values before publishing.

## Properties and common configuration

- **Data:** `ContactsListId` and the mapped field properties connect an existing directory. Inline mode uses `customList` for records maintained with the web part.
- **Hierarchy:** directorate, division, branch, and group mappings provide organization and cascading filters when those columns contain consistent values.
- **Cards:** title, name, email, phone, biography, VOIP, display-photo, and image-link fields determine the information shown for each person.
- **Appearance:** image size/shape, default image, panel header, title, tile background, typography, alignment, and corners control the card design.
- **Advanced:** the provisioning action can create the expected list schema. `overridecss` loads site-specific CSS and `enableDiagnostics` logs list, record, and image operations.

Provisioning creates the list structure, not contact records. Visitors need read access to the selected list and any image locations.

## Common scenarios

- An employee directory filtered by directorate and division.
- A leadership or command roster with photographs and biography links.
- A support directory organized by region, service, or escalation team.
- A small page-specific contact list maintained inline without a separate SharePoint list.
