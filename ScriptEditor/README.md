# SPS Script Editor

SPFx 1.5.1 web part for rendering custom HTML and JavaScript on a SharePoint page.

HTML code can be entered directly in the property pane or loaded from a local `.html`, `.htm`, `.txt`, or `.js` file. Selecting a file replaces the current HTML Code value.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

## Usage

1. Confirm that custom script is permitted and restrict page editing to trusted authors.
2. Add **SPS Script Editor** to a page and enter reviewed HTML, CSS, or JavaScript, or load a supported local text/script file.
3. Enable classic page-context compatibility only when legacy code requires it, and remove web-part padding only when the embedded layout needs edge-to-edge rendering.
4. Test the script in page edit and read modes, including navigation and repeated rendering, before publishing.

## Properties and common configuration

- **Content:** the script property stores raw HTML/JavaScript; the file picker accepts supported HTML, text, and JavaScript files as an authoring convenience.
- **Edit experience:** the title identifies the instance while page authors edit it.
- **Compatibility:** the page-context option exposes classic `_spPageContextInfo` data for legacy scripts that require it.
- **Layout:** the remove-padding option modifies ancestor spacing to let embedded content use the web-part area.
- **Appearance:** there are no styling properties; embedded HTML and CSS define the result.

The control deliberately executes unsanitized code and therefore has `requiresCustomScript` enabled. Treat configuration as code: review it, limit author permissions, avoid secrets, and load external assets only from trusted locations.

## Common scenarios

- Embed a trusted internal widget that is not available as an SPFx component.
- Add page-specific HTML and CSS for a controlled legacy experience.
- Run a small approved script that depends on classic SharePoint page context.
- Load a maintained script file into the editor for deployment on a limited set of pages.

## Build and package

```powershell
npm install
gulp clean
gulp bundle --ship
gulp package-solution --ship
```

The deployable package is generated under `sharepoint/solution/`.
