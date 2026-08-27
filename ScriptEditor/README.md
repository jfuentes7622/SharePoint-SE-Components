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

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
