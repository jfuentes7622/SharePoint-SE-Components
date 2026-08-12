# SharePoint Server Subscription Edition Port Roadmap

## Goal

Move SharePoint Dynamic Form from SPFx 1.22 to a SharePoint Server Subscription Edition-compatible SPFx baseline.

## Compatibility target

Microsoft's SPFx compatibility reference limits SharePoint Server Subscription Edition to SPFx v1.0 through v1.5.

## Stage 1

- Freeze the current repo state as the modern baseline.
- Mark the project as an active SE port instead of implying SE compatibility.
- Identify the legacy SPFx build scaffold and dependency set needed for SPFx 1.5.

## Stage 2

- Replace the modern build toolchain with the legacy SPFx/Gulp toolchain.
- Align package versions to the SE-supported SPFx release line.
- Restore a legacy `tsconfig.json` and supporting build config.

## Stage 3

- Refactor React hook-based components into patterns compatible with the legacy React/SPFx baseline.
- Remove or replace any packages that are not available on the target SPFx line.
- Revalidate field rendering, SharePoint REST access, and property pane behavior.

## Stage 4

- Build and package against the SE target.
- Test deployment on a SharePoint Server Subscription Edition farm.
- Document any remaining platform-specific limitations.