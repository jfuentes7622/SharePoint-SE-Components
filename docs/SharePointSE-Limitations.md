# SharePoint SE constraints

- Target SharePoint Server Subscription Edition first and preserve existing Online
  support. Keep the [SPFx 1.5.1 baseline](SPFx-Version.md).
- Do not add Microsoft Graph clients/toolkits, Adaptive Card Extensions, Viva APIs,
  Teams personal-app integrations or SPO-only API dependencies.
- Reuse existing compatible PnP/REST implementations. A package name or an HTTP
  endpoint alone is not evidence that an API works on the target farm.
- Do not add modern Fluent UI packages or modern SPFx modules to legacy projects.
- Preserve solution/component/feature IDs, `skipFeatureDeployment` and app scope.
  Inspect generated feature XML; do not assume every package has a source feature.xml.

`validate-se` checks known prohibited import/endpoint patterns and dependency
regressions. It is not a comprehensive API compatibility analyzer. Architecture
review and testing on the target SharePoint farm remain mandatory. Existing
exceptions, if any, are not permission to introduce more unsupported APIs.

Never automatically deploy from untrusted PR code. A test-farm deployment must use
a separate authorized process and credentials, not the public validation workflow.
