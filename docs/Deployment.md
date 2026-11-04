# Build and deployment validation

Keep [SPFx 1.5.1](SPFx-Version.md). Never regenerate component/solution IDs.

## Repository checks (Node 22+ and PowerShell 7)

From the repository root:

```powershell
npm test
npm run validate-se
$base = git rev-parse origin/main
npm run validate-se -- --base $base
```

Use the actual PR base commit if the PR does not target `main`.
The base comparison preserves IDs and deployment flags, disallows version regressions
and removal/renaming of existing project/component manifests, and prevents adding known
prohibited API patterns. Without `--base`, only current-state invariants are checked.
Intentional component retirement or a manifest move requires a separately reviewed
validator change; never work around an ID check by generating replacement IDs.

Checks include dependency/lockfile SPFx baseline, duplicate IDs within each namespace,
solution/feature/component versions, source manifest references, serve port, and CDN
configuration shape. Existing generated CDN placeholders are valid source configuration.
Legacy SPFx transitive modules can have older versions; direct SPFx dependencies must
resolve to 1.5.1. Static checks do not resolve every TypeScript import: the build does.

## Component builds (isolated Node 10.24.1 / npm 6.14.12)

In the affected project, with its checked-in v1 lockfile:

```powershell
npm ci
npm test # if provided; use npm run test:unit for DynamicForms
npx gulp clean
npx gulp build
npx gulp bundle --ship
npx gulp package-solution --ship
```

Check each exit code. Do not substitute non-production bundling or an editor error
list for the ship build. CI runs affected projects, or all projects when shared validation
changes; manual workflow dispatch runs all projects. The old runtime must not host
privileged persistent services.

Some legacy `gulp test` tasks report **0 tests completed** and a Karma warning while
exiting zero (observed in Accordion). This is not passing component test coverage.
Read the test output and report missing tests separately; root validation tests and
the existing documented regression tests do not establish full component coverage.

From the root after a fresh build:

```powershell
.\scripts\validate-package.ps1 -Project Accordion
```

The package inspector parses XML with DTDs prohibited, checks packaged AppManifest
solution ID/version and configured feature ID/version against generated feature XML,
checks embedded component IDs/versions, SPFx resource versions and bundled asset
references, and verifies that the package exists and opens as a ZIP. CI uploads the freshly built
`.sppkg` as an artifact; validation never publishes a package or deploys it.
Do not mistake validation of a checked-in package for a fresh-build check.

## On-premises acceptance

- Record farm build/CU, App Catalog scope and target page/site.
- Deploy the freshly built artifact to a test App Catalog using an authorized operator.
- Verify feature activation and `skipFeatureDeployment` behavior, asset loading,
  permissions, REST behavior, relevant classic/modern pages and existing component instances.
- Exercise issue acceptance criteria, diagnostics and upgrade from the installed version.
- Record results and a rollback package/version in the PR before production release.

CI checks cannot certify farm-level deployment compatibility or replace this checklist.
The existing release-publishing workflow remains separate and unchanged.
