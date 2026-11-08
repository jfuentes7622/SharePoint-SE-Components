# SPFx baseline

All components stay on **SPFx 1.5.1**, with the existing legacy TypeScript/React
toolchain in each project. No framework upgrade or downgrade is authorized.

The maintainer reports this baseline has been tested and works in the target
SharePoint Server Subscription Edition environment with its latest installed update.
The exact farm build and CU have not yet been recorded. Do not assume March 2026,
or describe this repository-specific decision as a Microsoft support statement.
Record the actual farm build/CU and deployment evidence before production rollout.

## Separate toolchains

- Repository scripts: Node 22 or later and PowerShell 7 (`pwsh`) for XML/package
  inspection and tests; no third-party npm runtime dependencies.
- SPFx builds: isolated Node 10.24.1 / npm 6.14.12, compatible with the legacy
  Gulp 3 and v1 lockfiles. The CI workflow pins these versions. They are end-of-life:
  use only on ephemeral CI runners, never a production server or a persistent
  privileged runner. Successful restore/build in CI is still a rollout prerequisite.
- Keep each component's dependency manifest and lockfile intact. Never run a
  modern npm install in a component as a toolchain migration workaround.

See [Deployment](Deployment.md) for commands and required evidence.
