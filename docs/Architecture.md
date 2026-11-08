# Repository architecture

This repository contains independent top-level SPFx component projects. Each owns
its package manifest/lockfile, source, TypeScript/lint configuration, Gulp build,
SPFx configuration and SharePoint solution package. The root npm package only
provides repository validation; it is not an SPFx application or a workspace migration.

Discover projects by their top-level `package.json` and
`config/package-solution.json`; do not maintain a second hard-coded component list.
Exclude dependencies, old dependency backups and generated outputs from source checks.

Prefer the affected project's existing SharePoint REST/SPHttpClient/PnP patterns.
Preserve its React/UI and diagnostics conventions. Do not copy modern dependencies
from an Online-targeted sibling repository.

Solution IDs, feature IDs and component IDs have different namespaces: a feature
may intentionally share an ID with its component. Duplicate IDs within the same
namespace across projects are invalid. Never regenerate existing IDs to fix a check.
Source component manifests can use `version: "*"`; compare explicit versions and
package/solution/feature versions against the PR base rather than requiring them
all to be equal.

Agent responsibilities and authorization live in [Agent Workflow](Agent-Workflow.md).
The deterministic validation scripts and required CI, not an agent's prose, enforce
the machine-checkable deployment constraints.
