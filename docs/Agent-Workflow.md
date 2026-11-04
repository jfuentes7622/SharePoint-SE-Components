# GitHub agent workflow: staged pilot

## What is automated

GitHub Actions initializes the six workflow labels on the first issue event or
manual routing run, maintains a routing comment, records maintainer approvals,
and revokes stale approval. PRs run deterministic SE validation. **No label launches
an AI session, assigns Copilot, writes code, or deploys anything in this pilot.**
Maintainers manually select the repository's custom agent in GitHub and provide
the issue/PR URL. Read-only agents return comment-ready reports; a maintainer posts them.
Copilot access, repository policy and agent availability must be enabled in GitHub settings.

## Issue lifecycle

1. Add `triage`; select **SE Issue Triage**. Post its report on the issue.
2. Resolve missing requirements; add `design`; select **SE SPFx Architecture**.
   Post its complete proposal as a separate issue comment.
3. A collaborator with write, maintain or admin permission comments:
   `/approve-design <proposal-comment-id>`. Obtain the numeric ID from the comment's
   permalink (`#issuecomment-123456` means `123456`).
4. The router records a hash of the current issue title/body and proposal text,
   approves that exact snapshot, and adds `design-approved`. Review the successful
   routing run before proceeding.
5. A maintainer adds `coding` and manually assigns **SE Approved Implementation**
   the issue, proposal and approval links. Approval does not automatically start work.
6. Open a PR; add `review` / `spfx-validation` on the issue as routing requests and manually select
   the relevant specialist with the PR link. Validation CI does not depend on labels.
7. Human review, required checks and a test-farm deployment precede merging/releasing.

Editing the issue title/body or approved proposal, deleting the proposal, or removing
`design-approved` revokes approval and removes `coding`. Reapprove the new snapshot.
An unauthorized approval/coding request is rejected. Do not edit bot-owned approval
comments. Approval invalidation occurs asynchronously: wait for routing to finish and
recheck the snapshot immediately before coding. Profiles are guidance, not an execution
sandbox; no automatic coding job is configured.

## Branches and permissions

Preferred branches: `feature/<issue>-<slug>`, `bugfix/<issue>-<slug>`,
`refactor/<issue>-<slug>`. Cloud agents may require their own prefix (such as `copilot/`);
keep that unprotected working branch and document the exception.

Configure GitHub rulesets for `main`, `develop`, `release`, and `release/*`:
require PRs, human approval, resolved conversations, and **SE validation required**.
Disallow force pushes/deletion and direct bot pushes; restrict bypass actors.
Do not grant agents administrative credentials. Repository instructions do not enforce
branch protection. Settings are an administrator activation step, not applied by this PR.

The router has contents-read/issues-write permissions and executes only default-branch
code, never issue content. PR validation has contents-read only, no deployment secrets,
and no `pull_request_target` execution. Do not put credentials in issues or agent prompts.

## Activation and next steps

1. Review and merge the profiles, scripts, docs and workflows into the default branch.
2. Confirm Actions and Copilot/custom agents are enabled for this repository/account.
3. Open a test issue or dispatch **SE agent routing** with an existing issue number.
   Confirm labels/routing and test approval, editing and revocation without coding.
4. Run **SE PR validation** manually to build all 19 projects. Resolve any existing
   restore/build/package blockers explicitly without changing SPFx 1.5.1.
5. Add **SE validation required** to branch rulesets once the workflow has run and the
   check is available; verify a deliberately failing test PR cannot merge.
6. Pilot a small approved change using manual agent selection and test-farm deployment.
7. Record the actual farm build/CU and deployment evidence. Only then consider
   event-triggered AI triage/design and, separately, autonomous coding credentials.

Custom profiles live in [the agents directory](../.github/agents).
The pilot router handles issues only, not PR comments/labels. Supply the PR link to the
review/validation specialist; every PR gets validation CI independently of routing.
Cloud agents do not support VS Code's `handoffs` as an orchestration mechanism;
the explicit workflow and maintainer handoffs here are intentional.
