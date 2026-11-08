'use strict';

const crypto = require('node:crypto');
const approvalMarker = '<!-- se-design-approval:';
const routeMarker = '<!-- se-agent-routing -->';
const stages = ['triage', 'design', 'coding', 'review', 'spfx-validation'];
const labels = {
  triage: ['bfd4f2', 'Analyze the issue only; no implementation'],
  design: ['d4c5f9', 'Prepare an SE-compatible implementation proposal'],
  coding: ['0e8a16', 'Maintainer requests manual implementation of approved design'],
  review: ['fbca04', 'Request manual specialist PR review'],
  'spfx-validation': ['1d76db', 'Request SE validation; PR CI always runs regardless of label'],
  'design-approved': ['0e8a16', 'Automation-managed current design approval']
};

function digest(issue, proposal) {
  return crypto.createHash('sha256').update(JSON.stringify([
    issue.title, issue.body || '', proposal.id, proposal.body || ''
  ])).digest('hex');
}

function approvalFrom(comments) {
  for (const comment of [...comments].reverse()) {
    if (comment.user.login !== 'github-actions[bot]' || !comment.body.startsWith(approvalMarker)) continue;
    const match = comment.body.match(/^<!-- se-design-approval:(\{[^\n]*\}) -->/);
    if (!match) throw new Error('Malformed automation approval comment');
    return { comment, value: JSON.parse(match[1]) };
  }
  return undefined;
}

function currentApproval(approval, issue, comments) {
  if (!approval || approval.value.revoked) return false;
  const proposal = comments.find(comment => comment.id === approval.value.proposalId);
  return !!proposal && approval.value.digest === digest(issue, proposal);
}

async function route({ github, context, core }) {
  const { owner, repo } = context.repo;
  if (context.payload.issue && context.payload.issue.pull_request) return;
  const number = context.payload.issue ? context.payload.issue.number : Number(context.payload.inputs.issue);
  if (!Number.isSafeInteger(number) || number <= 0) throw new Error('Expected a positive issue number');
  const api = github.rest.issues;
  const issue = (await api.get({ owner, repo, issue_number: number })).data;
  if (issue.pull_request) throw new Error('Issue routing does not accept a pull request number');
  const permission = (await github.rest.repos.getCollaboratorPermissionLevel({
    owner, repo, username: context.actor
  })).data.permission;
  const authorized = ['admin', 'maintain', 'write'].includes(permission);
  const existing = await github.paginate(api.listLabelsForRepo, { owner, repo, per_page: 100 });
  for (const [name, [color, description]] of Object.entries(labels)) {
    if (existing.some(label => label.name === name)) continue;
    try {
      await api.createLabel({ owner, repo, name, color, description });
    } catch (error) {
      if (error.status !== 422 || !error.response || !error.response.data ||
        !Array.isArray(error.response.data.errors) ||
        !error.response.data.errors.some(item => item.code === 'already_exists')) throw error;
      core.info(`Label ${name} was initialized by another issue run.`);
    }
  }
  const comments = await github.paginate(api.listComments, { owner, repo, issue_number: number, per_page: 100 });
  let approval = approvalFrom(comments);
  const remove = async name => {
    if (issue.labels.some(label => label.name === name)) {
      await api.removeLabel({ owner, repo, issue_number: number, name });
      issue.labels = issue.labels.filter(label => label.name !== name);
    }
  };
  const record = async value => {
    const body = `${approvalMarker}${JSON.stringify(value)} -->\n` +
      (value.revoked ? 'Design approval revoked. Reapprove the current proposal before coding.' :
        `Design comment ${value.proposalId} approved by @${value.approvedBy}. This approval applies only to the current issue and proposal text.`);
    const data = approval
      ? (await api.updateComment({ owner, repo, comment_id: approval.comment.id, body })).data
      : (await api.createComment({ owner, repo, issue_number: number, body })).data;
    approval = { comment: data, value };
  };
  const eventComment = context.payload.comment;
  const command = context.payload.action === 'created' && eventComment &&
    eventComment.body.trim().match(/^\/approve-design ([1-9]\d*)$/);
  let message = '';
  if (command) {
    if (!authorized) message = 'Design approval denied: write/maintain/admin permission is required.';
    else {
      const proposal = comments.find(comment => String(comment.id) === command[1]);
      if (!proposal || !proposal.body.trim() || proposal.body.startsWith(approvalMarker) || proposal.body.startsWith(routeMarker)) {
        throw new Error('Approval must reference a nonempty design proposal comment on this issue');
      }
      await record({ proposalId: proposal.id, digest: digest(issue, proposal), approvedBy: context.actor });
      await remove('coding');
      await api.addLabels({ owner, repo, issue_number: number, labels: ['design-approved'] });
      message = 'Design approval recorded. A maintainer can now add `coding` and manually select SE Approved Implementation.';
    }
  }
  if (!currentApproval(approval, issue, comments)) {
    if (approval && !approval.value.revoked) await record({ ...approval.value, revoked: true });
    await remove('design-approved');
    await remove('coding');
  } else if (context.payload.action === 'labeled' && context.payload.label.name === 'design-approved' && !authorized) {
    await record({ ...approval.value, revoked: true });
    await remove('design-approved');
    await remove('coding');
  }
  if (context.payload.action === 'unlabeled' && context.payload.label.name === 'design-approved') {
    if (approval && !approval.value.revoked) await record({ ...approval.value, revoked: true });
    await remove('coding');
  }
  if (context.payload.action === 'labeled' && context.payload.label.name === 'coding' && !authorized) {
    await remove('coding');
    message = 'Coding request denied: a maintainer must apply the coding label.';
  }
  const requested = issue.labels.filter(label => stages.includes(label.name)).map(label => label.name);
  const summary = `${routeMarker}\n## SE agent routing (manual pilot)\n\n` +
    `Requested stages: ${requested.length ? requested.map(name => '`' + name + '`').join(', ') : 'none'}.\n\n` +
    'No AI session or coding is started by this workflow. Select the corresponding repository custom agent on GitHub and supply this issue.\n\n' +
    '- `triage`: SE Issue Triage (analysis only).\n' +
    '- `design`: SE SPFx Architecture (proposal only).\n' +
    '- Approval: a maintainer posts `/approve-design <proposal-comment-id>` on this issue.\n' +
    '- `coding`: requires current recorded approval, `design-approved`, and a maintainer request; manually select SE Approved Implementation.\n' +
    '- `review`: SE Pull Request Review; provide the PR link.\n' +
    '- `spfx-validation`: SE Deployment Validation; PR validation is mandatory even without this label.\n\n' +
    'Issue/proposal edits revoke approval. All work stays on SPFx 1.5.1. See `docs/Agent-Workflow.md`.\n\n' + message;
  const previous = comments.find(comment => comment.user.login === 'github-actions[bot]' && comment.body.startsWith(routeMarker));
  if (previous) {
    if (previous.body !== summary) await api.updateComment({ owner, repo, comment_id: previous.id, body: summary });
  } else await api.createComment({ owner, repo, issue_number: number, body: summary });
  core.info(`Routed issue ${number}; no agent execution was requested.`);
}

module.exports = { route, digest, approvalFrom, currentApproval };
