'use strict';

const assert = require('node:assert/strict');
const { test } = require('node:test');
const { route, digest, currentApproval, approvalFrom } = require('../../.github/scripts/issue-routing.cjs');

function fixture(permission = 'write') {
  const issue = { number: 1, title: 'Feature', body: 'Criteria', labels: [] };
  const proposal = { id: 100, body: 'Ready for maintainer approval: use REST', user: { login: 'maintainer' } };
  const comments = [proposal];
  const labels = [];
  let nextId = 200;
  const api = {
    get: async () => ({ data: issue }),
    listLabelsForRepo: async () => labels,
    createLabel: async value => { labels.push(value); },
    listComments: async () => comments,
    removeLabel: async value => { issue.labels = issue.labels.filter(label => label.name !== value.name); },
    addLabels: async value => { issue.labels.push(...value.labels.map(name => ({ name }))); },
    updateComment: async value => {
      const comment = comments.find(item => item.id === value.comment_id);
      comment.body = value.body;
      return { data: comment };
    },
    createComment: async value => {
      const comment = { id: nextId++, body: value.body, user: { login: 'github-actions[bot]' } };
      comments.push(comment);
      return { data: comment };
    }
  };
  const github = {
    rest: { issues: api, repos: { getCollaboratorPermissionLevel: async () => ({ data: { permission } }) } },
    paginate: async (method, args) => method(args)
  };
  const context = {
    repo: { owner: 'example', repo: 'components' }, actor: 'maintainer',
    payload: { issue: { number: 1 }, action: 'created', comment: { body: '/approve-design 100' } }
  };
  return { issue, proposal, comments, github, context, core: { info: () => {} } };
}

test('maintainer approval records exact issue and proposal; no coding starts', async () => {
  const f = fixture();
  await route(f);
  assert.ok(currentApproval(approvalFrom(f.comments), f.issue, f.comments));
  assert.ok(f.issue.labels.some(label => label.name === 'design-approved'));
  assert.ok(!f.issue.labels.some(label => label.name === 'coding'));
  assert.equal(f.comments.length, 3);
});

test('outsider cannot approve or authorize coding', async () => {
  const f = fixture('read');
  await route(f);
  assert.equal(approvalFrom(f.comments), undefined);
  f.issue.labels.push({ name: 'coding' });
  f.context.payload = { issue: { number: 1 }, action: 'labeled', label: { name: 'coding' } };
  await route(f);
  assert.ok(!f.issue.labels.some(label => label.name === 'coding'));
});

for (const target of ['issue', 'proposal']) {
  test(`${target} edits revoke design approval and coding request`, async () => {
    const f = fixture();
    await route(f);
    f.issue.labels.push({ name: 'coding' });
    if (target === 'issue') f.issue.body += ' changed';
    else f.proposal.body += ' changed';
    f.context.payload = { issue: { number: 1 }, action: 'edited' };
    await route(f);
    assert.equal(approvalFrom(f.comments).value.revoked, true);
    assert.ok(!f.issue.labels.some(label => ['coding', 'design-approved'].includes(label.name)));
  });
}

test('deleting a proposal invalidates its approval', async () => {
  const f = fixture();
  await route(f);
  f.comments.splice(0, 1);
  f.context.payload = { issue: { number: 1 }, action: 'deleted' };
  await route(f);
  assert.equal(approvalFrom(f.comments).value.revoked, true);
});

test('a forged human approval marker is ignored', () => {
  const f = fixture();
  f.comments.push({ id: 2, user: { login: 'outsider' }, body: '<!-- se-design-approval:{"proposalId":100} -->' });
  assert.equal(approvalFrom(f.comments), undefined);
});

test('approval digest changes with title, issue criteria or proposal text', () => {
  const f = fixture();
  const original = digest(f.issue, f.proposal);
  assert.notEqual(digest({ ...f.issue, title: 'Changed' }, f.proposal), original);
  assert.notEqual(digest(f.issue, { ...f.proposal, body: 'Changed' }), original);
});

test('approval revocation cannot be reversed by simply relabeling', async () => {
  const f = fixture();
  await route(f);
  f.context.payload = { issue: { number: 1 }, action: 'unlabeled', label: { name: 'design-approved' } };
  await route(f);
  f.issue.labels.push({ name: 'design-approved' }, { name: 'coding' });
  f.context.payload = { issue: { number: 1 }, action: 'labeled', label: { name: 'coding' } };
  await route(f);
  assert.ok(!f.issue.labels.some(label => ['coding', 'design-approved'].includes(label.name)));
});

test('routing is idempotent and skips PR comments', async () => {
  const f = fixture();
  f.context.payload = { issue: { number: 1 }, action: 'opened' };
  await route(f);
  await route(f);
  assert.equal(f.comments.filter(comment => comment.body.startsWith('<!-- se-agent-routing -->')).length, 1);
  f.context.payload.issue.pull_request = {};
  await route(f);
  assert.equal(f.comments.length, 2);
});
