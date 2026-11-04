'use strict';

const { spawnSync } = require('node:child_process');
const path = require('node:path');
const { projects } = require('./validate-se');

function affected(names, files) {
  const selected = new Set();
  for (const file of files) {
    const first = file.split('/')[0];
    if (names.includes(first)) selected.add(first);
    else if (!file.startsWith('docs/') && !file.startsWith('.github/agents/') &&
      !/\.md$/i.test(file) && file !== '.github/workflows/agent-routing.yml' &&
      !file.startsWith('.github/scripts/')) return names;
  }
  return [...selected].sort();
}

if (require.main === module) {
  const root = path.resolve(__dirname, '..');
  const base = process.argv[2];
  const names = projects(root);
  if (!base) console.log(JSON.stringify(names));
  else {
    if (!/^[0-9a-f]{7,40}$/i.test(base)) throw new Error('Expected base commit SHA');
    const result = spawnSync('git', ['diff', '--name-only', '-z', base, 'HEAD'], { cwd: root, encoding: 'utf8' });
    if (result.error) throw result.error;
    if (result.status !== 0) throw new Error(result.stderr);
    console.log(JSON.stringify(affected(names, result.stdout.split('\0').filter(Boolean))));
  }
}

module.exports = { affected };
