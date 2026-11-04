'use strict';

const assert = require('node:assert/strict');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const { test } = require('node:test');

test('package XML, manifest and asset checks reject corrupt or stale artifacts', () => {
  const result = spawnSync('pwsh', ['-NoProfile', '-File', path.join(__dirname, 'validate-package.tests.ps1')], {
    encoding: 'utf8',
    timeout: 60000
  });
  assert.ifError(result.error);
  assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
});
