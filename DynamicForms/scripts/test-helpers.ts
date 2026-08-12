/**
 * Minimal test helpers for ts-node based test scripts.
 */

let _passed = 0;
let _failed = 0;
const _failures: string[] = [];

export function resetCounters(): void {
  _passed = 0;
  _failed = 0;
  _failures.length = 0;
}

export function assertEqual<T>(actual: T, expected: T, label: string): void {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`${label}: expected ${JSON.stringify(expected)} but got ${JSON.stringify(actual)}`);
  }
}

export function assertTrue(value: boolean, label: string): void {
  if (!value) throw new Error(`${label}: expected true but got false`);
}

export function assertFalse(value: boolean, label: string): void {
  if (value) throw new Error(`${label}: expected false but got true`);
}

export function assertNull(value: any, label: string): void {
  if (value !== null) throw new Error(`${label}: expected null but got ${JSON.stringify(value)}`);
}

export function test(name: string, fn: () => void): void {
  try {
    fn();
    _passed++;
  } catch (err: any) {
    _failed++;
    _failures.push(`  FAIL: ${name} — ${err.message}`);
  }
}

export function printSummary(suiteName: string): void {
  console.log(`\n=== ${suiteName} ===`);
  console.log(`Passed: ${_passed}`);
  console.log(`Failed: ${_failed}`);
  if (_failures.length > 0) {
    console.log('\nFailures:');
    _failures.forEach(f => console.log(f));
    (process as any).exit(1);
  } else {
    console.log('All tests passed!\n');
    _passed = 0;
    _failed = 0;
    _failures.length = 0;
  }
}
