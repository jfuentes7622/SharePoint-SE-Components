/**
 * Unit tests for SharePointValueMapper
 */

import {
  getODataFieldName,
  convertValueForSP,
  extractFieldValue,
  parseTaxonomyString,
  getItemValueByFieldName,
} from '../src/formEngine/data/SharePointValueMapper';
import { assertEqual, assertTrue, assertNull, test, printSummary } from './test-helpers';

// ============================================================================
// getODataFieldName
// ============================================================================

test('getODataFieldName: returns field unchanged for normal names', () => {
  assertEqual(getODataFieldName('Title'), 'Title', 'normal name');
});

test('getODataFieldName: adds OData__ prefix for _-prefixed names', () => {
  assertEqual(getODataFieldName('_MyField'), 'OData__MyField', 'underscore prefix');
});

test('getODataFieldName: does not double-prefix OData__ names', () => {
  assertEqual(getODataFieldName('OData__MyField'), 'OData__MyField', 'already prefixed');
});

test('getODataFieldName: handles empty string', () => {
  assertEqual(getODataFieldName(''), '', 'empty string');
});

// ============================================================================
// convertValueForSP (Form → SharePoint)
// ============================================================================

test('convertValueForSP: null → null', () => {
  assertNull(convertValueForSP(null, { type: 'text' }), 'null');
});

test('convertValueForSP: undefined → null', () => {
  assertNull(convertValueForSP(undefined, { type: 'text' }), 'undefined');
});

test('convertValueForSP: text passthrough', () => {
  assertEqual(convertValueForSP('hello', { type: 'text' }), 'hello', 'text');
});

test('convertValueForSP: multiselect to { results }', () => {
  const result = convertValueForSP(['a', 'b'], { type: 'multiselect' });
  assertEqual(result, { results: ['a', 'b'] }, 'multiselect');
});

test('convertValueForSP: multiselect empty array → null', () => {
  assertNull(convertValueForSP([], { type: 'multiselect' }), 'empty multiselect');
});

test('convertValueForSP: datetime to ISO string', () => {
  const result = convertValueForSP('2024-01-15T00:00:00.000Z', { type: 'datetime' });
  assertEqual(result, new Date('2024-01-15T00:00:00.000Z').toISOString(), 'datetime ISO');
});

test('convertValueForSP: datetime empty → null', () => {
  assertNull(convertValueForSP('', { type: 'datetime' }), 'empty datetime');
});

test('convertValueForSP: number valid → number', () => {
  assertEqual(convertValueForSP('42', { type: 'number' }), 42, 'number from string');
});

test('convertValueForSP: number empty → null', () => {
  assertNull(convertValueForSP('', { type: 'number' }), 'empty number');
});

test('convertValueForSP: boolean passthrough', () => {
  assertEqual(convertValueForSP(true, { type: 'boolean' }), true, 'boolean true');
  assertEqual(convertValueForSP(false, { type: 'boolean' }), false, 'boolean false');
});

test('convertValueForSP: url string → { Url, Description }', () => {
  assertEqual(
    convertValueForSP('https://example.com', { type: 'url' }),
    { Url: 'https://example.com', Description: '' },
    'url string'
  );
});

test('convertValueForSP: url object → { Url, Description }', () => {
  assertEqual(
    convertValueForSP({ url: 'https://example.com', description: 'Example' }, { type: 'url' }),
    { Url: 'https://example.com', Description: 'Example' },
    'url object'
  );
});

test('convertValueForSP: image with serverRelativeUrl → JSON string', () => {
  const result = convertValueForSP({ serverRelativeUrl: '/images/test.jpg', fileName: 'test.jpg' }, { type: 'image' });
  assertEqual(result, JSON.stringify({ serverRelativeUrl: '/images/test.jpg', fileName: 'test.jpg' }), 'image');
});

test('convertValueForSP: image string → JSON string', () => {
  const result = convertValueForSP('/images/test.jpg', { type: 'image' });
  assertEqual(result, JSON.stringify({ serverRelativeUrl: '/images/test.jpg' }), 'image string');
});

test('convertValueForSP: taxonomy single → object', () => {
  const result = convertValueForSP({ Label: 'IT', TermGuid: 'abc-123', WssId: 1 }, { type: 'taxonomy' });
  assertEqual(result, { Label: 'IT', TermGuid: 'abc-123', WssId: 1 }, 'taxonomy single');
});

test('convertValueForSP: taxonomy multiple → array', () => {
  const result = convertValueForSP([
    { Label: 'IT', TermGuid: 'abc-123', WssId: 1 },
    { Label: 'HR', TermGuid: 'def-456', WssId: 2 },
  ], { type: 'taxonomy' });
  assertEqual(result.length, 2, 'taxonomy multiple count');
});

test('convertValueForSP: taxonomy empty array → null', () => {
  assertNull(convertValueForSP([], { type: 'taxonomy' }), 'empty taxonomy');
});

// ============================================================================
// parseTaxonomyString
// ============================================================================

test('parseTaxonomyString: single term WssId;#Label|Guid', () => {
  const result = parseTaxonomyString('1;#IT|abc12300-def0-1234-5678-abcdef123456');
  assertEqual(result, { Label: 'IT', TermGuid: 'abc12300-def0-1234-5678-abcdef123456', WssId: 1 }, 'single term');
});

test('parseTaxonomyString: multiple terms', () => {
  const result = parseTaxonomyString('1;#IT|abc12300-def0-1234-5678-abcdef123456;#2;#HR|def45600-abcd-1234-5678-abcdef123456');
  assertTrue(Array.isArray(result), 'should be array');
  if (Array.isArray(result)) {
    assertEqual(result.length, 2, 'two terms');
  }
});

test('parseTaxonomyString: empty string → undefined', () => {
  assertEqual(parseTaxonomyString(''), undefined, 'empty');
});

test('parseTaxonomyString: plain text → returned as-is', () => {
  assertEqual(parseTaxonomyString('SomeLabel'), 'SomeLabel', 'plain text');
});

// ============================================================================
// getItemValueByFieldName
// ============================================================================

test('getItemValueByFieldName: exact match', () => {
  assertEqual(getItemValueByFieldName({ Title: 'Hello' }, 'Title'), 'Hello', 'exact match');
});

test('getItemValueByFieldName: case-insensitive match', () => {
  assertEqual(getItemValueByFieldName({ title: 'Hello' }, 'Title'), 'Hello', 'case-insensitive');
});

test('getItemValueByFieldName: no match → undefined', () => {
  assertEqual(getItemValueByFieldName({ Title: 'Hello' }, 'Name'), undefined, 'no match');
});

test('getItemValueByFieldName: null item → undefined', () => {
  assertEqual(getItemValueByFieldName(null, 'Title'), undefined, 'null item');
});

test('getItemValueByFieldName: empty fieldName → undefined', () => {
  assertEqual(getItemValueByFieldName({ Title: 'Hello' }, ''), undefined, 'empty field');
});

// ============================================================================
// extractFieldValue (SharePoint → Form)
// ============================================================================

test('extractFieldValue: text passthrough', () => {
  assertEqual(extractFieldValue({ Title: 'Hello' }, 'Title', 'text'), 'Hello', 'text');
});

test('extractFieldValue: null value → undefined', () => {
  assertEqual(extractFieldValue({ Title: null }, 'Title', 'text'), undefined, 'null → undefined');
});

test('extractFieldValue: missing value → undefined', () => {
  assertEqual(extractFieldValue({}, 'Title', 'text'), undefined, 'missing → undefined');
});

test('extractFieldValue: number value', () => {
  assertEqual(extractFieldValue({ Amount: 42 }, 'Amount', 'number'), 42, 'number');
});

test('extractFieldValue: person object passthrough', () => {
  const person = { Id: 1, Title: 'John' };
  assertEqual(extractFieldValue({ Manager: person }, 'Manager', 'person'), person, 'person object');
});

test('extractFieldValue: person from Id fallback', () => {
  assertEqual(extractFieldValue({ ManagerId: 5 }, 'Manager', 'person'), { Id: 5 }, 'person id fallback');
});

test('extractFieldValue: lookup from object', () => {
  const lookup = { Id: 10, Title: 'Item' };
  assertEqual(extractFieldValue({ Ref: lookup }, 'Ref', 'lookup'), lookup, 'lookup object');
});

test('extractFieldValue: multiselect from array', () => {
  assertEqual(extractFieldValue({ Choices: ['a', 'b'] }, 'Choices', 'multiselect'), ['a', 'b'], 'multiselect array');
});

test('extractFieldValue: multiselect from { results }', () => {
  assertEqual(
    extractFieldValue({ Choices: { results: ['a', 'b'] } }, 'Choices', 'multiselect'),
    ['a', 'b'],
    'multiselect results'
  );
});

test('extractFieldValue: multiselect null → []', () => {
  assertEqual(extractFieldValue({ Choices: null }, 'Choices', 'multiselect'), [], 'multiselect null');
});

test('extractFieldValue: boolean true', () => {
  assertEqual(extractFieldValue({ Active: true }, 'Active', 'boolean'), true, 'boolean true');
});

test('extractFieldValue: boolean from string "true"', () => {
  assertEqual(extractFieldValue({ Active: 'true' }, 'Active', 'boolean'), true, 'boolean string');
});

test('extractFieldValue: datetime valid ISO', () => {
  const iso = '2024-01-15T00:00:00.000Z';
  assertEqual(extractFieldValue({ Date: iso }, 'Date', 'datetime'), iso, 'datetime');
});

test('extractFieldValue: datetime invalid → undefined', () => {
  assertEqual(extractFieldValue({ Date: 'not-a-date' }, 'Date', 'datetime'), undefined, 'invalid datetime');
});

test('extractFieldValue: url from SP object', () => {
  assertEqual(
    extractFieldValue({ Link: { Url: 'https://example.com', Description: 'Example' } }, 'Link', 'url'),
    { url: 'https://example.com', description: 'Example' },
    'url SP object'
  );
});

test('extractFieldValue: url from string', () => {
  assertEqual(
    extractFieldValue({ Link: 'https://example.com' }, 'Link', 'url'),
    'https://example.com',
    'url string'
  );
});

test('extractFieldValue: image from SP object', () => {
  const img = { serverRelativeUrl: '/images/test.jpg', fileName: 'test.jpg' };
  assertEqual(extractFieldValue({ Pic: img }, 'Pic', 'image'), img, 'image object');
});

test('extractFieldValue: image from string → { url }', () => {
  assertEqual(
    extractFieldValue({ Pic: '/images/test.jpg' }, 'Pic', 'image'),
    { url: '/images/test.jpg' },
    'image string'
  );
});

// ============================================================================
// Summary
// ============================================================================

printSummary('SharePointValueMapper Tests');
