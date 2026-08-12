/**
 * Unit tests for ValidationEngine
 */

import { ValidationEngine } from '../src/formEngine/core/ValidationEngine';
import { FormSchema, FormField, ValidationRule } from '../src/formEngine/core/types';
import { assertEqual, assertTrue, assertFalse, test, printSummary } from './test-helpers';

// Helper to create a minimal FormField
function makeField(overrides: Partial<FormField> = {}): FormField {
  return {
    id: 'f1',
    type: 'text',
    label: 'Test Field',
    fieldName: 'TestField',
    ...overrides,
  };
}

// Minimal schema for ValidationEngine
const schema: FormSchema = {
  id: 'test',
  name: 'test',
  mode: 'new',
  listName: 'TestList',
  steps: [{ id: 's1', title: 'Step 1', fields: [] }],
};

// ============================================================================
// Required validation
// ============================================================================

test('required field: empty string → error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({ required: true });
  const errors = engine.validateField(field, '', {});
  assertTrue(errors.length > 0, 'should have error');
  assertTrue(errors[0].includes('Test Field'), 'error should contain field label');
});

test('required field: non-empty string → no error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({ required: true });
  const errors = engine.validateField(field, 'hello', {});
  assertEqual(errors.length, 0, 'no errors');
});

test('required field: null → error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({ required: true });
  const errors = engine.validateField(field, null, {});
  assertTrue(errors.length > 0, 'null should fail required');
});

test('required field: empty array → error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({ required: true });
  const errors = engine.validateField(field, [], {});
  assertTrue(errors.length > 0, 'empty array should fail required');
});

test('required field: non-empty array → no error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({ required: true });
  const errors = engine.validateField(field, ['a'], {});
  assertEqual(errors.length, 0, 'non-empty array passes');
});

test('non-required field: empty → no error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({ required: false });
  const errors = engine.validateField(field, '', {});
  assertEqual(errors.length, 0, 'optional empty is ok');
});

// ============================================================================
// minLength / maxLength validation
// ============================================================================

test('minLength: too short → error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'minLength', value: 5, message: 'Too short' }],
  });
  const errors = engine.validateField(field, 'ab', {});
  assertTrue(errors.length > 0, 'should have minLength error');
});

test('minLength: exact length → no error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'minLength', value: 3, message: 'Too short' }],
  });
  const errors = engine.validateField(field, 'abc', {});
  assertEqual(errors.length, 0, 'exact length passes');
});

test('maxLength: too long → error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'maxLength', value: 3, message: 'Too long' }],
  });
  const errors = engine.validateField(field, 'abcd', {});
  assertTrue(errors.length > 0, 'should have maxLength error');
});

test('maxLength: exact length → no error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'maxLength', value: 3, message: 'Too long' }],
  });
  const errors = engine.validateField(field, 'abc', {});
  assertEqual(errors.length, 0, 'exact length passes');
});

test('minLength/maxLength: empty value → no error (skip when empty)', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'minLength', value: 5, message: 'Too short' }],
  });
  const errors = engine.validateField(field, '', {});
  assertEqual(errors.length, 0, 'empty value skips minLength');
});

// ============================================================================
// min / max validation (numeric)
// ============================================================================

test('min: value below min → error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'min', value: 10, message: 'Too small' }],
  });
  const errors = engine.validateField(field, 5, {});
  assertTrue(errors.length > 0, 'should have min error');
});

test('min: value at min → no error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'min', value: 10, message: 'Too small' }],
  });
  const errors = engine.validateField(field, 10, {});
  assertEqual(errors.length, 0, 'at min passes');
});

test('max: value above max → error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'max', value: 100, message: 'Too big' }],
  });
  const errors = engine.validateField(field, 200, {});
  assertTrue(errors.length > 0, 'should have max error');
});

test('min/max: empty/null/undefined values → skip', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'min', value: 10, message: 'Too small' }],
  });
  assertEqual(engine.validateField(field, null, {}).length, 0, 'null skips');
  assertEqual(engine.validateField(field, undefined, {}).length, 0, 'undefined skips');
  assertEqual(engine.validateField(field, '', {}).length, 0, 'empty skips');
});

test('min: string number value → coerced and validated', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'min', value: 10, message: 'Too small' }],
  });
  assertTrue(engine.validateField(field, '5', {}).length > 0, 'string 5 fails min 10');
  assertEqual(engine.validateField(field, '15', {}).length, 0, 'string 15 passes min 10');
});

// ============================================================================
// pattern validation
// ============================================================================

test('pattern: matching → no error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'pattern', value: '^[A-Z]{3}$', message: 'Invalid code' }],
  });
  assertEqual(engine.validateField(field, 'ABC', {}).length, 0, 'ABC matches');
});

test('pattern: not matching → error', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'pattern', value: '^[A-Z]{3}$', message: 'Invalid code' }],
  });
  assertTrue(engine.validateField(field, 'abc', {}).length > 0, 'abc fails uppercase pattern');
});

test('pattern: empty value → skip', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'pattern', value: '^[A-Z]{3}$', message: 'Invalid code' }],
  });
  assertEqual(engine.validateField(field, '', {}).length, 0, 'empty skips pattern');
});

test('pattern: invalid regex → skip (no crash)', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'pattern', value: '[invalid(', message: 'Bad pattern' }],
  });
  assertEqual(engine.validateField(field, 'anything', {}).length, 0, 'invalid regex skips');
});

test('pattern: caches compiled regex', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{ type: 'pattern', value: '^\\d+$', message: 'Digits only' }],
  });
  engine.validateField(field, '123', {});
  engine.validateField(field, '456', {});
  // Should not throw and should work correctly (cache hit)
  assertEqual(engine.validateField(field, 'abc', {}).length, 1, 'non-digits fail');
});

// ============================================================================
// applyWhen conditional validation
// ============================================================================

test('applyWhen: rule applies when condition is true', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{
      type: 'required',
      message: 'Required when active',
      applyWhen: "Status eq 'Active'",
    }],
  });
  const errors = engine.validateField(field, '', { Status: 'Active' });
  assertTrue(errors.length > 0, 'should error when condition met');
});

test('applyWhen: rule skips when condition is false', () => {
  const engine = new ValidationEngine(schema);
  const field = makeField({
    validation: [{
      type: 'required',
      message: 'Required when active',
      applyWhen: "Status eq 'Active'",
    }],
  });
  const errors = engine.validateField(field, '', { Status: 'Closed' });
  assertEqual(errors.length, 0, 'should skip when condition not met');
});

// ============================================================================
// validateForm (full form)
// ============================================================================

test('validateForm: returns valid=true when all fields pass', () => {
  const formSchema: FormSchema = {
    ...schema,
    steps: [{
      id: 's1', title: 'Step 1',
      fields: [makeField({ id: 'f1' }), makeField({ id: 'f2' })],
    }],
  };
  const engine = new ValidationEngine(formSchema);
  const result = engine.validateForm({ f1: 'val1', f2: 'val2' });
  assertTrue(result.valid, 'form should be valid');
  assertEqual(Object.keys(result.errors).length, 0, 'no errors');
});

test('validateForm: returns valid=false with errors for failing fields', () => {
  const formSchema: FormSchema = {
    ...schema,
    steps: [{
      id: 's1', title: 'Step 1',
      fields: [makeField({ id: 'f1', required: true })],
    }],
  };
  const engine = new ValidationEngine(formSchema);
  const result = engine.validateForm({ f1: '' });
  assertFalse(result.valid, 'form should be invalid');
  assertTrue(result.errors['f1'].length > 0, 'f1 should have errors');
});

test('validateForm: respects visibleFields filter', () => {
  const formSchema: FormSchema = {
    ...schema,
    steps: [{
      id: 's1', title: 'Step 1',
      fields: [
        makeField({ id: 'f1', required: true }),
        makeField({ id: 'f2', required: true }),
      ],
    }],
  };
  const engine = new ValidationEngine(formSchema);
  // Only validate f2 (f1 is hidden)
  const visibleFields = new Set(['f2']);
  const result = engine.validateForm({ f1: '', f2: 'value' }, visibleFields);
  assertTrue(result.valid, 'form valid when hidden fields fail');
});

test('validateForm: skips richtext fields', () => {
  const formSchema: FormSchema = {
    ...schema,
    steps: [{
      id: 's1', title: 'Step 1',
      fields: [makeField({ id: 'f1', type: 'richtext', required: true })],
    }],
  };
  const engine = new ValidationEngine(formSchema);
  const result = engine.validateForm({ f1: '' });
  assertTrue(result.valid, 'richtext fields are skipped');
});

// ============================================================================
// Custom validation messages
// ============================================================================

test('uses custom validation messages when provided', () => {
  const engine = new ValidationEngine(schema, {
    required: (label) => `${label} is mandatory!`,
    minLength: (n) => `At least ${n} chars needed`,
    maxLength: (n) => `Max ${n} chars allowed`,
    min: (n) => `Min is ${n}`,
    max: (n) => `Max is ${n}`,
    pattern: () => 'Wrong format',
  });
  const field = makeField({ required: true });
  const errors = engine.validateField(field, '', {});
  assertTrue(errors[0].includes('mandatory'), 'custom message used');
});

// ============================================================================
// Summary
// ============================================================================

printSummary('ValidationEngine Tests');
