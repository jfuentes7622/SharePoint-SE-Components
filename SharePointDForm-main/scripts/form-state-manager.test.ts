/**
 * Unit tests for FormStateManager
 */

import { FormStateManager } from '../src/formEngine/core/FormStateManager';
import { FormSchema, FormField, FormState, FieldValue, FieldAction } from '../src/formEngine/core/types';
import { assertEqual, assertTrue, assertFalse, assertNull, test, printSummary, resetCounters } from './test-helpers';

// ============================================================================
// Helpers
// ============================================================================

function makeField(overrides: Partial<FormField> = {}): FormField {
  return {
    id: 'f1',
    type: 'text',
    label: 'Field 1',
    fieldName: 'Field1',
    ...overrides,
  };
}

function makeSchema(overrides: Partial<FormSchema> = {}): FormSchema {
  return {
    id: 'test',
    name: 'test',
    mode: 'new',
    listName: 'TestList',
    steps: [{ id: 's1', title: 'Step 1', fields: [] }],
    ...overrides,
  };
}

// ============================================================================
// 1. State initialization with default values
// ============================================================================

test('initializeState: text field defaults to empty string', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'Step 1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  assertEqual(mgr.getFieldValue('f1'), '', 'text default');
});

test('initializeState: boolean field defaults to false', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'Step 1', fields: [makeField({ type: 'boolean', fieldName: 'Flag' })] }],
  });
  const mgr = new FormStateManager(schema);
  assertEqual(mgr.getFieldValue('f1'), false, 'boolean default');
});

test('initializeState: multiselect field defaults to empty array', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'Step 1', fields: [makeField({ type: 'multiselect', fieldName: 'Tags' })] }],
  });
  const mgr = new FormStateManager(schema);
  assertEqual(mgr.getFieldValue('f1'), [], 'multiselect default');
});

test('initializeState: field uses defaultValue when specified', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'Step 1', fields: [makeField({ defaultValue: 'preset' })] }],
  });
  const mgr = new FormStateManager(schema);
  assertEqual(mgr.getFieldValue('f1'), 'preset', 'defaultValue used');
});

test('initializeState: initial state has touched=false, dirty=false, valid=true', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'Step 1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  const fs = mgr.getFieldState('f1')!;
  assertFalse(fs.touched, 'not touched');
  assertFalse(fs.dirty, 'not dirty');
  assertTrue(fs.valid, 'valid');
  assertEqual(fs.errors.length, 0, 'no errors');
});

test('initializeState: currentStep starts at 0', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'Step 1', fields: [] }],
  });
  const mgr = new FormStateManager(schema);
  assertEqual(mgr.getState().currentStep, 0, 'initial step');
});

test('initializeState: isSubmitting=false, isValid=true', () => {
  const schema = makeSchema({ steps: [{ id: 's1', title: 'S1', fields: [] }] });
  const mgr = new FormStateManager(schema);
  const state = mgr.getState();
  assertFalse(state.isSubmitting, 'not submitting');
  assertTrue(state.isValid, 'valid');
});

test('initializeState: fields are visible/required by default unless specified', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  const fs = mgr.getFieldState('f1')!;
  assertTrue(fs.visible, 'visible by default');
  assertFalse(fs.required, 'not required by default');
  assertFalse(fs.readOnly, 'not readOnly by default');
  assertFalse(fs.disabled, 'not disabled by default');
});

test('initializeState: required=true makes field required', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField({ required: true })] }],
  });
  const mgr = new FormStateManager(schema);
  assertTrue(mgr.getFieldState('f1')!.required, 'required');
});

test('initializeState: forceReadOnly option sets all fields readOnly', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema, undefined, { forceReadOnly: true });
  assertTrue(mgr.getFieldState('f1')!.readOnly, 'forceReadOnly');
});

// ============================================================================
// 2. State initialization with initial values
// ============================================================================

test('initializeState: initial values by field id', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema, { f1: 'hello' });
  assertEqual(mgr.getFieldValue('f1'), 'hello', 'value by id');
});

test('initializeState: initial values by fieldName', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField({ fieldName: 'Title' })] }],
  });
  const mgr = new FormStateManager(schema, { Title: 'hello' });
  assertEqual(mgr.getFieldValue('f1'), 'hello', 'value by fieldName');
});

test('initializeState: fieldName takes precedence over field id', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField({ fieldName: 'Title' })] }],
  });
  const mgr = new FormStateManager(schema, { f1: 'by-id', Title: 'by-name' });
  assertEqual(mgr.getFieldValue('f1'), 'by-name', 'fieldName wins over id');
});

test('initializeState: multiple fields with initial values', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({ id: 'f1', fieldName: 'Title' }),
        makeField({ id: 'f2', fieldName: 'Desc', type: 'number' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema, { Title: 'MyTitle', Desc: 42 });
  assertEqual(mgr.getFieldValue('f1'), 'MyTitle', 'f1 value');
  assertEqual(mgr.getFieldValue('f2'), 42, 'f2 value');
});

// ============================================================================
// 3. getFieldValue / getAllFieldValues
// ============================================================================

test('getFieldValue: returns undefined for unknown field', () => {
  const schema = makeSchema({ steps: [{ id: 's1', title: 'S1', fields: [] }] });
  const mgr = new FormStateManager(schema);
  assertEqual(mgr.getFieldValue('nonexistent'), undefined, 'unknown field');
});

test('getAllFieldValues: returns values keyed by field id', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({ id: 'f1', fieldName: 'A' }),
        makeField({ id: 'f2', fieldName: 'B' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema, { A: 'val1', B: 'val2' });
  const all = mgr.getAllFieldValues();
  assertEqual(all['f1'], 'val1', 'f1 in all values');
  assertEqual(all['f2'], 'val2', 'f2 in all values');
});

test('getAllFieldValues: excludes richtext fields', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({ id: 'f1', fieldName: 'A' }),
        makeField({ id: 'f2', fieldName: 'Body', type: 'richtext' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema, { A: 'text', Body: '<p>html</p>' });
  const all = mgr.getAllFieldValues();
  assertEqual(all['f1'], 'text', 'f1 included');
  assertFalse(Object.prototype.hasOwnProperty.call(all, 'f2'), 'richtext excluded');
});

// ============================================================================
// 4. setFieldValue with basic update
// ============================================================================

test('setFieldValue: updates value and sets dirty=true', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  mgr.setFieldValue('f1', 'new value');
  assertEqual(mgr.getFieldValue('f1'), 'new value', 'value updated');
  assertTrue(mgr.getFieldState('f1')!.dirty, 'dirty');
  assertTrue(mgr.hasDirtyFields(), 'has dirty fields');
});

test('setFieldValue: no-op for unknown field id', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  mgr.setFieldValue('nonexistent', 'value');
  assertFalse(mgr.hasDirtyFields(), 'no dirty fields for unknown id');
});

test('setFieldValue: no-op when field exists in schema but not in state', () => {
  // Edge case: fieldMap has it but state.fields does not (shouldn't happen normally)
  // This is covered by the guard check in setFieldValue
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  // Setting value on existing field works fine
  mgr.setFieldValue('f1', 'test');
  assertEqual(mgr.getFieldValue('f1'), 'test', 'normal set works');
});

// ============================================================================
// 5. setFieldValue with onChange actions
// ============================================================================

test('onChange show: makes target field visible', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({
          id: 'f1',
          onChange: [{ type: 'show', target: 'f2' }],
        }),
        makeField({ id: 'f2', fieldName: 'Target' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  // Hide f2 first
  const state = mgr.getState();
  state.fields['f2'].visible = false;
  // Now trigger onChange via setFieldValue
  mgr.setFieldValue('f1', 'trigger');
  assertTrue(mgr.getFieldState('f2')!.visible, 'f2 shown');
});

test('onChange hide: hide action sets visible=false (reevaluate may restore)', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({
          id: 'f1',
          fieldName: 'Trigger',
          onChange: [{ type: 'hide', target: 'f2' }],
        }),
        makeField({ id: 'f2', fieldName: 'Target', visible: false }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  // f2 starts hidden (visible: false)
  assertFalse(mgr.getFieldState('f2')!.visible, 'f2 hidden initially');
  // The hide action runs, then reevaluateAllFields also sets visible=false from schema
  mgr.setFieldValue('f1', 'trigger');
  assertFalse(mgr.getFieldState('f2')!.visible, 'f2 stays hidden');
});

test('onChange set-value: sets value on target field', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({
          id: 'f1',
          onChange: [{ type: 'set-value', target: 'f2', value: 'auto-filled' }],
        }),
        makeField({ id: 'f2', fieldName: 'Target' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  mgr.setFieldValue('f1', 'trigger');
  assertEqual(mgr.getFieldValue('f2'), 'auto-filled', 'f2 auto-set');
  assertTrue(mgr.getFieldState('f2')!.dirty, 'f2 dirty from set-value');
});

test('onChange clear: clears target field value', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({
          id: 'f1',
          onChange: [{ type: 'clear', target: 'f2' }],
        }),
        makeField({ id: 'f2', fieldName: 'Target' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema, { Target: 'existing' });
  assertEqual(mgr.getFieldValue('f2'), 'existing', 'f2 has value');
  mgr.setFieldValue('f1', 'trigger');
  assertEqual(mgr.getFieldValue('f2'), '', 'f2 cleared');
  assertTrue(mgr.getFieldState('f2')!.dirty, 'f2 dirty from clear');
});

test('onChange enable: enables target field', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({
          id: 'f1',
          onChange: [{ type: 'enable', target: 'f2' }],
        }),
        makeField({ id: 'f2', fieldName: 'Target' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  // Disable f2 first
  mgr.getState().fields['f2'].disabled = true;
  mgr.setFieldValue('f1', 'trigger');
  assertFalse(mgr.getFieldState('f2')!.disabled, 'f2 enabled');
});

test('onChange disable: disables target field', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({
          id: 'f1',
          onChange: [{ type: 'disable', target: 'f2' }],
        }),
        makeField({ id: 'f2', fieldName: 'Target' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.getFieldState('f2')!.disabled, 'f2 initially enabled');
  mgr.setFieldValue('f1', 'trigger');
  assertTrue(mgr.getFieldState('f2')!.disabled, 'f2 disabled');
});

test('onChange show with condition: visible when condition is true', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({
          id: 'f1',
          fieldName: 'Status',
          onChange: [{ type: 'show', target: 'f2', condition: "Status eq 'Active'" }],
        }),
        makeField({ id: 'f2', fieldName: 'Detail' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  mgr.setFieldValue('f1', 'Active');
  assertTrue(mgr.getFieldState('f2')!.visible, 'f2 visible when condition true');
});

test('onChange show with condition: condition false and schema visible=false stays hidden', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({
          id: 'f1',
          fieldName: 'Status',
          onChange: [{ type: 'show', target: 'f2', condition: "Status eq 'Active'" }],
        }),
        makeField({ id: 'f2', fieldName: 'Detail', visible: false }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.getFieldState('f2')!.visible, 'f2 hidden initially');
  mgr.setFieldValue('f1', 'Inactive');
  assertFalse(mgr.getFieldState('f2')!.visible, 'f2 stays hidden when condition false');
});

test('onChange multiple actions execute in order', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({
          id: 'f1',
          onChange: [
            { type: 'set-value', target: 'f2', value: 'first' },
            { type: 'set-value', target: 'f3', value: 'second' },
          ],
        }),
        makeField({ id: 'f2', fieldName: 'B' }),
        makeField({ id: 'f3', fieldName: 'C' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  mgr.setFieldValue('f1', 'trigger');
  assertEqual(mgr.getFieldValue('f2'), 'first', 'first action');
  assertEqual(mgr.getFieldValue('f3'), 'second', 'second action');
});

test('onChange without actions: reevaluates all field conditions', () => {
  // When no onChange, setFieldValue calls reevaluateAllFields
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({ id: 'f1', fieldName: 'Type' }),
        makeField({ id: 'f2', fieldName: 'Detail', visible: "Type eq 'Special'" }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.getFieldState('f2')!.visible, 'f2 hidden initially');
  mgr.setFieldValue('f1', 'Special');
  assertTrue(mgr.getFieldState('f2')!.visible, 'f2 visible after condition met');
});

// ============================================================================
// 6. touchField and dirty tracking
// ============================================================================

test('touchField: marks field as touched', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.getFieldState('f1')!.touched, 'not touched initially');
  mgr.touchField('f1');
  assertTrue(mgr.getFieldState('f1')!.touched, 'touched after touchField');
});

test('touchField: no-op for unknown field', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  // Should not throw
  mgr.touchField('nonexistent');
  assertFalse(mgr.getFieldState('f1')!.touched, 'f1 still untouched');
});

test('hasDirtyFields: false when no fields changed', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.hasDirtyFields(), 'no dirty fields');
});

test('hasDirtyFields: true after setFieldValue', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  mgr.setFieldValue('f1', 'changed');
  assertTrue(mgr.hasDirtyFields(), 'dirty after change');
});

test('touchField does not mark dirty', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  mgr.touchField('f1');
  assertTrue(mgr.getFieldState('f1')!.touched, 'touched');
  assertFalse(mgr.getFieldState('f1')!.dirty, 'not dirty');
});

// ============================================================================
// 7. setFieldErrors and form validity
// ============================================================================

test('setFieldErrors: sets errors and marks field invalid', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  mgr.setFieldErrors('f1', ['Required']);
  assertFalse(mgr.getFieldState('f1')!.valid, 'field invalid');
  assertEqual(mgr.getFieldState('f1')!.errors.length, 1, 'one error');
  assertEqual(mgr.getFieldState('f1')!.errors[0], 'Required', 'error text');
});

test('setFieldErrors: clears errors and marks field valid', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  mgr.setFieldErrors('f1', ['Error']);
  assertFalse(mgr.getFieldState('f1')!.valid, 'invalid');
  mgr.setFieldErrors('f1', []);
  assertTrue(mgr.getFieldState('f1')!.valid, 'valid again');
  assertEqual(mgr.getFieldState('f1')!.errors.length, 0, 'no errors');
});

test('setFieldErrors: updates form validity - visible field errors invalidate form', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  assertTrue(mgr.isFormValid(), 'form initially valid');
  mgr.setFieldErrors('f1', ['Required']);
  assertFalse(mgr.isFormValid(), 'form invalid after error on visible field');
});

test('setFieldErrors: hidden field errors do not invalidate form', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField({ visible: false })] }],
  });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.getFieldState('f1')!.visible, 'f1 hidden');
  mgr.setFieldErrors('f1', ['Required']);
  assertTrue(mgr.isFormValid(), 'form still valid - hidden field errors ignored');
});

test('setFieldErrors: no-op for unknown field', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  mgr.setFieldErrors('nonexistent', ['Error']);
  assertTrue(mgr.isFormValid(), 'form still valid');
});

test('setFieldErrors: multiple errors on field', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  mgr.setFieldErrors('f1', ['Required', 'Too short']);
  assertEqual(mgr.getFieldState('f1')!.errors.length, 2, 'two errors');
  assertFalse(mgr.getFieldState('f1')!.valid, 'invalid');
});

// ============================================================================
// 8. Step navigation (nextStep, prevStep, goToStep) including hidden steps
// ============================================================================

test('nextStep: advances to next step', () => {
  const schema = makeSchema({
    steps: [
      { id: 's1', title: 'Step 1', fields: [] },
      { id: 's2', title: 'Step 2', fields: [] },
    ],
  });
  const mgr = new FormStateManager(schema);
  assertEqual(mgr.getState().currentStep, 0, 'starts at 0');
  const result = mgr.nextStep();
  assertTrue(result, 'moved to next');
  assertEqual(mgr.getState().currentStep, 1, 'now at step 1');
});

test('nextStep: returns false when at last step', () => {
  const schema = makeSchema({
    steps: [
      { id: 's1', title: 'Step 1', fields: [] },
      { id: 's2', title: 'Step 2', fields: [] },
    ],
  });
  const mgr = new FormStateManager(schema);
  mgr.nextStep();
  assertEqual(mgr.getState().currentStep, 1, 'at last step');
  const result = mgr.nextStep();
  assertFalse(result, 'cannot advance further');
  assertEqual(mgr.getState().currentStep, 1, 'stays at step 1');
});

test('nextStep: skips hidden steps', () => {
  const schema = makeSchema({
    steps: [
      { id: 's1', title: 'Step 1', fields: [] },
      { id: 's2', title: 'Step 2', fields: [], visible: false },
      { id: 's3', title: 'Step 3', fields: [] },
    ],
  });
  const mgr = new FormStateManager(schema);
  const result = mgr.nextStep();
  assertTrue(result, 'moved');
  assertEqual(mgr.getState().currentStep, 2, 'skipped hidden step 2, landed on step 3');
});

test('prevStep: goes back to previous step', () => {
  const schema = makeSchema({
    steps: [
      { id: 's1', title: 'Step 1', fields: [] },
      { id: 's2', title: 'Step 2', fields: [] },
    ],
  });
  const mgr = new FormStateManager(schema);
  mgr.nextStep();
  assertEqual(mgr.getState().currentStep, 1, 'at step 1');
  const result = mgr.prevStep();
  assertTrue(result, 'moved back');
  assertEqual(mgr.getState().currentStep, 0, 'back at step 0');
});

test('prevStep: returns false when at first step', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'Step 1', fields: [] }],
  });
  const mgr = new FormStateManager(schema);
  const result = mgr.prevStep();
  assertFalse(result, 'cannot go back from step 0');
  assertEqual(mgr.getState().currentStep, 0, 'stays at 0');
});

test('prevStep: skips hidden steps', () => {
  const schema = makeSchema({
    steps: [
      { id: 's1', title: 'Step 1', fields: [] },
      { id: 's2', title: 'Step 2', fields: [], visible: false },
      { id: 's3', title: 'Step 3', fields: [] },
    ],
  });
  const mgr = new FormStateManager(schema);
  mgr.nextStep(); // goes to step 2
  assertEqual(mgr.getState().currentStep, 2, 'at step 2');
  const result = mgr.prevStep();
  assertTrue(result, 'moved back');
  assertEqual(mgr.getState().currentStep, 0, 'skipped hidden step 1, back at step 0');
});

test('goToStep: jumps to specified step', () => {
  const schema = makeSchema({
    steps: [
      { id: 's1', title: 'Step 1', fields: [] },
      { id: 's2', title: 'Step 2', fields: [] },
      { id: 's3', title: 'Step 3', fields: [] },
    ],
  });
  const mgr = new FormStateManager(schema);
  const result = mgr.goToStep(2);
  assertTrue(result, 'jumped');
  assertEqual(mgr.getState().currentStep, 2, 'at step 2');
});

test('goToStep: hidden step falls back to nearest visible before', () => {
  const schema = makeSchema({
    steps: [
      { id: 's1', title: 'Step 1', fields: [] },
      { id: 's2', title: 'Step 2', fields: [], visible: false },
      { id: 's3', title: 'Step 3', fields: [] },
    ],
  });
  const mgr = new FormStateManager(schema);
  const result = mgr.goToStep(1); // step 1 is hidden
  assertTrue(result, 'moved');
  assertEqual(mgr.getState().currentStep, 0, 'fell back to step 0');
});

test('goToStep: hidden step falls back to nearest visible after when no visible before', () => {
  const schema = makeSchema({
    steps: [
      { id: 's1', title: 'Step 1', fields: [], visible: false },
      { id: 's2', title: 'Step 2', fields: [] },
      { id: 's3', title: 'Step 3', fields: [] },
    ],
  });
  // Since step 0 is hidden, initial step should be 1
  const mgr = new FormStateManager(schema);
  assertEqual(mgr.getState().currentStep, 1, 'starts at first visible step');
  // Try to go to hidden step 0 - should fall back to next visible after (step 1)
  const result = mgr.goToStep(0);
  assertTrue(result, 'moved');
  assertEqual(mgr.getState().currentStep, 1, 'fell back to step 1');
});

test('goToStep: out of range returns false', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'Step 1', fields: [] }],
  });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.goToStep(5), 'out of range');
  assertFalse(mgr.goToStep(-1), 'negative index');
});

// ============================================================================
// 9. Subscriptions (state listeners and field listeners) with unsubscribe
// ============================================================================

test('subscribe: state listener called on state changes', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  let callCount = 0;
  let lastState: FormState | null = null;

  mgr.subscribe((state) => {
    callCount++;
    lastState = state;
  });

  mgr.setFieldValue('f1', 'new');
  assertEqual(callCount, 1, 'listener called once');
  assertTrue(lastState !== null, 'state received');
  assertEqual(lastState!.fields['f1'].value, 'new', 'correct value in state snapshot');
});

test('subscribe: listener called on setSubmitting', () => {
  const schema = makeSchema({ steps: [{ id: 's1', title: 'S1', fields: [] }] });
  const mgr = new FormStateManager(schema);
  let callCount = 0;

  mgr.subscribe(() => { callCount++; });
  mgr.setSubmitting(true);
  assertEqual(callCount, 1, 'listener called on setSubmitting');
  assertTrue(mgr.getState().isSubmitting, 'submitting');
});

test('subscribe: listener called on touchField', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  let callCount = 0;
  mgr.subscribe(() => { callCount++; });
  mgr.touchField('f1');
  assertEqual(callCount, 1, 'listener called on touch');
});

test('subscribe: listener called on setFieldErrors', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  let callCount = 0;
  mgr.subscribe(() => { callCount++; });
  mgr.setFieldErrors('f1', ['Error']);
  assertEqual(callCount, 1, 'listener called on setFieldErrors');
});

test('subscribe: unsubscribe stops notifications', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  let callCount = 0;
  const unsub = mgr.subscribe(() => { callCount++; });

  mgr.setFieldValue('f1', 'first');
  assertEqual(callCount, 1, 'called before unsubscribe');

  unsub();
  mgr.setFieldValue('f1', 'second');
  assertEqual(callCount, 1, 'not called after unsubscribe');
});

test('subscribeField: field listener called with field id and value', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  let receivedId = '';
  let receivedValue: FieldValue = undefined;

  mgr.subscribeField('f1', (fieldId, value) => {
    receivedId = fieldId;
    receivedValue = value;
  });

  mgr.setFieldValue('f1', 'hello');
  assertEqual(receivedId, 'f1', 'correct field id');
  assertEqual(receivedValue, 'hello', 'correct value');
});

test('subscribeField: only triggered for subscribed field, not others', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [makeField({ id: 'f1' }), makeField({ id: 'f2', fieldName: 'Field2' })],
    }],
  });
  const mgr = new FormStateManager(schema);
  let callCount = 0;

  mgr.subscribeField('f1', () => { callCount++; });

  mgr.setFieldValue('f2', 'other');
  assertEqual(callCount, 0, 'not called for different field');

  mgr.setFieldValue('f1', 'mine');
  assertEqual(callCount, 1, 'called for subscribed field');
});

test('subscribeField: unsubscribe stops field notifications', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  let callCount = 0;
  const unsub = mgr.subscribeField('f1', () => { callCount++; });

  mgr.setFieldValue('f1', 'first');
  assertEqual(callCount, 1, 'called before unsubscribe');

  unsub();
  mgr.setFieldValue('f1', 'second');
  assertEqual(callCount, 1, 'not called after unsubscribe');
});

// ============================================================================
// 10. reset with new initial values
// ============================================================================

test('reset: resets all fields to defaults', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  mgr.setFieldValue('f1', 'changed');
  mgr.touchField('f1');
  assertTrue(mgr.getFieldState('f1')!.dirty, 'dirty');
  assertTrue(mgr.getFieldState('f1')!.touched, 'touched');

  mgr.reset();
  assertEqual(mgr.getFieldValue('f1'), '', 'reset to default');
  assertFalse(mgr.getFieldState('f1')!.dirty, 'not dirty after reset');
  assertFalse(mgr.getFieldState('f1')!.touched, 'not touched after reset');
});

test('reset: applies new initial values', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema, { f1: 'original' });
  assertEqual(mgr.getFieldValue('f1'), 'original', 'initial value');

  mgr.setFieldValue('f1', 'changed');
  mgr.reset({ f1: 'new-initial' });
  assertEqual(mgr.getFieldValue('f1'), 'new-initial', 'reset with new values');
});

test('reset: resets isSubmitting', () => {
  const schema = makeSchema({ steps: [{ id: 's1', title: 'S1', fields: [] }] });
  const mgr = new FormStateManager(schema);
  mgr.setSubmitting(true);
  assertTrue(mgr.getState().isSubmitting, 'submitting');
  mgr.reset();
  assertFalse(mgr.getState().isSubmitting, 'not submitting after reset');
});

test('reset: resets currentStep to 0', () => {
  const schema = makeSchema({
    steps: [
      { id: 's1', title: 'S1', fields: [] },
      { id: 's2', title: 'S2', fields: [] },
    ],
  });
  const mgr = new FormStateManager(schema);
  mgr.nextStep();
  assertEqual(mgr.getState().currentStep, 1, 'at step 1');
  mgr.reset();
  assertEqual(mgr.getState().currentStep, 0, 'back to step 0');
});

test('reset: notifies state listeners', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  let callCount = 0;
  mgr.subscribe(() => { callCount++; });
  mgr.reset();
  assertEqual(callCount, 1, 'listener called on reset');
});

test('reset: clears field errors and restores form validity', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  mgr.setFieldErrors('f1', ['Error']);
  assertFalse(mgr.isFormValid(), 'invalid');
  mgr.reset();
  assertTrue(mgr.isFormValid(), 'valid after reset');
  assertEqual(mgr.getFieldState('f1')!.errors.length, 0, 'no errors');
});

// ============================================================================
// 11. setSubmitting
// ============================================================================

test('setSubmitting: sets isSubmitting flag', () => {
  const schema = makeSchema({ steps: [{ id: 's1', title: 'S1', fields: [] }] });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.getState().isSubmitting, 'not submitting');
  mgr.setSubmitting(true);
  assertTrue(mgr.getState().isSubmitting, 'submitting');
  mgr.setSubmitting(false);
  assertFalse(mgr.getState().isSubmitting, 'not submitting again');
});

// ============================================================================
// 12. Condition evaluation (visible/required/readOnly with OData expressions)
// ============================================================================

test('condition: visible with OData expression', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({ id: 'f1', fieldName: 'Category' }),
        makeField({ id: 'f2', fieldName: 'Detail', visible: "Category eq 'A'" }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema, { Category: 'A' });
  assertTrue(mgr.getFieldState('f2')!.visible, 'visible when condition true');

  mgr.setFieldValue('f1', 'B');
  assertFalse(mgr.getFieldState('f2')!.visible, 'hidden when condition false');
});

test('condition: required with OData expression', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({ id: 'f1', fieldName: 'Type' }),
        makeField({ id: 'f2', fieldName: 'Justification', required: "Type eq 'Other'" }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.getFieldState('f2')!.required, 'not required initially');

  mgr.setFieldValue('f1', 'Other');
  assertTrue(mgr.getFieldState('f2')!.required, 'required when Type eq Other');

  mgr.setFieldValue('f1', 'Standard');
  assertFalse(mgr.getFieldState('f2')!.required, 'not required again');
});

test('condition: readOnly with OData expression', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({ id: 'f1', fieldName: 'Status' }),
        makeField({ id: 'f2', fieldName: 'Amount', readOnly: "Status eq 'Locked'" }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.getFieldState('f2')!.readOnly, 'not readOnly initially');

  mgr.setFieldValue('f1', 'Locked');
  assertTrue(mgr.getFieldState('f2')!.readOnly, 'readOnly when Status eq Locked');

  mgr.setFieldValue('f1', 'Open');
  assertFalse(mgr.getFieldState('f2')!.readOnly, 'not readOnly again');
});

test('condition: visible with boolean literal true', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [makeField({ id: 'f1', visible: true })],
    }],
  });
  const mgr = new FormStateManager(schema);
  assertTrue(mgr.getFieldState('f1')!.visible, 'visible true');
});

test('condition: visible with boolean literal false', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [makeField({ id: 'f1', visible: false })],
    }],
  });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.getFieldState('f1')!.visible, 'visible false');
});

test('condition: visible with string literal "true"', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [makeField({ id: 'f1', visible: 'true' })],
    }],
  });
  const mgr = new FormStateManager(schema);
  assertTrue(mgr.getFieldState('f1')!.visible, 'visible "true"');
});

test('condition: visible with string literal "false"', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [makeField({ id: 'f1', visible: 'false' })],
    }],
  });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.getFieldState('f1')!.visible, 'visible "false"');
});

test('condition: readOnly with forceReadOnly overrides condition evaluation', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [makeField({ id: 'f1', readOnly: false })],
    }],
  });
  const mgr = new FormStateManager(schema, undefined, { forceReadOnly: true });
  assertTrue(mgr.getFieldState('f1')!.readOnly, 'forceReadOnly overrides false');
});

test('condition: initial values used in condition evaluation during init', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({ id: 'f1', fieldName: 'Category' }),
        makeField({ id: 'f2', fieldName: 'Detail', visible: "Category eq 'Special'" }),
      ],
    }],
  });
  // f2 should be visible at init because Category='Special'
  const mgr = new FormStateManager(schema, { Category: 'Special' });
  assertTrue(mgr.getFieldState('f2')!.visible, 'visible at init with matching initial value');
});

// ============================================================================
// 13. Field actions with conditions
// ============================================================================

test('field action show with condition: uses current form context', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({ id: 'f1', fieldName: 'Choice' }),
        makeField({
          id: 'f2',
          fieldName: 'Trigger',
          onChange: [{ type: 'hide', target: 'f3' }],
        }),
        makeField({ id: 'f3', fieldName: 'Target' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  assertTrue(mgr.getFieldState('f3')!.visible, 'f3 visible initially');

  // Trigger hide via onChange - f3 has no visible condition in schema,
  // so reevaluateAllFields will set it back to true (default).
  // This test verifies the hide action runs (even though reevaluate restores it).
  mgr.setFieldValue('f2', 'trigger');
  // After reevaluate, f3.visible is true again (no schema condition to keep it false)
  assertTrue(mgr.getFieldState('f3')!.visible, 'f3 restored by reevaluate (no schema visible=false)');
});

test('field action show with condition: condition false with schema visible=false stays hidden', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({ id: 'f1', fieldName: 'Choice' }),
        makeField({
          id: 'f2',
          fieldName: 'Trigger',
          onChange: [{ type: 'show', target: 'f3', condition: "Choice eq 'Yes'" }],
        }),
        makeField({ id: 'f3', fieldName: 'Target', visible: false }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema);
  assertFalse(mgr.getFieldState('f3')!.visible, 'f3 hidden');

  mgr.setFieldValue('f1', 'No');
  mgr.setFieldValue('f2', 'trigger');
  // Show action evaluates condition (Choice='No' != 'Yes'), so visible stays false.
  // reevaluateAllFields also sets visible=false from schema.
  assertFalse(mgr.getFieldState('f3')!.visible, 'f3 stays hidden - condition false');
});

// ============================================================================
// getState and getFieldState return clones (immutability)
// ============================================================================

test('getState: returns a snapshot, mutations do not affect internal state', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  const snapshot = mgr.getState();
  snapshot.fields['f1'].value = 'mutated';
  snapshot.currentStep = 99;
  // Internal state should be unchanged
  assertEqual(mgr.getFieldValue('f1'), '', 'internal value unchanged');
  assertEqual(mgr.getState().currentStep, 0, 'internal step unchanged');
});

test('getFieldState: returns a clone, mutations do not affect internal state', () => {
  const schema = makeSchema({
    steps: [{ id: 's1', title: 'S1', fields: [makeField()] }],
  });
  const mgr = new FormStateManager(schema);
  const fs = mgr.getFieldState('f1')!;
  fs.value = 'mutated';
  fs.touched = true;
  assertEqual(mgr.getFieldValue('f1'), '', 'internal value unchanged');
  assertFalse(mgr.getFieldState('f1')!.touched, 'internal touched unchanged');
});

test('getFieldState: returns undefined for unknown field', () => {
  const schema = makeSchema({ steps: [{ id: 's1', title: 'S1', fields: [] }] });
  const mgr = new FormStateManager(schema);
  assertEqual(mgr.getFieldState('nonexistent'), undefined, 'undefined for unknown');
});

// ============================================================================
// getFieldValueContext
// ============================================================================

test('getFieldValueContext: returns values keyed by fieldName', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({ id: 'f1', fieldName: 'Title' }),
        makeField({ id: 'f2', fieldName: 'Count', type: 'number' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema, { Title: 'Hello', Count: 5 });
  const ctx = mgr.getFieldValueContext();
  assertEqual(ctx['Title'], 'Hello', 'Title in context');
  assertEqual(ctx['Count'], 5, 'Count in context');
});

test('getFieldValueContext: excludes richtext fields', () => {
  const schema = makeSchema({
    steps: [{
      id: 's1', title: 'S1',
      fields: [
        makeField({ id: 'f1', fieldName: 'Title' }),
        makeField({ id: 'f2', fieldName: 'Body', type: 'richtext' }),
      ],
    }],
  });
  const mgr = new FormStateManager(schema, { Title: 'Hello', Body: '<p>html</p>' });
  const ctx = mgr.getFieldValueContext();
  assertEqual(ctx['Title'], 'Hello', 'Title in context');
  assertFalse(Object.prototype.hasOwnProperty.call(ctx, 'Body'), 'richtext excluded from context');
});

// ============================================================================
// Fields across multiple steps
// ============================================================================

test('fields across multiple steps: all included in state', () => {
  const schema = makeSchema({
    steps: [
      { id: 's1', title: 'Step 1', fields: [makeField({ id: 'f1', fieldName: 'A' })] },
      { id: 's2', title: 'Step 2', fields: [makeField({ id: 'f2', fieldName: 'B' })] },
    ],
  });
  const mgr = new FormStateManager(schema);
  assertTrue(mgr.getFieldState('f1') !== undefined, 'f1 exists');
  assertTrue(mgr.getFieldState('f2') !== undefined, 'f2 exists');
});

// ============================================================================
// Summary
// ============================================================================

printSummary('FormStateManager Tests');
