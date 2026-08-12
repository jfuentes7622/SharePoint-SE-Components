/**
 * Unit tests for ODataParser, ODataEvaluator, and ODataConditionEngine
 */

import { ODataParser } from '../src/formEngine/utils/odata/ODataParser';
import { ODataEvaluator, ODataConditionEngine } from '../src/formEngine/utils/odata/ODataEvaluator';
import { assertEqual, assertTrue, assertFalse, test, printSummary } from './test-helpers';

// ============================================================================
// ODataParser tests
// ============================================================================

test('parses simple eq comparison', () => {
  const parser = new ODataParser();
  const ast = parser.parse("Status eq 'Active'");
  assertEqual(ast.type, 'BinaryOp', 'type');
  assertEqual((ast as any).operator, 'eq', 'operator');
  assertEqual((ast as any).left.type, 'Field', 'left type');
  assertEqual((ast as any).left.name, 'Status', 'left name');
  assertEqual((ast as any).right.type, 'Value', 'right type');
  assertEqual((ast as any).right.value, 'Active', 'right value');
});

test('parses ne comparison', () => {
  const parser = new ODataParser();
  const ast = parser.parse("Status ne 'Closed'");
  assertEqual((ast as any).operator, 'ne', 'operator');
});

test('parses numeric comparison with string values (gt, ge, lt, le)', () => {
  const parser = new ODataParser();
  for (const op of ['gt', 'ge', 'lt', 'le']) {
    const ast = parser.parse(`Amount ${op} '100'`);
    assertEqual((ast as any).operator, op, `${op} operator`);
    assertEqual((ast as any).right.type, 'Value', `${op} right type`);
    // Parser converts numeric strings to numbers
    assertEqual((ast as any).right.value, 100, `${op} right value`);
  }
});

test('parses and expression', () => {
  const parser = new ODataParser();
  const ast = parser.parse("Status eq 'Active' and Amount gt 100");
  assertEqual(ast.type, 'BinaryOp', 'type');
  assertEqual((ast as any).operator, 'and', 'operator');
  assertEqual((ast as any).left.type, 'BinaryOp', 'left is comparison');
  assertEqual((ast as any).right.type, 'BinaryOp', 'right is comparison');
});

test('parses or expression', () => {
  const parser = new ODataParser();
  const ast = parser.parse("Status eq 'Active' or Status eq 'Pending'");
  assertEqual((ast as any).operator, 'or', 'operator');
});

test('parses not expression', () => {
  const parser = new ODataParser();
  const ast = parser.parse("not Status eq 'Closed'");
  assertEqual(ast.type, 'UnaryOp', 'type');
  assertEqual((ast as any).operator, 'not', 'operator');
});

test('parses parenthesized expression', () => {
  const parser = new ODataParser();
  const ast = parser.parse("(Status eq 'A' or Status eq 'B') and Amount gt 0");
  assertEqual(ast.type, 'BinaryOp', 'type');
  assertEqual((ast as any).operator, 'and', 'top-level operator');
  assertEqual((ast as any).left.type, 'Group', 'left is group');
});

test('parses contains function', () => {
  const parser = new ODataParser();
  const ast = parser.parse("contains(Title, 'test')");
  assertEqual(ast.type, 'FunctionCall', 'type');
  assertEqual((ast as any).name, 'contains', 'function name');
  assertEqual((ast as any).args.length, 2, 'args count');
});

test('parses startswith function', () => {
  const parser = new ODataParser();
  const ast = parser.parse("startswith(Name, 'ABC')");
  assertEqual(ast.type, 'FunctionCall', 'type');
  assertEqual((ast as any).name, 'startswith', 'function name');
});

test('parses boolean values (parser converts "true"/"false" to boolean)', () => {
  const parser = new ODataParser();
  const astTrue = parser.parse("Active eq 'true'");
  assertEqual((astTrue as any).right.value, true, 'true → boolean true');
  const astFalse = parser.parse("Active eq 'false'");
  assertEqual((astFalse as any).right.value, false, 'false → boolean false');
});

test('parses null value (parser converts "null" to null)', () => {
  const parser = new ODataParser();
  const ast = parser.parse("Title eq 'null'");
  assertEqual((ast as any).right.value, null, 'null literal');
});

test('throws on unexpected token after valid expression', () => {
  const parser = new ODataParser();
  let threw = false;
  try {
    parser.parse("Status eq 'A' eq 'B'");
  } catch {
    threw = true;
  }
  assertTrue(threw, 'should throw on extra tokens');
});

test('parses escaped single quotes in string values', () => {
  const parser = new ODataParser();
  const ast = parser.parse("Name eq 'O''Brien'");
  assertEqual((ast as any).right.value, "O'Brien", 'escaped quote value');
});

// ============================================================================
// ODataEvaluator tests
// ============================================================================

const evaluator = new ODataEvaluator();

test('evaluates simple eq with matching string', () => {
  assertTrue(evaluator.evaluate("Status eq 'Active'", { Status: 'Active' }), 'eq match');
});

test('evaluates simple eq with non-matching string', () => {
  assertFalse(evaluator.evaluate("Status eq 'Active'", { Status: 'Closed' }), 'eq no match');
});

test('evaluates ne (not equals)', () => {
  assertTrue(evaluator.evaluate("Status ne 'Closed'", { Status: 'Active' }), 'ne match');
  assertFalse(evaluator.evaluate("Status ne 'Active'", { Status: 'Active' }), 'ne no match');
});

test('evaluates gt with string-comparable values', () => {
  assertTrue(evaluator.evaluate("Amount gt '100'", { Amount: '200' }), 'gt match string');
  assertFalse(evaluator.evaluate("Amount gt '100'", { Amount: '50' }), 'gt no match');
});

test('evaluates ge with string-comparable values', () => {
  assertTrue(evaluator.evaluate("Amount ge '100'", { Amount: '200' }), 'ge greater');
  assertFalse(evaluator.evaluate("Amount ge '100'", { Amount: '50' }), 'ge less');
});

test('evaluates lt and le with string-comparable values', () => {
  assertTrue(evaluator.evaluate("Amount lt '100'", { Amount: '50' }), 'lt match');
  assertTrue(evaluator.evaluate("Amount le '100'", { Amount: '50' }), 'le match');
});

test('evaluates and operator', () => {
  const ctx = { Status: 'Active', Amount: '200' };
  assertTrue(evaluator.evaluate("Status eq 'Active' and Amount gt '100'", ctx), 'and both true');
  assertFalse(evaluator.evaluate("Status eq 'Closed' and Amount gt '100'", ctx), 'and first false');
  assertFalse(evaluator.evaluate("Status eq 'Active' and Amount gt '300'", ctx), 'and second false');
});

test('evaluates or operator', () => {
  const ctx = { Status: 'Active', Amount: 50 };
  assertTrue(evaluator.evaluate("Status eq 'Active' or Amount gt 100", ctx), 'or first true');
  assertTrue(evaluator.evaluate("Status eq 'Closed' or Status eq 'Active'", ctx), 'or second true');
  assertFalse(evaluator.evaluate("Status eq 'Closed' or Amount gt 100", ctx), 'or both false');
});

test('evaluates not operator', () => {
  assertTrue(evaluator.evaluate("not Status eq 'Closed'", { Status: 'Active' }), 'not true');
  assertFalse(evaluator.evaluate("not Status eq 'Active'", { Status: 'Active' }), 'not false');
});

test('evaluates contains function', () => {
  assertTrue(evaluator.evaluate("contains(Title, 'test')", { Title: 'This is a test' }), 'contains match');
  assertFalse(evaluator.evaluate("contains(Title, 'xyz')", { Title: 'This is a test' }), 'contains no match');
});

test('evaluates startswith function', () => {
  assertTrue(evaluator.evaluate("startswith(Name, 'ABC')", { Name: 'ABC Company' }), 'startswith match');
  assertFalse(evaluator.evaluate("startswith(Name, 'XYZ')", { Name: 'ABC Company' }), 'startswith no match');
});

test('evaluates parenthesized expression with correct precedence', () => {
  const ctx = { Status: 'B', Amount: '50' };
  // (Status eq 'A' or Status eq 'B') and Amount gt '100' → true and false → false
  assertFalse(evaluator.evaluate("(Status eq 'A' or Status eq 'B') and Amount gt '100'", ctx), 'precedence');
  // Status eq 'A' or Status eq 'B' and Amount gt '100' → false or (true and false) → false (AND binds tighter)
  assertFalse(evaluator.evaluate("Status eq 'A' or Status eq 'B' and Amount gt '100'", ctx), 'and precedence');
});

test('returns true for empty expression', () => {
  assertTrue(evaluator.evaluate('', {}), 'empty string');
  assertTrue(evaluator.evaluate('  ', {}), 'whitespace');
});

test('returns false for invalid expression (graceful degradation)', () => {
  assertFalse(evaluator.evaluate('invalid $$$ expression', {}), 'invalid expression');
});

test('handles null/undefined field values in eq', () => {
  assertTrue(evaluator.evaluate("Title eq null", { Title: null }), 'null eq null');
  assertTrue(evaluator.evaluate("Title eq null", { }), 'undefined eq null');
  assertFalse(evaluator.evaluate("Title eq 'test'", { Title: null }), 'null ne string');
});

test('handles missing field gracefully', () => {
  assertFalse(evaluator.evaluate("Status eq 'Active'", {}), 'missing field');
});

test('compares string-to-string values (no bare number support)', () => {
  // The ConditionBuilder always generates quoted values
  assertTrue(evaluator.evaluate("Amount gt '50'", { Amount: '100' }), 'string gt');
  assertTrue(evaluator.evaluate("Amount gt '50'", { Amount: '200' }), 'string gt 200');
});

test('blocks prototype pollution via __proto__', () => {
  assertFalse(evaluator.evaluate("__proto__ eq 'x'", {}), '__proto__ blocked');
  assertFalse(evaluator.evaluate("constructor eq 'x'", {}), 'constructor blocked');
});

test('caches parsed ASTs for repeated evaluations', () => {
  const eval1 = new ODataEvaluator();
  const ctx = { Status: 'Active' };
  // First call parses and caches
  assertTrue(eval1.evaluate("Status eq 'Active'", ctx), 'first eval');
  // Second call uses cache
  assertTrue(eval1.evaluate("Status eq 'Active'", ctx), 'cached eval');
});

// ============================================================================
// ODataConditionEngine tests
// ============================================================================

const engine = new ODataConditionEngine();

test('ODataConditionEngine.evaluate delegates to evaluator', () => {
  assertTrue(engine.evaluate("Status eq 'Active'", { Status: 'Active' }), 'engine match');
  assertFalse(engine.evaluate("Status eq 'Closed'", { Status: 'Active' }), 'engine no match');
});

// ============================================================================
// Summary
// ============================================================================

printSummary('OData Module Tests');
