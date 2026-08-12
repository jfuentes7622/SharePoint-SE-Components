export interface IGridValidationField {
  id: string;
  fieldName: string;
  label?: string;
}

export interface IGridAdvancedValidationRule {
  id?: string;
  expression: string;
  message: string;
  targetField?: string;
}

interface IValidationToken {
  type: 'identifier' | 'number' | 'string' | 'operator' | 'paren' | 'comma' | 'eof';
  value: string;
}

function isValueEmpty(value: any): boolean {
  return value === undefined || value === null || value === '' || (Array.isArray(value) && value.length === 0);
}

function toComparableValue(value: any): any {
  if (value instanceof Date) {
    return value.getTime();
  }
  if (typeof value === 'string') {
    var trimmed = value.trim();
    if (!trimmed) { return trimmed; }
    if (!isNaN(Number(trimmed))) { return Number(trimmed); }
    var parsedDate = new Date(trimmed);
    if (!isNaN(parsedDate.getTime())) { return parsedDate.getTime(); }
    return trimmed.toLowerCase();
  }
  return value;
}

function compareValues(left: any, right: any): number {
  var comparableLeft = toComparableValue(left);
  var comparableRight = toComparableValue(right);
  if (comparableLeft === comparableRight) { return 0; }
  if (comparableLeft > comparableRight) { return 1; }
  if (comparableLeft < comparableRight) { return -1; }
  return 0;
}

function tokenize(expression: string): IValidationToken[] {
  var tokens: IValidationToken[] = [];
  var index = 0;
  while (index < expression.length) {
    var character = expression.charAt(index);
    if (/\s/.test(character)) { index += 1; continue; }
    var pair = expression.substr(index, 2);
    if (pair === '&&' || pair === '||' || pair === '==' || pair === '!=' || pair === '>=' || pair === '<=') {
      tokens.push({ type: 'operator', value: pair }); index += 2; continue;
    }
    if (character === '>' || character === '<' || character === '!') {
      tokens.push({ type: 'operator', value: character }); index += 1; continue;
    }
    if (character === '(' || character === ')') {
      tokens.push({ type: 'paren', value: character }); index += 1; continue;
    }
    if (character === ',') {
      tokens.push({ type: 'comma', value: character }); index += 1; continue;
    }
    if (character === '"' || character === '\'') {
      var quote = character;
      var stringValue = '';
      index += 1;
      while (index < expression.length) {
        var stringCharacter = expression.charAt(index);
        if (stringCharacter === '\\' && index + 1 < expression.length) {
          stringValue += expression.charAt(index + 1); index += 2; continue;
        }
        if (stringCharacter === quote) { break; }
        stringValue += stringCharacter;
        index += 1;
      }
      if (index >= expression.length || expression.charAt(index) !== quote) {
        throw new Error('Unterminated string literal in expression.');
      }
      index += 1;
      tokens.push({ type: 'string', value: stringValue });
      continue;
    }
    if (/\d/.test(character)) {
      var numberValue = character;
      index += 1;
      while (index < expression.length && /[\d.]/.test(expression.charAt(index))) {
        numberValue += expression.charAt(index); index += 1;
      }
      tokens.push({ type: 'number', value: numberValue });
      continue;
    }
    if (/[A-Za-z_]/.test(character)) {
      var identifier = character;
      index += 1;
      while (index < expression.length && /[A-Za-z0-9_]/.test(expression.charAt(index))) {
        identifier += expression.charAt(index); index += 1;
      }
      tokens.push({ type: 'identifier', value: identifier });
      continue;
    }
    throw new Error('Unsupported token in expression near: ' + expression.substr(index, 10));
  }
  tokens.push({ type: 'eof', value: '' });
  return tokens;
}

export function evaluateGridValidationExpression(expression: string, fields: IGridValidationField[], values: { [fieldId: string]: any }): any {
  var tokens = tokenize(expression);
  var tokenIndex = 0;
  var fieldsByKey: { [key: string]: IGridValidationField } = {};
  for (var fieldIndex = 0; fieldIndex < fields.length; fieldIndex += 1) {
    var field = fields[fieldIndex];
    if (field.id) { fieldsByKey[field.id.toLowerCase()] = field; }
    if (field.fieldName) { fieldsByKey[field.fieldName.toLowerCase()] = field; }
    if (field.label) { fieldsByKey[field.label.toLowerCase()] = field; }
  }

  var peek = function(): IValidationToken { return tokens[tokenIndex] || { type: 'eof', value: '' }; };
  var consume = function(expectedType?: string, expectedValue?: string): IValidationToken {
    var current = peek();
    if (expectedType && current.type !== expectedType) {
      throw new Error('Unexpected token: expected ' + expectedType + ' but found ' + current.type + '.');
    }
    if (expectedValue && current.value !== expectedValue) {
      throw new Error('Unexpected token: expected "' + expectedValue + '" but found "' + current.value + '".');
    }
    tokenIndex += 1;
    return current;
  };
  var resolveFieldValue = function(fieldReference: any): any {
    var key = String(fieldReference === undefined || fieldReference === null ? '' : fieldReference).trim().toLowerCase();
    var resolvedField = fieldsByKey[key];
    return resolvedField ? values[resolvedField.id] : undefined;
  };
  var parseExpression: () => any;
  var parseOr: () => any;
  var parseAnd: () => any;
  var parseComparison: () => any;
  var parseUnary: () => any;
  var parsePrimary: () => any;

  var evaluateFunction = function(name: string, args: any[]): any {
    var functionName = name.toLowerCase();
    if (functionName === 'field') { return args.length > 0 ? resolveFieldValue(args[0]) : undefined; }
    if (functionName === 'today') { var today = new Date(); today.setHours(0, 0, 0, 0); return today; }
    if (functionName === 'now') { return new Date(); }
    if (functionName === 'daysfromtoday') {
      var date = new Date(); date.setHours(0, 0, 0, 0);
      var delta = Number(args[0] || 0); date.setDate(date.getDate() + (isNaN(delta) ? 0 : delta)); return date;
    }
    if (functionName === 'isempty') { return isValueEmpty(args[0]); }
    if (functionName === 'hasvalue') { return !isValueEmpty(args[0]); }
    if (functionName === 'between') {
      return args.length >= 3 && compareValues(args[0], args[1]) >= 0 && compareValues(args[0], args[2]) <= 0;
    }
    throw new Error('Unsupported function: ' + name);
  };

  parseExpression = function(): any { return parseOr(); };
  parseOr = function(): any {
    var left = parseAnd();
    while (peek().type === 'operator' && peek().value === '||') {
      consume('operator', '||');
      var right = parseAnd();
      left = !!left || !!right;
    }
    return left;
  };
  parseAnd = function(): any {
    var left = parseComparison();
    while (peek().type === 'operator' && peek().value === '&&') {
      consume('operator', '&&');
      var right = parseComparison();
      left = !!left && !!right;
    }
    return left;
  };
  parseComparison = function(): any {
    var left = parseUnary();
    if (peek().type === 'operator' && ['==', '!=', '>', '>=', '<', '<='].indexOf(peek().value) >= 0) {
      var operator = consume('operator').value;
      var comparison = compareValues(left, parseUnary());
      if (operator === '==') { return comparison === 0; }
      if (operator === '!=') { return comparison !== 0; }
      if (operator === '>') { return comparison > 0; }
      if (operator === '>=') { return comparison >= 0; }
      if (operator === '<') { return comparison < 0; }
      return comparison <= 0;
    }
    return left;
  };
  parseUnary = function(): any {
    if (peek().type === 'operator' && peek().value === '!') { consume('operator', '!'); return !parseUnary(); }
    return parsePrimary();
  };
  parsePrimary = function(): any {
    var token = peek();
    if (token.type === 'paren' && token.value === '(') {
      consume('paren', '('); var parenthesized = parseExpression(); consume('paren', ')'); return parenthesized;
    }
    if (token.type === 'number') { consume('number'); return Number(token.value); }
    if (token.type === 'string') { consume('string'); return token.value; }
    if (token.type === 'identifier') {
      consume('identifier');
      var identifier = token.value;
      var normalizedIdentifier = identifier.toLowerCase();
      if (normalizedIdentifier === 'true') { return true; }
      if (normalizedIdentifier === 'false') { return false; }
      if (normalizedIdentifier === 'null') { return null; }
      if (peek().type === 'paren' && peek().value === '(') {
        consume('paren', '(');
        var args: any[] = [];
        if (!(peek().type === 'paren' && peek().value === ')')) {
          while (true) {
            args.push(parseExpression());
            if (peek().type === 'comma') { consume('comma', ','); continue; }
            break;
          }
        }
        consume('paren', ')');
        return evaluateFunction(identifier, args);
      }
      return resolveFieldValue(identifier);
    }
    throw new Error('Unexpected token in expression: ' + token.type + ' ' + token.value);
  };

  var result = parseExpression();
  if (peek().type !== 'eof') { throw new Error('Unexpected trailing tokens in expression.'); }
  return result;
}