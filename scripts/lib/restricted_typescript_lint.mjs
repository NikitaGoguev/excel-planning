import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const ts = require("typescript");

const VBA_RESERVED_NAMES = new Set(
  `addressof alias and as attribute base binary boolean byref byte byval call case class close compare const currency date decimal declare defbool defbyte defcur defdate defdbl defdec defint deflng defobj defsng defstr defvar dim do double each else elseif empty end enum eqv erase error event exit explicit false filecopy for friend function get global gosub goto if imp implements in input integer is len let lib like line load lock long loop lset me mid mod name new next not nothing null object on open option optional or output paramarray preserve print private property public put random randomize redim rem reset resume return rset seek select set single static stop string sub then true type typeof unload until variant wend while with withevents write xor`
    .split(/\s+/)
    .filter(Boolean),
);

const FORBIDDEN_GLOBALS = new Set([
  "date",
  "map",
  "math",
  "promise",
  "set",
  "weakmap",
  "weakset",
]);

const ALLOWED_BINARY_OPERATORS = new Set([
  ts.SyntaxKind.PlusToken,
  ts.SyntaxKind.MinusToken,
  ts.SyntaxKind.AsteriskToken,
  ts.SyntaxKind.SlashToken,
  ts.SyntaxKind.EqualsEqualsToken,
  ts.SyntaxKind.EqualsEqualsEqualsToken,
  ts.SyntaxKind.ExclamationEqualsToken,
  ts.SyntaxKind.ExclamationEqualsEqualsToken,
  ts.SyntaxKind.LessThanToken,
  ts.SyntaxKind.LessThanEqualsToken,
  ts.SyntaxKind.GreaterThanToken,
  ts.SyntaxKind.GreaterThanEqualsToken,
  ts.SyntaxKind.AmpersandAmpersandToken,
  ts.SyntaxKind.BarBarToken,
]);

const ASSIGNMENT_OPERATORS = new Set([
  ts.SyntaxKind.EqualsToken,
  ts.SyntaxKind.PlusEqualsToken,
  ts.SyntaxKind.MinusEqualsToken,
]);

function lineAndColumn(sourceFile, node) {
  const start = node.getStart(sourceFile, false);
  const position = sourceFile.getLineAndCharacterOfPosition(start);
  return { line: position.line + 1, column: position.character + 1 };
}

function isExported(node) {
  return Boolean(node.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword));
}

function isAsync(node) {
  return Boolean(node.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.AsyncKeyword));
}

function isSimpleIdentifier(node) {
  return ts.isIdentifier(node);
}

function supportedTypeKind(typeNode) {
  if (!typeNode) return null;
  if (typeNode.kind === ts.SyntaxKind.NumberKeyword) return { kind: "number", array: false };
  if (typeNode.kind === ts.SyntaxKind.BooleanKeyword) return { kind: "boolean", array: false };
  if (typeNode.kind === ts.SyntaxKind.StringKeyword) return { kind: "string", array: false };
  if (typeNode.kind === ts.SyntaxKind.VoidKeyword) return { kind: "void", array: false };
  if (!ts.isArrayTypeNode(typeNode)) return null;
  const element = supportedTypeKind(typeNode.elementType);
  if (!element || element.array || element.kind === "void") return null;
  return { kind: element.kind, array: true };
}

function rootIdentifier(node) {
  let current = node;
  while (ts.isElementAccessExpression(current)) current = current.expression;
  return ts.isIdentifier(current) ? current : null;
}

function diagnosticKey(diagnostic) {
  return `${diagnostic.filePath}:${diagnostic.line}:${diagnostic.column}:${diagnostic.code}:${diagnostic.message}`;
}

export function formatRestrictedTypeScriptDiagnostic(diagnostic) {
  return `${diagnostic.filePath}:${diagnostic.line}:${diagnostic.column} [${diagnostic.code}] ${diagnostic.message}`;
}

export function lintRestrictedTypeScript(sourceText, filePath = "<memory>") {
  const sourceFile = ts.createSourceFile(filePath, sourceText, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
  const diagnostics = [];
  const add = (node, code, message) => {
    const position = lineAndColumn(sourceFile, node);
    diagnostics.push({ filePath, ...position, code, message });
  };

  for (const parseDiagnostic of sourceFile.parseDiagnostics) {
    const position = sourceFile.getLineAndCharacterOfPosition(parseDiagnostic.start ?? 0);
    diagnostics.push({
      filePath,
      line: position.line + 1,
      column: position.character + 1,
      code: "TS_PARSE",
      message: ts.flattenDiagnosticMessageText(parseDiagnostic.messageText, "\n"),
    });
  }

  const functionNames = new Map();

  function validateIdentifier(node, scope, role) {
    if (!isSimpleIdentifier(node)) {
      add(node, "IDENTIFIER_SHAPE", `${role} must use a simple identifier`);
      return null;
    }
    const name = node.text;
    const folded = name.toLowerCase();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
      add(node, "IDENTIFIER_ASCII", `${role} must be an ASCII VBA-safe identifier`);
    }
    if (VBA_RESERVED_NAMES.has(folded)) {
      add(node, "VBA_RESERVED_NAME", `${role} '${name}' is reserved by VBA`);
    }
    if (folded.startsWith("ts_")) {
      add(node, "RUNTIME_NAME_COLLISION", `${role} '${name}' collides with the generated TS_ runtime namespace`);
    }
    if (scope?.has(folded)) {
      add(node, "CASE_INSENSITIVE_COLLISION", `${role} '${name}' collides with '${scope.get(folded).name}' in VBA`);
      return null;
    }
    return { name, folded };
  }

  function declareIdentifier(node, scope, role, metadata) {
    const checked = validateIdentifier(node, scope, role);
    if (!checked) return;
    scope.set(checked.folded, { name: checked.name, ...metadata });
  }

  function resolveIdentifier(node, scope) {
    const folded = node.text.toLowerCase();
    const symbol = scope.get(folded);
    if (symbol) return symbol;
    if (FORBIDDEN_GLOBALS.has(folded)) {
      add(node, "FORBIDDEN_GLOBAL", `global '${node.text}' is not available in the restricted subset`);
    } else {
      add(node, "UNDECLARED_IDENTIFIER", `identifier '${node.text}' is not declared in the current function`);
    }
    return null;
  }

  function validateExpression(node, scope, context = {}) {
    if (
      ts.isNumericLiteral(node) ||
      ts.isStringLiteral(node) ||
      node.kind === ts.SyntaxKind.TrueKeyword ||
      node.kind === ts.SyntaxKind.FalseKeyword
    ) {
      return;
    }
    if (ts.isParenthesizedExpression(node)) {
      validateExpression(node.expression, scope, context);
      return;
    }
    if (ts.isIdentifier(node)) {
      resolveIdentifier(node, scope);
      return;
    }
    if (ts.isPrefixUnaryExpression(node) && node.operator === ts.SyntaxKind.ExclamationToken) {
      validateExpression(node.operand, scope, context);
      return;
    }
    if (ts.isBinaryExpression(node)) {
      if (!ALLOWED_BINARY_OPERATORS.has(node.operatorToken.kind)) {
        add(node.operatorToken, "BINARY_OPERATOR", `operator '${node.operatorToken.getText(sourceFile)}' is not supported`);
      }
      validateExpression(node.left, scope, context);
      validateExpression(node.right, scope, context);
      return;
    }
    if (ts.isElementAccessExpression(node)) {
      const root = rootIdentifier(node);
      if (!root || !ts.isIdentifier(node.expression)) {
        add(node, "MULTIDIMENSIONAL_ARRAY", "only one-dimensional array access is allowed");
      } else {
        const symbol = resolveIdentifier(root, scope);
        if (symbol && !symbol.type?.array) {
          add(root, "ARRAY_TYPE", `'${root.text}' is not declared as a one-dimensional array`);
        }
      }
      if (!node.argumentExpression) {
        add(node, "ARRAY_INDEX", "array access requires an index");
      } else {
        validateExpression(node.argumentExpression, scope);
      }
      return;
    }
    if (ts.isPropertyAccessExpression(node)) {
      if (
        context.allowArrayLength &&
        node.name.text === "length" &&
        ts.isIdentifier(node.expression)
      ) {
        const symbol = resolveIdentifier(node.expression, scope);
        if (symbol && !symbol.type?.array) {
          add(node.expression, "ARRAY_TYPE", `'${node.expression.text}' is not an array`);
        }
        return;
      }
      add(node, "PROPERTY_ACCESS", `property '${node.name.text}' is not allowed here`);
      return;
    }
    if (ts.isCallExpression(node)) {
      if (ts.isPropertyAccessExpression(node.expression)) {
        const owner = node.expression.expression.getText(sourceFile);
        if (ts.isIdentifier(node.expression.expression) && FORBIDDEN_GLOBALS.has(owner.toLowerCase())) {
          add(node.expression.expression, "FORBIDDEN_GLOBAL", `global '${owner}' is not available in the restricted subset`);
        }
        add(node, "METHOD_NOT_ALLOWED", `method '${owner}.${node.expression.name.text}' is not in the allowlist`);
      } else {
        add(node, "HELPER_CALL", "helper and global function calls are not supported by the emitter");
      }
      for (const argument of node.arguments) validateExpression(argument, scope);
      return;
    }
    if (ts.isArrayLiteralExpression(node)) {
      add(node, "ARRAY_LITERAL", "array literals are not part of the restricted engine subset");
      return;
    }
    if (ts.isObjectLiteralExpression(node)) {
      add(node, "OBJECT_LITERAL", "object literals are not part of the restricted engine subset");
      return;
    }
    add(node, "EXPRESSION_SYNTAX", `expression '${ts.SyntaxKind[node.kind]}' is not supported`);
  }

  function validateAssignmentTarget(node, scope) {
    if (ts.isIdentifier(node)) {
      const symbol = resolveIdentifier(node, scope);
      if (symbol?.role === "parameter") {
        add(node, "PARAMETER_WRITE", `parameter '${node.text}' is read-only because the emitter uses ByVal Variant`);
      }
      return;
    }
    if (ts.isElementAccessExpression(node)) {
      const root = rootIdentifier(node);
      const symbol = root ? resolveIdentifier(root, scope) : null;
      if (symbol?.role === "parameter") {
        add(node, "PARAMETER_WRITE", `array parameter '${root.text}' cannot be used as an output channel`);
      } else {
        add(node, "ARRAY_WRITE", "array mutation is not allowed in the restricted subset");
      }
      if (node.argumentExpression) validateExpression(node.argumentExpression, scope);
      return;
    }
    add(node, "ASSIGNMENT_TARGET", "assignment target must be a local scalar variable");
  }

  function validateConsoleLog(node, scope) {
    if (
      !ts.isPropertyAccessExpression(node.expression) ||
      !ts.isIdentifier(node.expression.expression) ||
      node.expression.expression.text !== "console" ||
      node.expression.name.text !== "log"
    ) {
      if (ts.isPropertyAccessExpression(node.expression)) {
        const owner = node.expression.expression.getText(sourceFile);
        if (ts.isIdentifier(node.expression.expression) && FORBIDDEN_GLOBALS.has(owner.toLowerCase())) {
          add(node.expression.expression, "FORBIDDEN_GLOBAL", `global '${owner}' is not available in the restricted subset`);
        }
        add(node, "METHOD_NOT_ALLOWED", `method '${node.expression.getText(sourceFile)}' is not in the allowlist`);
      } else {
        add(node, "HELPER_CALL", "helper and global function calls are not supported by the emitter");
      }
      return;
    }
    for (const argument of node.arguments) validateExpression(argument, scope);
  }

  function validateVariableStatement(node, scope) {
    const flags = node.declarationList.flags;
    const isConst = (flags & ts.NodeFlags.Const) !== 0;
    const isLet = (flags & ts.NodeFlags.Let) !== 0;
    if (!isConst && !isLet) add(node, "VAR_DECLARATION", "use only let or const; var is not allowed");

    for (const declaration of node.declarationList.declarations) {
      const type = declaration.type ? supportedTypeKind(declaration.type) : null;
      if (declaration.type && !type) {
        add(declaration.type, "VARIABLE_TYPE", "variable type must be number, boolean, string, or a one-dimensional array");
      }
      declareIdentifier(declaration.name, scope, "local variable", { role: "local", type });
      if (!declaration.initializer) {
        add(declaration, "VARIABLE_INITIALIZER", "local variables require an initializer");
      } else {
        validateExpression(declaration.initializer, scope);
      }
    }
  }

  function validateForStatement(node, scope) {
    if (!node.initializer || !ts.isVariableDeclarationList(node.initializer)) {
      add(node, "FOR_INITIALIZER", "for loop must start with 'let index = 0'");
      return;
    }
    const declarations = node.initializer.declarations;
    const declaration = declarations.length === 1 ? declarations[0] : null;
    if (
      !declaration ||
      !ts.isIdentifier(declaration.name) ||
      !declaration.initializer ||
      !ts.isNumericLiteral(declaration.initializer) ||
      declaration.initializer.text !== "0" ||
      (node.initializer.flags & ts.NodeFlags.Let) === 0
    ) {
      add(node.initializer, "FOR_INITIALIZER", "for loop must start with exactly 'let index = 0'");
      return;
    }
    declareIdentifier(declaration.name, scope, "loop variable", {
      role: "loop",
      type: { kind: "number", array: false },
    });

    if (
      !node.condition ||
      !ts.isBinaryExpression(node.condition) ||
      node.condition.operatorToken.kind !== ts.SyntaxKind.LessThanToken ||
      !ts.isIdentifier(node.condition.left) ||
      node.condition.left.text.toLowerCase() !== declaration.name.text.toLowerCase() ||
      !ts.isPropertyAccessExpression(node.condition.right) ||
      node.condition.right.name.text !== "length"
    ) {
      add(node.condition ?? node, "FOR_CONDITION", "for condition must be 'index < oneDimensionalArray.length'");
    } else {
      validateExpression(node.condition.right, scope, { allowArrayLength: true });
    }

    if (
      !node.incrementor ||
      !ts.isPostfixUnaryExpression(node.incrementor) ||
      node.incrementor.operator !== ts.SyntaxKind.PlusPlusToken ||
      !ts.isIdentifier(node.incrementor.operand) ||
      node.incrementor.operand.text.toLowerCase() !== declaration.name.text.toLowerCase()
    ) {
      add(node.incrementor ?? node, "FOR_INCREMENT", "for loop must increment its index with index++");
    }

    if (!ts.isBlock(node.statement)) {
      add(node.statement, "BLOCK_REQUIRED", "for body must be a block");
    } else {
      validateStatements(node.statement.statements, scope);
    }
  }

  function validateExpressionStatement(node, scope) {
    const expression = node.expression;
    if (ts.isCallExpression(expression)) {
      validateConsoleLog(expression, scope);
      return;
    }
    if (ts.isPostfixUnaryExpression(expression)) {
      if (
        expression.operator !== ts.SyntaxKind.PlusPlusToken &&
        expression.operator !== ts.SyntaxKind.MinusMinusToken
      ) {
        add(expression, "UPDATE_OPERATOR", "only ++ and -- updates are allowed");
      }
      validateAssignmentTarget(expression.operand, scope);
      return;
    }
    if (ts.isBinaryExpression(expression) && ASSIGNMENT_OPERATORS.has(expression.operatorToken.kind)) {
      validateAssignmentTarget(expression.left, scope);
      validateExpression(expression.right, scope);
      return;
    }
    add(expression, "EXPRESSION_STATEMENT", "statement is not an allowed assignment, update, or console.log call");
  }

  function validateStatement(node, scope) {
    if (ts.isVariableStatement(node)) {
      validateVariableStatement(node, scope);
      return;
    }
    if (ts.isIfStatement(node)) {
      validateExpression(node.expression, scope);
      if (!ts.isBlock(node.thenStatement)) add(node.thenStatement, "BLOCK_REQUIRED", "if body must be a block");
      else validateStatements(node.thenStatement.statements, scope);
      if (node.elseStatement) {
        if (ts.isIfStatement(node.elseStatement)) validateStatement(node.elseStatement, scope);
        else if (ts.isBlock(node.elseStatement)) validateStatements(node.elseStatement.statements, scope);
        else add(node.elseStatement, "BLOCK_REQUIRED", "else body must be a block");
      }
      return;
    }
    if (ts.isForStatement(node)) {
      validateForStatement(node, scope);
      return;
    }
    if (ts.isExpressionStatement(node)) {
      validateExpressionStatement(node, scope);
      return;
    }
    if (ts.isReturnStatement(node)) {
      add(node, "UNSUPPORTED_RETURN", "typescript-to-vba 1.0.1 drops return statements");
      if (node.expression) validateExpression(node.expression, scope);
      return;
    }
    if (ts.isWhileStatement(node) || ts.isDoStatement(node)) {
      add(node, "UNSUPPORTED_LOOP", "typescript-to-vba 1.0.1 does not emit while/do loop control flow");
      return;
    }
    if (ts.isBlock(node)) {
      validateStatements(node.statements, scope);
      return;
    }
    add(node, "STATEMENT_SYNTAX", `statement '${ts.SyntaxKind[node.kind]}' is not supported`);
  }

  function validateStatements(statements, scope) {
    for (const statement of statements) validateStatement(statement, scope);
  }

  function validateFunction(node) {
    if (!node.name) {
      add(node, "FUNCTION_NAME", "exported function requires a name");
      return;
    }
    if (!isExported(node)) add(node, "FUNCTION_EXPORT", "top-level function must be exported");
    if (node.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.DefaultKeyword)) {
      add(node, "FUNCTION_DEFAULT", "default exports are forbidden; use a named export function");
    }
    if (isAsync(node) || node.asteriskToken) add(node, "FUNCTION_ASYNC", "async and generator functions are forbidden");
    if (node.typeParameters?.length) add(node, "FUNCTION_GENERIC", "generic functions are forbidden");

    const checkedName = validateIdentifier(node.name, functionNames, "function");
    if (checkedName) functionNames.set(checkedName.folded, { name: checkedName.name });

    const returnType = supportedTypeKind(node.type);
    if (!returnType || returnType.kind !== "void" || returnType.array) {
      add(node.type ?? node, "FUNCTION_RETURN", "only an explicit void return type is allowed");
    }
    if (!node.body) {
      add(node, "FUNCTION_BODY", "function implementation is required");
      return;
    }

    const scope = new Map();
    for (const parameter of node.parameters) {
      if (parameter.dotDotDotToken || parameter.questionToken || parameter.initializer) {
        add(parameter, "PARAMETER_SHAPE", "rest, optional, and default parameters are forbidden");
      }
      const type = supportedTypeKind(parameter.type);
      if (!type || type.kind === "void") {
        add(parameter.type ?? parameter, "PARAMETER_TYPE", "parameter type must be number, boolean, string, or a one-dimensional array");
      }
      declareIdentifier(parameter.name, scope, "parameter", { role: "parameter", type });
    }
    validateStatements(node.body.statements, scope);
  }

  for (const statement of sourceFile.statements) {
    if (ts.isFunctionDeclaration(statement)) {
      validateFunction(statement);
    } else if (ts.isVariableStatement(statement)) {
      add(statement, "TOP_LEVEL_STATE", "top-level mutable or constant state is forbidden in the PoC engine subset");
    } else if (ts.isImportDeclaration(statement) || ts.isExportDeclaration(statement)) {
      add(statement, "IMPORT_EXPORT", "imports and re-exports are not emitted safely by typescript-to-vba 1.0.1");
    } else if (ts.isClassDeclaration(statement) || ts.isInterfaceDeclaration(statement) || ts.isTypeAliasDeclaration(statement)) {
      add(statement, "TOP_LEVEL_TYPE", "classes, interfaces, and type aliases are outside the restricted subset");
    } else if (statement.kind !== ts.SyntaxKind.EndOfFileToken) {
      add(statement, "TOP_LEVEL_SYNTAX", `top-level '${ts.SyntaxKind[statement.kind]}' is not supported`);
    }
  }

  return [...new Map(diagnostics.map((diagnostic) => [diagnosticKey(diagnostic), diagnostic])).values()].sort(
    (left, right) => left.line - right.line || left.column - right.column || left.code.localeCompare(right.code),
  );
}

export function lintRestrictedTypeScriptOrThrow(sourceText, filePath = "<memory>") {
  const diagnostics = lintRestrictedTypeScript(sourceText, filePath);
  if (diagnostics.length) {
    const error = new Error(diagnostics.map(formatRestrictedTypeScriptDiagnostic).join("\n"));
    error.name = "RestrictedTypeScriptLintError";
    error.diagnostics = diagnostics;
    throw error;
  }
}

export function relativeDiagnosticPath(filePath, repoRoot) {
  const relative = path.relative(repoRoot, filePath);
  return relative && !relative.startsWith("..") ? relative.replaceAll(path.sep, "/") : filePath;
}
