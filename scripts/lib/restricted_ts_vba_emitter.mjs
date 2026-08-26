import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const ts = require("typescript");

const VBA_RESERVED_NAMES = new Set(
  `addressof alias and as attribute base binary boolean byref byte byval call case class close compare const currency date decimal declare defbool defbyte defcur defdate defdbl defdec defint deflng defobj defsng defstr defvar dim do double each else elseif empty end enum eqv erase error event exit explicit false filecopy for friend function get global gosub goto if imp implements in input integer is len let lib like line load lock long loop lset me mid mod name new next not nothing null object on open option optional or output paramarray preserve print private property public put random randomize redim rem reset resume return rset seek select set single static stop string sub then true type typeof unload until variant wend while with withevents write xor`
    .split(/\s+/)
    .filter(Boolean),
);

const SAFE_BINARY_OPERATORS = new Map([
  [ts.SyntaxKind.PlusToken, "+"],
  [ts.SyntaxKind.MinusToken, "-"],
  [ts.SyntaxKind.AsteriskToken, "*"],
  [ts.SyntaxKind.SlashToken, "/"],
  [ts.SyntaxKind.PercentToken, "Mod"],
  [ts.SyntaxKind.EqualsEqualsToken, "="],
  [ts.SyntaxKind.EqualsEqualsEqualsToken, "="],
  [ts.SyntaxKind.ExclamationEqualsToken, "<>"],
  [ts.SyntaxKind.ExclamationEqualsEqualsToken, "<>"],
  [ts.SyntaxKind.LessThanToken, "<"],
  [ts.SyntaxKind.LessThanEqualsToken, "<="],
  [ts.SyntaxKind.GreaterThanToken, ">"],
  [ts.SyntaxKind.GreaterThanEqualsToken, ">="],
  [ts.SyntaxKind.AmpersandAmpersandToken, "And"],
  [ts.SyntaxKind.BarBarToken, "Or"],
]);

const ASSIGNMENT_OPERATORS = new Map([
  [ts.SyntaxKind.EqualsToken, "="],
  [ts.SyntaxKind.PlusEqualsToken, "+"],
  [ts.SyntaxKind.MinusEqualsToken, "-"],
  [ts.SyntaxKind.AsteriskEqualsToken, "*"],
  [ts.SyntaxKind.SlashEqualsToken, "/"],
]);

const EMITTER_RESERVED_IDENTIFIERS = new Set(["console", "math"]);
const RUNTIME_FUNCTION_SUFFIXES = new Set(["ceil", "assertzerobasedarray"]);

function positionOf(sourceFile, node) {
  const position = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile, false));
  return { line: position.line + 1, column: position.character + 1 };
}

function supportedType(typeNode) {
  if (!typeNode) return null;
  if (typeNode.kind === ts.SyntaxKind.NumberKeyword) return { kind: "number", array: false };
  if (typeNode.kind === ts.SyntaxKind.BooleanKeyword) return { kind: "boolean", array: false };
  if (typeNode.kind === ts.SyntaxKind.StringKeyword) return { kind: "string", array: false };
  if (typeNode.kind === ts.SyntaxKind.VoidKeyword) return { kind: "void", array: false };
  if (!ts.isArrayTypeNode(typeNode)) return null;
  const element = supportedType(typeNode.elementType);
  if (!element || element.array || element.kind === "void") return null;
  return { kind: element.kind, array: true };
}

function isExported(node) {
  return Boolean(node.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword));
}

function isAsync(node) {
  return Boolean(node.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.AsyncKeyword));
}

function vbaType(type) {
  if (type.array) return "Variant";
  if (type.kind === "number") return "Double";
  if (type.kind === "boolean") return "Boolean";
  if (type.kind === "string") return "String";
  throw new Error(`No VBA type mapping for ${JSON.stringify(type)}`);
}

function quoteVbaString(value) {
  return `"${value.replaceAll('"', '""')}"`;
}

function createVirtualProgram(sourceText, filePath, compilerOptions) {
  const resolvedFilePath = path.resolve(filePath);
  const host = ts.createCompilerHost(compilerOptions, true);
  const originalGetSourceFile = host.getSourceFile.bind(host);
  const originalReadFile = host.readFile.bind(host);
  const originalFileExists = host.fileExists.bind(host);

  host.fileExists = (candidate) =>
    path.resolve(candidate) === resolvedFilePath || originalFileExists(candidate);
  host.readFile = (candidate) =>
    path.resolve(candidate) === resolvedFilePath ? sourceText : originalReadFile(candidate);
  host.getSourceFile = (candidate, languageVersion, onError, shouldCreateNewSourceFile) => {
    if (path.resolve(candidate) === resolvedFilePath) {
      return ts.createSourceFile(resolvedFilePath, sourceText, languageVersion, true, ts.ScriptKind.TS);
    }
    return originalGetSourceFile(candidate, languageVersion, onError, shouldCreateNewSourceFile);
  };

  const program = ts.createProgram([resolvedFilePath], compilerOptions, host);
  return { program, sourceFile: program.getSourceFile(resolvedFilePath) };
}

function createCompilerContext(sourceText, filePath) {
  const compilerOptions = {
    target: ts.ScriptTarget.ES2022,
    module: ts.ModuleKind.ESNext,
    strict: true,
    noEmit: true,
    noImplicitReturns: true,
    skipLibCheck: true,
  };
  const { program, sourceFile } = createVirtualProgram(sourceText, filePath, compilerOptions);
  if (!sourceFile) throw new Error(`TypeScript source was not loaded: ${filePath}`);
  return { program, sourceFile, checker: program.getTypeChecker() };
}

function flattenTypeScriptDiagnostics(program, sourceFile, filePath) {
  return ts
    .getPreEmitDiagnostics(program, sourceFile)
    .map((diagnostic) => {
      const start = diagnostic.start ?? 0;
      const position = sourceFile.getLineAndCharacterOfPosition(start);
      return {
        filePath,
        line: position.line + 1,
        column: position.character + 1,
        code: `TS${diagnostic.code}`,
        message: ts.flattenDiagnosticMessageText(diagnostic.messageText, "\n"),
      };
    });
}

export function formatRestrictedEmitterDiagnostic(diagnostic) {
  return `${diagnostic.filePath}:${diagnostic.line}:${diagnostic.column} [${diagnostic.code}] ${diagnostic.message}`;
}

export function analyzeRestrictedTypeScriptForVba(sourceText, filePath = "engine.ts", options = {}) {
  const prefix = options.prefix ?? "QPT_";
  const { program, sourceFile, checker } = createCompilerContext(sourceText, filePath);
  const diagnostics = flattenTypeScriptDiagnostics(program, sourceFile, filePath);
  const functions = new Map();
  const functionOrder = [];

  const add = (node, code, message) => {
    const position = positionOf(sourceFile, node);
    diagnostics.push({ filePath, ...position, code, message });
  };

  function validateName(node, registry, role) {
    if (!ts.isIdentifier(node)) {
      add(node, "IDENTIFIER_SHAPE", `${role} must be a simple identifier`);
      return null;
    }
    const name = node.text;
    const folded = name.toLowerCase();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) {
      add(node, "IDENTIFIER_ASCII", `${role} must be an ASCII identifier`);
    }
    if (VBA_RESERVED_NAMES.has(folded)) {
      add(node, "VBA_RESERVED_NAME", `${role} '${name}' is reserved by VBA`);
    }
    if (EMITTER_RESERVED_IDENTIFIERS.has(folded)) {
      add(node, "EMITTER_RESERVED_NAME", `${role} '${name}' is reserved by the restricted emitter`);
    }
    if (role === "function" && RUNTIME_FUNCTION_SUFFIXES.has(folded)) {
      add(node, "RUNTIME_NAME_COLLISION", `function '${name}' collides with a generated VBA runtime helper`);
    }
    if (`${prefix}${name}`.length > 200) {
      add(node, "IDENTIFIER_LENGTH", `${role} '${name}' is too long after applying the VBA prefix`);
    }
    if (registry.has(folded)) {
      add(node, "CASE_INSENSITIVE_COLLISION", `${role} '${name}' collides with '${registry.get(folded).name}' in VBA`);
      return null;
    }
    return { name, folded };
  }

  for (const statement of sourceFile.statements) {
    if (!ts.isFunctionDeclaration(statement)) continue;
    if (!statement.name) {
      add(statement, "FUNCTION_NAME", "function declaration requires a name");
      continue;
    }
    const checkedName = validateName(statement.name, functions, "function");
    if (!checkedName) continue;
    const returnType = supportedType(statement.type);
    if (!returnType || returnType.array) {
      add(statement.type ?? statement, "FUNCTION_RETURN_TYPE", "return type must be void, number, boolean, or string");
    }
    const symbol = checker.getSymbolAtLocation(statement.name);
    const metadata = {
      ...checkedName,
      node: statement,
      symbol,
      exported: isExported(statement),
      returnType,
      vbaName: `${prefix}${statement.name.text}`,
      parameters: [],
    };
    functions.set(checkedName.folded, metadata);
    functionOrder.push(metadata);
  }

  function validateFunction(metadata) {
    const node = metadata.node;
    if (node.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.DefaultKeyword)) {
      add(node, "DEFAULT_EXPORT", "default exports are forbidden");
    }
    if (isAsync(node) || node.asteriskToken || node.typeParameters?.length) {
      add(node, "FUNCTION_SHAPE", "async, generator, and generic functions are forbidden");
    }
    if (!node.body) {
      add(node, "FUNCTION_BODY", "function implementation is required");
      return;
    }

    const scope = metadata.scope;

    function resolveIdentifier(identifier) {
      const local = scope.get(identifier.text.toLowerCase());
      if (local) return local;
      add(identifier, "UNDECLARED_IDENTIFIER", `identifier '${identifier.text}' is not declared in the current function`);
      return null;
    }

    function validateExpression(expression, context = {}) {
      if (ts.isNumericLiteral(expression)) return;
      if (ts.isStringLiteral(expression)) {
        if (!/^[\x20-\x7E]*$/.test(expression.text)) {
          add(expression, "STRING_ASCII", "string literals must be ASCII for deterministic VBA source encoding");
        }
        return;
      }
      if (expression.kind === ts.SyntaxKind.TrueKeyword || expression.kind === ts.SyntaxKind.FalseKeyword) return;
      if (ts.isParenthesizedExpression(expression)) {
        validateExpression(expression.expression, context);
        return;
      }
      if (ts.isIdentifier(expression)) {
        resolveIdentifier(expression);
        return;
      }
      if (ts.isPrefixUnaryExpression(expression)) {
        if (![ts.SyntaxKind.ExclamationToken, ts.SyntaxKind.MinusToken, ts.SyntaxKind.PlusToken].includes(expression.operator)) {
          add(expression, "UNARY_OPERATOR", `unary operator '${expression.getText(sourceFile)}' is not supported`);
        }
        validateExpression(expression.operand, context);
        return;
      }
      if (ts.isBinaryExpression(expression)) {
        if (!SAFE_BINARY_OPERATORS.has(expression.operatorToken.kind)) {
          add(expression.operatorToken, "BINARY_OPERATOR", `operator '${expression.operatorToken.getText(sourceFile)}' is not supported`);
        }
        if (
          [ts.SyntaxKind.AmpersandAmpersandToken, ts.SyntaxKind.BarBarToken].includes(expression.operatorToken.kind)
        ) {
          const leftType = checker.getTypeAtLocation(expression.left);
          const rightType = checker.getTypeAtLocation(expression.right);
          if (
            (leftType.flags & ts.TypeFlags.BooleanLike) === 0 ||
            (rightType.flags & ts.TypeFlags.BooleanLike) === 0
          ) {
            add(expression.operatorToken, "LOGICAL_TYPE", "&& and || require boolean operands in the restricted subset");
          }
          const containsNonScalarOperation = (node) => {
            if (ts.isCallExpression(node) || ts.isElementAccessExpression(node) || ts.isPropertyAccessExpression(node)) {
              return true;
            }
            return node.getChildren(sourceFile).some(containsNonScalarOperation);
          };
          if (containsNonScalarOperation(expression.left) || containsNonScalarOperation(expression.right)) {
            add(
              expression.operatorToken,
              "LOGICAL_SHORT_CIRCUIT",
              "&& and || operands must be scalar-only because VBA And/Or do not short-circuit",
            );
          }
        }
        validateExpression(expression.left, context);
        validateExpression(expression.right, context);
        return;
      }
      if (ts.isElementAccessExpression(expression)) {
        if (!ts.isIdentifier(expression.expression)) {
          add(expression, "MULTIDIMENSIONAL_ARRAY", "only direct one-dimensional array access is allowed");
        } else {
          const value = resolveIdentifier(expression.expression);
          if (value?.type && !value.type.array) add(expression.expression, "ARRAY_TYPE", `'${value.name}' is not an array`);
        }
        if (!expression.argumentExpression) add(expression, "ARRAY_INDEX", "array access requires an index");
        else validateExpression(expression.argumentExpression);
        return;
      }
      if (ts.isPropertyAccessExpression(expression)) {
        if (context.allowLength && expression.name.text === "length" && ts.isIdentifier(expression.expression)) {
          const value = resolveIdentifier(expression.expression);
          if (value?.type && !value.type.array) add(expression.expression, "ARRAY_TYPE", `'${value.name}' is not an array`);
          return;
        }
        add(expression, "PROPERTY_ACCESS", `property '${expression.getText(sourceFile)}' is not supported`);
        return;
      }
      if (ts.isCallExpression(expression)) {
        if (
          ts.isPropertyAccessExpression(expression.expression) &&
          ts.isIdentifier(expression.expression.expression) &&
          expression.expression.expression.text === "Math" &&
          expression.expression.name.text === "ceil"
        ) {
          if (expression.arguments.length !== 1) add(expression, "CALL_ARITY", "Math.ceil requires exactly one argument");
          for (const argument of expression.arguments) validateExpression(argument);
          return;
        }
        if (ts.isIdentifier(expression.expression)) {
          const target = functions.get(expression.expression.text.toLowerCase());
          const callSymbol = checker.getSymbolAtLocation(expression.expression);
          if (!target || (target.symbol && callSymbol && target.symbol !== callSymbol)) {
            add(expression.expression, "HELPER_CALL", `call target '${expression.expression.text}' is not a top-level PoC function`);
          } else if (expression.arguments.length !== target.parameters.length) {
            add(expression, "CALL_ARITY", `function '${target.name}' expects ${target.parameters.length} arguments`);
          }
          for (const argument of expression.arguments) validateExpression(argument);
          return;
        }
        add(expression, "METHOD_NOT_ALLOWED", `call '${expression.expression.getText(sourceFile)}' is not supported`);
        for (const argument of expression.arguments) validateExpression(argument);
        return;
      }
      add(expression, "EXPRESSION_SYNTAX", `expression '${ts.SyntaxKind[expression.kind]}' is not supported`);
    }

    function validateAssignmentTarget(target) {
      if (ts.isIdentifier(target)) {
        const value = resolveIdentifier(target);
        if (value?.constant) add(target, "CONST_WRITE", `const local '${value.name}' cannot be reassigned`);
        if (value?.type?.array) add(target, "ARRAY_REASSIGN", `array '${value.name}' cannot be rebound`);
        return;
      }
      if (ts.isElementAccessExpression(target)) {
        if (!ts.isIdentifier(target.expression)) {
          add(target, "MULTIDIMENSIONAL_ARRAY", "only direct one-dimensional array writes are allowed");
        } else {
          const value = resolveIdentifier(target.expression);
          if (value?.type && !value.type.array) add(target.expression, "ARRAY_TYPE", `'${value.name}' is not an array`);
          if (value?.constant) add(target, "CONST_WRITE", `const array '${value.name}' cannot be mutated`);
        }
        if (target.argumentExpression) validateExpression(target.argumentExpression);
        return;
      }
      add(target, "ASSIGNMENT_TARGET", "assignment target must be a declared scalar or one-dimensional array element");
    }

    function validateVariableStatement(statement) {
      const flags = statement.declarationList.flags;
      const constant = (flags & ts.NodeFlags.Const) !== 0;
      if (!constant && (flags & ts.NodeFlags.Let) === 0) {
        add(statement, "VAR_DECLARATION", "use only let or const");
      }
      for (const declaration of statement.declarationList.declarations) {
        const type = supportedType(declaration.type);
        if (!type || type.kind === "void" || type.array) {
          add(declaration.type ?? declaration, "LOCAL_TYPE", "local variables require an explicit scalar number, boolean, or string type");
        }
        const checked = validateName(declaration.name, scope, "local variable");
        if (checked) {
          const symbol = checker.getSymbolAtLocation(declaration.name);
          scope.set(checked.folded, { ...checked, node: declaration, symbol, role: "local", type, constant });
        }
        if (!declaration.initializer) add(declaration, "LOCAL_INITIALIZER", "local variables require an initializer");
        else validateExpression(declaration.initializer);
      }
    }

    function validateStatements(statements) {
      for (const statement of statements) validateStatement(statement);
    }

    function validateForStatement(statement) {
      const initializer = statement.initializer;
      if (!initializer || !ts.isVariableDeclarationList(initializer) || initializer.declarations.length !== 1) {
        add(statement, "FOR_INITIALIZER", "for loop must start with 'let index = 0'");
        return;
      }
      const declaration = initializer.declarations[0];
      if (
        (initializer.flags & ts.NodeFlags.Let) === 0 ||
        !ts.isIdentifier(declaration.name) ||
        !declaration.initializer ||
        !ts.isNumericLiteral(declaration.initializer) ||
        declaration.initializer.text !== "0"
      ) {
        add(initializer, "FOR_INITIALIZER", "for loop must start with exactly 'let index = 0'");
        return;
      }
      const checked = validateName(declaration.name, scope, "loop variable");
      if (checked) {
        const symbol = checker.getSymbolAtLocation(declaration.name);
        scope.set(checked.folded, {
          ...checked,
          node: declaration,
          symbol,
          role: "loop",
          type: { kind: "number", array: false },
          constant: false,
        });
      }
      if (
        !statement.condition ||
        !ts.isBinaryExpression(statement.condition) ||
        statement.condition.operatorToken.kind !== ts.SyntaxKind.LessThanToken ||
        !ts.isIdentifier(statement.condition.left) ||
        statement.condition.left.text.toLowerCase() !== declaration.name.text.toLowerCase() ||
        !ts.isPropertyAccessExpression(statement.condition.right) ||
        statement.condition.right.name.text !== "length"
      ) {
        add(statement.condition ?? statement, "FOR_CONDITION", "for condition must be 'index < array.length'");
      } else {
        validateExpression(statement.condition.right, { allowLength: true });
      }
      if (
        !statement.incrementor ||
        !ts.isPostfixUnaryExpression(statement.incrementor) ||
        statement.incrementor.operator !== ts.SyntaxKind.PlusPlusToken ||
        !ts.isIdentifier(statement.incrementor.operand) ||
        statement.incrementor.operand.text.toLowerCase() !== declaration.name.text.toLowerCase()
      ) {
        add(statement.incrementor ?? statement, "FOR_INCREMENT", "for loop must use index++");
      }
      if (!ts.isBlock(statement.statement)) add(statement.statement, "BLOCK_REQUIRED", "for body must be a block");
      else validateStatements(statement.statement.statements);
    }

    function validateCallStatement(call) {
      if (
        ts.isPropertyAccessExpression(call.expression) &&
        ts.isIdentifier(call.expression.expression) &&
        call.expression.expression.text === "console" &&
        call.expression.name.text === "log"
      ) {
        if (call.arguments.length !== 1) add(call, "CALL_ARITY", "console.log requires exactly one argument");
        for (const argument of call.arguments) validateExpression(argument);
        return;
      }
      validateExpression(call);
      if (ts.isIdentifier(call.expression)) {
        const target = functions.get(call.expression.text.toLowerCase());
        if (target?.returnType?.kind !== "void") {
          add(call, "DISCARDED_RETURN", `return value from '${target.name}' cannot be discarded`);
        }
      } else {
        add(call, "CALL_STATEMENT", "only console.log or a void helper may be used as a statement");
      }
    }

    function validateStatement(statement) {
      if (ts.isVariableStatement(statement)) {
        validateVariableStatement(statement);
        return;
      }
      if (ts.isIfStatement(statement)) {
        validateExpression(statement.expression);
        if (!ts.isBlock(statement.thenStatement)) add(statement.thenStatement, "BLOCK_REQUIRED", "if body must be a block");
        else validateStatements(statement.thenStatement.statements);
        if (statement.elseStatement) {
          if (ts.isIfStatement(statement.elseStatement)) validateStatement(statement.elseStatement);
          else if (ts.isBlock(statement.elseStatement)) validateStatements(statement.elseStatement.statements);
          else add(statement.elseStatement, "BLOCK_REQUIRED", "else body must be a block");
        }
        return;
      }
      if (ts.isForStatement(statement)) {
        validateForStatement(statement);
        return;
      }
      if (ts.isWhileStatement(statement)) {
        validateExpression(statement.expression);
        if (!ts.isBlock(statement.statement)) add(statement.statement, "BLOCK_REQUIRED", "while body must be a block");
        else validateStatements(statement.statement.statements);
        return;
      }
      if (ts.isDoStatement(statement)) {
        add(statement, "DO_LOOP", "do loops are outside the restricted subset; use while");
        return;
      }
      if (ts.isReturnStatement(statement)) {
        const returnType = metadata.returnType;
        if (!returnType) return;
        if (returnType.kind === "void" && statement.expression) {
          add(statement, "RETURN_VALUE", "void function cannot return a value");
        } else if (returnType.kind !== "void" && !statement.expression) {
          add(statement, "RETURN_VALUE", "non-void function must return a value");
        }
        if (statement.expression) validateExpression(statement.expression);
        return;
      }
      if (ts.isExpressionStatement(statement)) {
        const expression = statement.expression;
        if (ts.isCallExpression(expression)) {
          validateCallStatement(expression);
          return;
        }
        if (ts.isPostfixUnaryExpression(expression)) {
          if (![ts.SyntaxKind.PlusPlusToken, ts.SyntaxKind.MinusMinusToken].includes(expression.operator)) {
            add(expression, "UPDATE_OPERATOR", "only ++ and -- are supported");
          }
          validateAssignmentTarget(expression.operand);
          return;
        }
        if (ts.isBinaryExpression(expression) && ASSIGNMENT_OPERATORS.has(expression.operatorToken.kind)) {
          validateAssignmentTarget(expression.left);
          validateExpression(expression.right);
          return;
        }
        add(expression, "EXPRESSION_STATEMENT", "unsupported expression statement");
        return;
      }
      add(statement, "STATEMENT_SYNTAX", `statement '${ts.SyntaxKind[statement.kind]}' is not supported`);
    }

    validateStatements(node.body.statements);
  }

  for (const statement of sourceFile.statements) {
    if (ts.isFunctionDeclaration(statement)) continue;
    if (ts.isImportDeclaration(statement) || ts.isExportDeclaration(statement)) {
      add(statement, "IMPORT_EXPORT", "imports and re-exports are forbidden");
    } else if (ts.isVariableStatement(statement)) {
      add(statement, "TOP_LEVEL_STATE", "top-level state is forbidden");
    } else {
      add(statement, "TOP_LEVEL_SYNTAX", `top-level '${ts.SyntaxKind[statement.kind]}' is not supported`);
    }
  }
  for (const metadata of functionOrder) {
    const scope = new Map();
    metadata.scope = scope;
    for (const parameter of metadata.node.parameters) {
      if (parameter.dotDotDotToken || parameter.questionToken || parameter.initializer) {
        add(parameter, "PARAMETER_SHAPE", "rest, optional, and default parameters are forbidden");
      }
      const type = supportedType(parameter.type);
      if (!type || type.kind === "void") {
        add(parameter.type ?? parameter, "PARAMETER_TYPE", "parameter type must be number, boolean, string, or a one-dimensional array");
      }
      const checked = validateName(parameter.name, scope, "parameter");
      if (checked) {
        const symbol = checker.getSymbolAtLocation(parameter.name);
        const value = { ...checked, node: parameter, symbol, role: "parameter", type, constant: false };
        scope.set(checked.folded, value);
        metadata.parameters.push(value);
      }
    }
  }
  for (const metadata of functionOrder) validateFunction(metadata);

  const uniqueDiagnostics = [...new Map(
    diagnostics.map((diagnostic) => [
      `${diagnostic.line}:${diagnostic.column}:${diagnostic.code}:${diagnostic.message}`,
      diagnostic,
    ]),
  ).values()].sort((left, right) =>
    left.line - right.line || left.column - right.column || left.code.localeCompare(right.code),
  );

  return { sourceFile, checker, functions, functionOrder, diagnostics: uniqueDiagnostics, prefix };
}

function createEmitter(analysis, options) {
  const { checker, sourceFile, functions, prefix } = analysis;
  const lines = [];
  let indent = 0;
  let usesCeil = false;

  const write = (line = "") => lines.push(`${"    ".repeat(indent)}${line}`);

  function lookupLocal(metadata, identifier) {
    return metadata.scope.get(identifier.text.toLowerCase());
  }

  function emitExpression(expression, metadata) {
    if (ts.isNumericLiteral(expression)) return expression.text;
    if (ts.isStringLiteral(expression)) return quoteVbaString(expression.text);
    if (expression.kind === ts.SyntaxKind.TrueKeyword) return "True";
    if (expression.kind === ts.SyntaxKind.FalseKeyword) return "False";
    if (ts.isIdentifier(expression)) return lookupLocal(metadata, expression)?.name ?? expression.text;
    if (ts.isParenthesizedExpression(expression)) return `(${emitExpression(expression.expression, metadata)})`;
    if (ts.isPrefixUnaryExpression(expression)) {
      const operator = expression.operator === ts.SyntaxKind.ExclamationToken
        ? "Not "
        : expression.operator === ts.SyntaxKind.MinusToken ? "-" : "+";
      return `${operator}${emitExpression(expression.operand, metadata)}`;
    }
    if (ts.isBinaryExpression(expression)) {
      let operator = SAFE_BINARY_OPERATORS.get(expression.operatorToken.kind);
      if (
        expression.operatorToken.kind === ts.SyntaxKind.PlusToken &&
        (checker.getTypeAtLocation(expression).flags & ts.TypeFlags.StringLike) !== 0
      ) {
        operator = "&";
      }
      return `(${emitExpression(expression.left, metadata)} ${operator} ${emitExpression(expression.right, metadata)})`;
    }
    if (ts.isElementAccessExpression(expression)) {
      return `${emitExpression(expression.expression, metadata)}(${emitExpression(expression.argumentExpression, metadata)})`;
    }
    if (ts.isPropertyAccessExpression(expression) && expression.name.text === "length") {
      const owner = emitExpression(expression.expression, metadata);
      return `(UBound(${owner}) - LBound(${owner}) + 1)`;
    }
    if (ts.isCallExpression(expression)) {
      if (ts.isPropertyAccessExpression(expression.expression)) {
        usesCeil = true;
        return `${prefix}Ceil(${expression.arguments.map((argument) => emitExpression(argument, metadata)).join(", ")})`;
      }
      const target = functions.get(expression.expression.text.toLowerCase());
      return `${target.vbaName}(${expression.arguments.map((argument) => emitExpression(argument, metadata)).join(", ")})`;
    }
    throw new Error(`Emitter reached unsupported expression ${ts.SyntaxKind[expression.kind]}`);
  }

  function emitAssignment(left, operatorKind, right, metadata) {
    const target = emitExpression(left, metadata);
    const value = emitExpression(right, metadata);
    const operator = ASSIGNMENT_OPERATORS.get(operatorKind);
    if (operator === "=") write(`${target} = ${value}`);
    else write(`${target} = ${target} ${operator} ${value}`);
  }

  function emitStatements(statements, metadata) {
    for (const statement of statements) emitStatement(statement, metadata);
  }

  function emitIf(statement, metadata, asElseIf = false) {
    write(`${asElseIf ? "ElseIf" : "If"} ${emitExpression(statement.expression, metadata)} Then`);
    indent++;
    emitStatements(statement.thenStatement.statements, metadata);
    indent--;
    if (statement.elseStatement) {
      if (ts.isIfStatement(statement.elseStatement)) {
        emitIf(statement.elseStatement, metadata, true);
      } else {
        write("Else");
        indent++;
        emitStatements(statement.elseStatement.statements, metadata);
        indent--;
        write("End If");
      }
    } else {
      write("End If");
    }
  }

  function emitStatement(statement, metadata) {
    if (ts.isVariableStatement(statement)) {
      for (const declaration of statement.declarationList.declarations) {
        const local = metadata.scope.get(declaration.name.text.toLowerCase());
        write(`${local.name} = ${emitExpression(declaration.initializer, metadata)}`);
      }
      return;
    }
    if (ts.isIfStatement(statement)) {
      emitIf(statement, metadata);
      return;
    }
    if (ts.isForStatement(statement)) {
      const indexName = statement.initializer.declarations[0].name.text;
      const arrayName = statement.condition.right.expression.text;
      write(`For ${indexName} = 0 To UBound(${arrayName})`);
      indent++;
      emitStatements(statement.statement.statements, metadata);
      indent--;
      write(`Next ${indexName}`);
      return;
    }
    if (ts.isWhileStatement(statement)) {
      write(`Do While ${emitExpression(statement.expression, metadata)}`);
      indent++;
      emitStatements(statement.statement.statements, metadata);
      indent--;
      write("Loop");
      return;
    }
    if (ts.isReturnStatement(statement)) {
      if (metadata.returnType.kind === "void") {
        write("Exit Sub");
      } else {
        write(`${metadata.vbaName} = ${emitExpression(statement.expression, metadata)}`);
        write("Exit Function");
      }
      return;
    }
    if (ts.isExpressionStatement(statement)) {
      const expression = statement.expression;
      if (ts.isCallExpression(expression)) {
        if (ts.isPropertyAccessExpression(expression.expression)) {
          write(`Debug.Print ${emitExpression(expression.arguments[0], metadata)}`);
        } else {
          const target = functions.get(expression.expression.text.toLowerCase());
          const args = expression.arguments.map((argument) => emitExpression(argument, metadata)).join(", ");
          write(args ? `Call ${target.vbaName}(${args})` : `Call ${target.vbaName}`);
        }
        return;
      }
      if (ts.isPostfixUnaryExpression(expression)) {
        const target = emitExpression(expression.operand, metadata);
        const operator = expression.operator === ts.SyntaxKind.PlusPlusToken ? "+" : "-";
        write(`${target} = ${target} ${operator} 1`);
        return;
      }
      emitAssignment(expression.left, expression.operatorToken.kind, expression.right, metadata);
      return;
    }
    throw new Error(`Emitter reached unsupported statement ${ts.SyntaxKind[statement.kind]}`);
  }

  write("Option Explicit");
  write();
  write("' Generated deterministically by the restricted TypeScript-to-VBA PoC.");
  write("' Source paths and timestamps are intentionally omitted.");
  write();

  for (const metadata of analysis.functionOrder) {
    const visibility = metadata.exported ? "Public" : "Private";
    const isSub = metadata.returnType.kind === "void";
    const parameters = metadata.parameters.map((parameter) => {
      const passing = parameter.type.array ? "ByRef" : "ByVal";
      return `${passing} ${parameter.name} As ${vbaType(parameter.type)}`;
    }).join(", ");
    const signature = isSub
      ? `${visibility} Sub ${metadata.vbaName}(${parameters})`
      : `${visibility} Function ${metadata.vbaName}(${parameters}) As ${vbaType(metadata.returnType)}`;
    write(signature);
    indent++;
    for (const local of metadata.scope.values()) {
      if (local.role === "local") write(`Dim ${local.name} As ${vbaType(local.type)}`);
      if (local.role === "loop") write(`Dim ${local.name} As Long`);
    }
    for (const parameter of metadata.parameters.filter((candidate) => candidate.type.array)) {
      write(`${prefix}AssertZeroBasedArray ${parameter.name}, ${quoteVbaString(parameter.name)}`);
    }
    emitStatements(metadata.node.body.statements, metadata);
    indent--;
    write(isSub ? "End Sub" : "End Function");
    write();
  }

  const functionLines = [...lines];
  lines.length = 0;
  lines.push(...functionLines);
  if (usesCeil) {
    write(`Private Function ${prefix}Ceil(ByVal value As Double) As Double`);
    indent++;
    write(`${prefix}Ceil = -Int(-value)`);
    indent--;
    write("End Function");
    write();
  }
  write(`Private Sub ${prefix}AssertZeroBasedArray(ByRef values As Variant, ByVal parameterName As String)`);
  indent++;
  write("If LBound(values) <> 0 Then");
  indent++;
  write(`Err.Raise 5, ${quoteVbaString(options.moduleName)}, parameterName & ${quoteVbaString(" must be a zero-based array")}`);
  indent--;
  write("End If");
  indent--;
  write("End Sub");
  write();

  return `${lines.join("\r\n")}\r\n`;
}

export function emitRestrictedTypeScriptToVba(sourceText, options = {}) {
  const filePath = options.filePath ?? "engine.ts";
  const moduleName = options.moduleName ?? "RestrictedPlanningEngine";
  const prefix = options.prefix ?? "QPT_";
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(moduleName) || VBA_RESERVED_NAMES.has(moduleName.toLowerCase())) {
    throw new Error(`Invalid VBA module name '${moduleName}'`);
  }
  if (!/^[A-Za-z_][A-Za-z0-9_]*_$/.test(prefix)) {
    throw new Error(`VBA prefix must be ASCII and end with underscore: '${prefix}'`);
  }
  const analysis = analyzeRestrictedTypeScriptForVba(sourceText, filePath, { prefix });
  if (analysis.diagnostics.length) {
    const error = new Error(analysis.diagnostics.map(formatRestrictedEmitterDiagnostic).join("\n"));
    error.name = "RestrictedTypeScriptVbaEmitterError";
    error.diagnostics = analysis.diagnostics;
    throw error;
  }
  return {
    vba: createEmitter(analysis, { moduleName }),
    manifest: {
      schemaVersion: 1,
      emitter: "restricted-typescript-vba-poc",
      typescriptVersion: ts.version,
      moduleName,
      prefix,
      functions: analysis.functionOrder.map((metadata) => ({
        sourceName: metadata.name,
        vbaName: metadata.vbaName,
        visibility: metadata.exported ? "public" : "private",
        returnType: metadata.returnType.kind,
        parameters: metadata.parameters.map((parameter) => ({
          name: parameter.name,
          type: parameter.type.array ? `${parameter.type.kind}[]` : parameter.type.kind,
          passing: parameter.type.array ? "ByRef" : "ByVal",
        })),
      })),
    },
  };
}

export function emitRestrictedTypeScriptFileToVba(filePath, options = {}) {
  const sourceText = fs.readFileSync(filePath, "utf8");
  return emitRestrictedTypeScriptToVba(sourceText, { ...options, filePath });
}
