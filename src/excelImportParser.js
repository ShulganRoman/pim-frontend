import * as XLSX from "xlsx";
import {
  ATTRIBUTE_HEADERS,
  CONFIG_HEADERS,
  GROUP_HEADERS,
  ITEM_BASE_HEADERS,
  PRODUCT_SHEETS,
  SHEETS,
  TYPE_GROUP_BINDING_HEADERS,
  TYPE_HEADERS,
} from "./excelImportTemplate.js";

const VALID_IMPORT_MODES = new Set([
  "CREATE_ONLY",
  "UPDATE_ONLY",
  "CREATE_UPDATE",
]);

const VALID_ERROR_PROCESSING = new Set(["PROCESS_WARN", "WARN_REJECTED"]);

const ATTRIBUTE_TYPE_CODES = new Set([1, 2, 3, 4, 5, 6, 7, 8]);

function toText(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}

function normalizeIdentifier(value) {
  return toText(value).toLowerCase();
}

function isBlank(value) {
  if (value === null || value === undefined) {
    return true;
  }
  if (typeof value === "string") {
    return value.trim() === "";
  }
  return false;
}

function parseSheetWithHeaders(workbook, sheetName, requiredHeaders) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    return {
      exists: false,
      headers: [],
      rows: [],
      missingHeaders: requiredHeaders,
    };
  }

  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: false,
  });

  const rawHeaders = Array.isArray(matrix[0]) ? matrix[0] : [];
  const headers = rawHeaders.map((value) => toText(value));
  const headerMap = new Map();
  headers.forEach((header, index) => {
    if (header) {
      headerMap.set(header, index);
    }
  });

  const missingHeaders = requiredHeaders.filter((header) => !headerMap.has(header));

  const rows = [];
  for (let rowIndex = 1; rowIndex < matrix.length; rowIndex += 1) {
    const rowArray = Array.isArray(matrix[rowIndex]) ? matrix[rowIndex] : [];
    const rowObject = {};

    headers.forEach((header, index) => {
      rowObject[header] = rowArray[index];
    });

    rows.push({ rowNumber: rowIndex + 1, data: rowObject });
  }

  return {
    exists: true,
    headers,
    rows,
    missingHeaders,
  };
}

function addIssue(target, severity, sheet, row, field, message) {
  target.push({ severity, sheet, row, field, message });
}

function parseBoolean(value) {
  if (typeof value === "boolean") {
    return { ok: true, value };
  }

  if (typeof value === "number") {
    if (value === 1) return { ok: true, value: true };
    if (value === 0) return { ok: true, value: false };
    return { ok: false, error: "Expected boolean-like value" };
  }

  const text = toText(value).toLowerCase();
  if (text === "") return { ok: false, error: "Expected boolean-like value" };

  if (["true", "1", "yes", "y"].includes(text)) return { ok: true, value: true };
  if (["false", "0", "no", "n"].includes(text)) return { ok: true, value: false };

  return { ok: false, error: "Expected one of: true/false/1/0/yes/no" };
}

function parseInteger(value) {
  if (typeof value === "number" && Number.isInteger(value)) {
    return { ok: true, value };
  }

  const text = toText(value);
  if (!/^-?\d+$/.test(text)) {
    return { ok: false, error: "Expected integer" };
  }

  return { ok: true, value: Number.parseInt(text, 10) };
}

function parseFloatNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return { ok: true, value };
  }

  const text = toText(value);
  if (text === "") {
    return { ok: false, error: "Expected number" };
  }

  const parsed = Number.parseFloat(text);
  if (!Number.isFinite(parsed)) {
    return { ok: false, error: "Expected number" };
  }

  return { ok: true, value: parsed };
}

function parseCsvList(value) {
  if (Array.isArray(value)) {
    return value
      .map((entry) => normalizeIdentifier(entry))
      .filter((entry) => entry.length > 0);
  }

  const text = toText(value);
  if (!text) {
    return [];
  }

  return text
    .split(/[;,]/)
    .map((entry) => normalizeIdentifier(entry))
    .filter((entry) => entry.length > 0);
}

function parseJsonObject(value, allowBlank = true) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    return { ok: true, value };
  }

  const text = toText(value);
  if (text === "") {
    return allowBlank ? { ok: true, value: {} } : { ok: false, error: "JSON object is required" };
  }

  try {
    const parsed = JSON.parse(text);
    if (!parsed || Array.isArray(parsed) || typeof parsed !== "object") {
      return { ok: false, error: "Expected JSON object" };
    }
    return { ok: true, value: parsed };
  } catch {
    return { ok: false, error: "Invalid JSON object" };
  }
}

function rowHasAnyValue(rowData, headers) {
  return headers.some((header) => !isBlank(rowData[header]));
}

function coerceAttributeValue(rawValue, attribute, defaultLanguage) {
  if (rawValue === null || rawValue === undefined || toText(rawValue) === "") {
    return { ok: true, skip: true };
  }

  if (attribute.languageDependent) {
    if (rawValue && typeof rawValue === "object" && !Array.isArray(rawValue)) {
      return { ok: true, value: rawValue };
    }

    return {
      ok: true,
      value: {
        [defaultLanguage]: toText(rawValue),
      },
    };
  }

  switch (attribute.type) {
    case 2: {
      const parsed = parseBoolean(rawValue);
      if (!parsed.ok) return { ok: false, error: parsed.error };
      return { ok: true, value: parsed.value };
    }
    case 3: {
      const parsed = parseInteger(rawValue);
      if (!parsed.ok) return { ok: false, error: parsed.error };
      return { ok: true, value: parsed.value };
    }
    case 4: {
      const parsed = parseFloatNumber(rawValue);
      if (!parsed.ok) return { ok: false, error: parsed.error };
      return { ok: true, value: parsed.value };
    }
    case 1:
    case 5:
    case 6:
    case 7:
    case 8:
    default:
      return { ok: true, value: toText(rawValue) };
  }
}

function ensureRequiredSheets(workbook, issues) {
  const requiredSheets = [
    SHEETS.CONFIG,
    SHEETS.GROUPS,
    SHEETS.ATTRIBUTES,
    SHEETS.TYPES,
    SHEETS.TYPE_GROUP_BINDINGS,
    SHEETS.ITEM_PARENTS,
    ...PRODUCT_SHEETS,
  ];

  const available = new Set(workbook.SheetNames || []);
  for (const sheet of requiredSheets) {
    if (!available.has(sheet)) {
      addIssue(issues, "error", sheet, null, "sheet", "Sheet is required but missing");
    }
  }
}

function parseConfig(workbook, issues) {
  const parsed = parseSheetWithHeaders(workbook, SHEETS.CONFIG, CONFIG_HEADERS);
  if (!parsed.exists) {
    return {
      mode: "CREATE_UPDATE",
      errors: "PROCESS_WARN",
      defaultLanguage: "en",
    };
  }

  if (parsed.missingHeaders.length > 0) {
    parsed.missingHeaders.forEach((header) => {
      addIssue(issues, "error", SHEETS.CONFIG, 1, header, "Missing required header");
    });
  }

  const kv = new Map();
  parsed.rows.forEach(({ rowNumber, data }) => {
    if (rowHasAnyValue(data, CONFIG_HEADERS)) {
      const key = toText(data.key).toLowerCase();
      const value = toText(data.value);

      if (!key) {
        addIssue(issues, "error", SHEETS.CONFIG, rowNumber, "key", "Key is required");
        return;
      }

      kv.set(key, value);
    }
  });

  const mode = (kv.get("mode") || "CREATE_UPDATE").toUpperCase();
  const errors = (kv.get("errors") || "PROCESS_WARN").toUpperCase();
  const defaultLanguage = (kv.get("default_language") || "en").toLowerCase();

  if (!VALID_IMPORT_MODES.has(mode)) {
    addIssue(
      issues,
      "error",
      SHEETS.CONFIG,
      null,
      "mode",
      "mode must be CREATE_ONLY, UPDATE_ONLY, or CREATE_UPDATE",
    );
  }

  if (!VALID_ERROR_PROCESSING.has(errors)) {
    addIssue(
      issues,
      "error",
      SHEETS.CONFIG,
      null,
      "errors",
      "errors must be PROCESS_WARN or WARN_REJECTED",
    );
  }

  if (!/^[a-z]{2}(-[a-z]{2})?$/.test(defaultLanguage)) {
    addIssue(
      issues,
      "warning",
      SHEETS.CONFIG,
      null,
      "default_language",
      "default_language should look like en or en-us",
    );
  }

  return {
    mode,
    errors,
    defaultLanguage,
  };
}

function parseAttributeGroups(workbook, config, issues) {
  const parsed = parseSheetWithHeaders(workbook, SHEETS.GROUPS, GROUP_HEADERS);
  if (!parsed.exists) {
    return { payload: [], byIdentifier: new Map() };
  }

  if (parsed.missingHeaders.length > 0) {
    parsed.missingHeaders.forEach((header) => {
      addIssue(issues, "error", SHEETS.GROUPS, 1, header, "Missing required header");
    });
  }

  const byIdentifier = new Map();
  const payload = [];

  parsed.rows.forEach(({ rowNumber, data }) => {
    if (!rowHasAnyValue(data, GROUP_HEADERS)) {
      return;
    }

    const identifier = normalizeIdentifier(data.identifier);
    const name = toText(data.name_en);
    const orderText = toText(data.order);
    const visibleText = toText(data.visible);

    if (!identifier) {
      addIssue(issues, "error", SHEETS.GROUPS, rowNumber, "identifier", "identifier is required");
      return;
    }

    if (!name) {
      addIssue(issues, "error", SHEETS.GROUPS, rowNumber, "name_en", "name_en is required");
      return;
    }

    if (byIdentifier.has(identifier)) {
      addIssue(issues, "error", SHEETS.GROUPS, rowNumber, "identifier", "Duplicate group identifier");
      return;
    }

    let order;
    if (orderText) {
      const parsedOrder = parseInteger(orderText);
      if (!parsedOrder.ok) {
        addIssue(issues, "error", SHEETS.GROUPS, rowNumber, "order", parsedOrder.error);
        return;
      }
      order = parsedOrder.value;
    }

    let visible;
    if (visibleText) {
      const parsedVisible = parseBoolean(visibleText);
      if (!parsedVisible.ok) {
        addIssue(issues, "error", SHEETS.GROUPS, rowNumber, "visible", parsedVisible.error);
        return;
      }
      visible = parsedVisible.value;
    }

    const optionsParsed = parseJsonObject(data.options_json, true);
    if (!optionsParsed.ok) {
      addIssue(issues, "error", SHEETS.GROUPS, rowNumber, "options_json", optionsParsed.error);
      return;
    }

    const request = {
      identifier,
      name: { [config.defaultLanguage]: name },
    };

    if (order !== undefined) request.order = order;
    if (visible !== undefined) request.visible = visible;
    if (Object.keys(optionsParsed.value).length > 0) request.options = optionsParsed.value;

    byIdentifier.set(identifier, request);
    payload.push(request);
  });

  return { payload, byIdentifier };
}

function parseTypes(workbook, config, issues) {
  const parsed = parseSheetWithHeaders(workbook, SHEETS.TYPES, TYPE_HEADERS);
  if (!parsed.exists) {
    return { payload: [], byIdentifier: new Map() };
  }

  if (parsed.missingHeaders.length > 0) {
    parsed.missingHeaders.forEach((header) => {
      addIssue(issues, "error", SHEETS.TYPES, 1, header, "Missing required header");
    });
  }

  const byIdentifier = new Map();
  const payload = [];

  parsed.rows.forEach(({ rowNumber, data }) => {
    if (!rowHasAnyValue(data, TYPE_HEADERS)) {
      return;
    }

    const identifier = normalizeIdentifier(data.identifier);
    const name = toText(data.name_en);
    const parentIdentifier = normalizeIdentifier(data.parent_identifier);

    if (!identifier) {
      addIssue(issues, "error", SHEETS.TYPES, rowNumber, "identifier", "identifier is required");
      return;
    }

    if (!name) {
      addIssue(issues, "error", SHEETS.TYPES, rowNumber, "name_en", "name_en is required");
      return;
    }

    if (byIdentifier.has(identifier)) {
      addIssue(issues, "error", SHEETS.TYPES, rowNumber, "identifier", "Duplicate type identifier");
      return;
    }

    let fileValue;
    if (!isBlank(data.file)) {
      const parsedFile = parseBoolean(data.file);
      if (!parsedFile.ok) {
        addIssue(issues, "error", SHEETS.TYPES, rowNumber, "file", parsedFile.error);
        return;
      }
      fileValue = parsedFile.value;
    }

    const request = {
      identifier,
      name: { [config.defaultLanguage]: name },
    };

    if (parentIdentifier) {
      request.parentIdentifier = parentIdentifier;
    }

    const icon = toText(data.icon);
    const iconColor = toText(data.icon_color);

    if (icon) request.icon = icon;
    if (iconColor) request.iconColor = iconColor;
    if (fileValue !== undefined) request.file = fileValue;

    byIdentifier.set(identifier, request);
    payload.push(request);
  });

  payload.forEach((typeRequest) => {
    if (typeRequest.parentIdentifier && !byIdentifier.has(typeRequest.parentIdentifier)) {
      addIssue(
        issues,
        "warning",
        SHEETS.TYPES,
        null,
        "parent_identifier",
        `Parent type '${typeRequest.parentIdentifier}' is not defined in this workbook. It must already exist in PIM.`,
      );
    }

    if (typeRequest.parentIdentifier === typeRequest.identifier) {
      addIssue(
        issues,
        "error",
        SHEETS.TYPES,
        null,
        "parent_identifier",
        `Type '${typeRequest.identifier}' cannot reference itself as parent.`,
      );
    }
  });

  return { payload, byIdentifier };
}

function parseTypeGroupBindings(workbook, groupIdentifiers, typeIdentifiers, issues) {
  const parsed = parseSheetWithHeaders(
    workbook,
    SHEETS.TYPE_GROUP_BINDINGS,
    TYPE_GROUP_BINDING_HEADERS,
  );

  if (!parsed.exists) {
    return new Map();
  }

  if (parsed.missingHeaders.length > 0) {
    parsed.missingHeaders.forEach((header) => {
      addIssue(issues, "error", SHEETS.TYPE_GROUP_BINDINGS, 1, header, "Missing required header");
    });
  }

  const bindingMap = new Map();

  parsed.rows.forEach(({ rowNumber, data }) => {
    if (!rowHasAnyValue(data, TYPE_GROUP_BINDING_HEADERS)) {
      return;
    }

    const groupIdentifier = normalizeIdentifier(data.group_identifier);
    const typeIdentifier = normalizeIdentifier(data.type_identifier);

    if (!groupIdentifier) {
      addIssue(
        issues,
        "error",
        SHEETS.TYPE_GROUP_BINDINGS,
        rowNumber,
        "group_identifier",
        "group_identifier is required",
      );
      return;
    }

    if (!typeIdentifier) {
      addIssue(
        issues,
        "error",
        SHEETS.TYPE_GROUP_BINDINGS,
        rowNumber,
        "type_identifier",
        "type_identifier is required",
      );
      return;
    }

    if (!groupIdentifiers.has(groupIdentifier)) {
      addIssue(
        issues,
        "warning",
        SHEETS.TYPE_GROUP_BINDINGS,
        rowNumber,
        "group_identifier",
        `Group '${groupIdentifier}' is not defined in Attribute_Groups sheet`,
      );
    }

    if (!typeIdentifiers.has(typeIdentifier)) {
      addIssue(
        issues,
        "warning",
        SHEETS.TYPE_GROUP_BINDINGS,
        rowNumber,
        "type_identifier",
        `Type '${typeIdentifier}' is not defined in Types sheet`,
      );
    }

    const validParsed = isBlank(data.valid) ? { ok: true, value: true } : parseBoolean(data.valid);
    const visibleParsed = isBlank(data.visible)
      ? { ok: true, value: true }
      : parseBoolean(data.visible);

    if (!validParsed.ok) {
      addIssue(issues, "error", SHEETS.TYPE_GROUP_BINDINGS, rowNumber, "valid", validParsed.error);
      return;
    }

    if (!visibleParsed.ok) {
      addIssue(issues, "error", SHEETS.TYPE_GROUP_BINDINGS, rowNumber, "visible", visibleParsed.error);
      return;
    }

    const current = bindingMap.get(groupIdentifier) || [];
    current.push({
      typeIdentifier,
      valid: validParsed.value,
      visible: visibleParsed.value,
    });
    bindingMap.set(groupIdentifier, current);
  });

  return bindingMap;
}

function parseAttributes(workbook, config, groupMap, typeMap, bindingMap, issues) {
  const parsed = parseSheetWithHeaders(workbook, SHEETS.ATTRIBUTES, ATTRIBUTE_HEADERS);
  if (!parsed.exists) {
    return { payload: [], byIdentifier: new Map() };
  }

  if (parsed.missingHeaders.length > 0) {
    parsed.missingHeaders.forEach((header) => {
      addIssue(issues, "error", SHEETS.ATTRIBUTES, 1, header, "Missing required header");
    });
  }

  const byIdentifier = new Map();
  const payload = [];

  parsed.rows.forEach(({ rowNumber, data }) => {
    if (!rowHasAnyValue(data, ATTRIBUTE_HEADERS)) {
      return;
    }

    const identifier = normalizeIdentifier(data.identifier);
    const name = toText(data.name_en);
    const typeRaw = toText(data.type_code);

    if (!identifier) {
      addIssue(issues, "error", SHEETS.ATTRIBUTES, rowNumber, "identifier", "identifier is required");
      return;
    }

    if (!name) {
      addIssue(issues, "error", SHEETS.ATTRIBUTES, rowNumber, "name_en", "name_en is required");
      return;
    }

    if (!typeRaw) {
      addIssue(issues, "error", SHEETS.ATTRIBUTES, rowNumber, "type_code", "type_code is required");
      return;
    }

    if (byIdentifier.has(identifier)) {
      addIssue(issues, "error", SHEETS.ATTRIBUTES, rowNumber, "identifier", "Duplicate attribute identifier");
      return;
    }

    const parsedType = parseInteger(typeRaw);
    if (!parsedType.ok || !ATTRIBUTE_TYPE_CODES.has(parsedType.value)) {
      addIssue(
        issues,
        "error",
        SHEETS.ATTRIBUTES,
        rowNumber,
        "type_code",
        "type_code must be one of 1..8",
      );
      return;
    }

    const groups = parseCsvList(data.groups_csv);
    if (groups.length === 0) {
      addIssue(
        issues,
        "error",
        SHEETS.ATTRIBUTES,
        rowNumber,
        "groups_csv",
        "At least one group identifier is required",
      );
      return;
    }

    const unknownGroups = groups.filter((groupIdentifier) => !groupMap.has(groupIdentifier));
    if (unknownGroups.length > 0) {
      addIssue(
        issues,
        "error",
        SHEETS.ATTRIBUTES,
        rowNumber,
        "groups_csv",
        `Unknown group identifiers: ${unknownGroups.join(", ")}`,
      );
      return;
    }

    let order;
    if (!isBlank(data.order)) {
      const parsedOrder = parseInteger(data.order);
      if (!parsedOrder.ok) {
        addIssue(issues, "error", SHEETS.ATTRIBUTES, rowNumber, "order", parsedOrder.error);
        return;
      }
      order = parsedOrder.value;
    }

    const languageDependentParsed = isBlank(data.language_dependent)
      ? { ok: true, value: false }
      : parseBoolean(data.language_dependent);

    const richTextParsed = isBlank(data.rich_text)
      ? { ok: true, value: false }
      : parseBoolean(data.rich_text);

    const multiLineParsed = isBlank(data.multi_line)
      ? { ok: true, value: false }
      : parseBoolean(data.multi_line);

    if (!languageDependentParsed.ok) {
      addIssue(
        issues,
        "error",
        SHEETS.ATTRIBUTES,
        rowNumber,
        "language_dependent",
        languageDependentParsed.error,
      );
      return;
    }

    if (!richTextParsed.ok) {
      addIssue(issues, "error", SHEETS.ATTRIBUTES, rowNumber, "rich_text", richTextParsed.error);
      return;
    }

    if (!multiLineParsed.ok) {
      addIssue(issues, "error", SHEETS.ATTRIBUTES, rowNumber, "multi_line", multiLineParsed.error);
      return;
    }

    const optionsParsed = parseJsonObject(data.options_json, true);
    if (!optionsParsed.ok) {
      addIssue(issues, "error", SHEETS.ATTRIBUTES, rowNumber, "options_json", optionsParsed.error);
      return;
    }

    const validSet = new Set(parseCsvList(data.valid_types_csv));
    const visibleSet = new Set(parseCsvList(data.visible_types_csv));

    for (const groupIdentifier of groups) {
      const bindings = bindingMap.get(groupIdentifier) || [];
      for (const binding of bindings) {
        if (binding.valid) {
          validSet.add(binding.typeIdentifier);
        }
        if (binding.visible) {
          visibleSet.add(binding.typeIdentifier);
        }
      }
    }

    for (const typeIdentifier of [...validSet, ...visibleSet]) {
      if (!typeMap.has(typeIdentifier)) {
        addIssue(
          issues,
          "warning",
          SHEETS.ATTRIBUTES,
          rowNumber,
          "valid_types_csv",
          `Type '${typeIdentifier}' is not in Types sheet. It must already exist in PIM.`,
        );
      }
    }

    const request = {
      identifier,
      name: { [config.defaultLanguage]: name },
      groups,
      type: parsedType.value,
      languageDependent: languageDependentParsed.value,
      richText: richTextParsed.value,
      multiLine: multiLineParsed.value,
    };

    if (order !== undefined) request.order = order;

    const pattern = toText(data.pattern);
    if (pattern) request.pattern = pattern;

    const lovIdentifier = normalizeIdentifier(data.lov_identifier);
    if (lovIdentifier) request.lov = lovIdentifier;

    if (validSet.size > 0) request.valid = [...validSet];
    if (visibleSet.size > 0) request.visible = [...visibleSet];
    if (Object.keys(optionsParsed.value).length > 0) request.options = optionsParsed.value;

    const model = {
      identifier,
      type: parsedType.value,
      languageDependent: languageDependentParsed.value,
    };

    byIdentifier.set(identifier, model);
    payload.push(request);
  });

  return { payload, byIdentifier };
}

function collectDeclaredItemIdentifiers(workbook) {
  const sheetNames = [SHEETS.ITEM_PARENTS, ...PRODUCT_SHEETS];
  const declared = new Set();

  sheetNames.forEach((sheetName) => {
    const parsed = parseSheetWithHeaders(workbook, sheetName, ITEM_BASE_HEADERS);
    if (!parsed.exists) {
      return;
    }

    parsed.rows.forEach(({ data }) => {
      const identifier = normalizeIdentifier(data.identifier);
      if (identifier) {
        declared.add(identifier);
      }
    });
  });

  return declared;
}

function parseItemSheet(
  workbook,
  sheetName,
  config,
  attributeMap,
  typeMap,
  seenItemIdentifiers,
  declaredItemIdentifiers,
  issues,
) {
  const parsed = parseSheetWithHeaders(workbook, sheetName, ITEM_BASE_HEADERS);
  if (!parsed.exists) {
    return [];
  }

  if (parsed.missingHeaders.length > 0) {
    parsed.missingHeaders.forEach((header) => {
      addIssue(issues, "error", sheetName, 1, header, "Missing required header");
    });
  }

  const attrHeaders = parsed.headers.filter((header) => header.startsWith("attr:"));
  const payload = [];

  parsed.rows.forEach(({ rowNumber, data }) => {
    if (!rowHasAnyValue(data, parsed.headers)) {
      return;
    }

    const identifier = normalizeIdentifier(data.identifier);
    const name = toText(data.name_en);
    const typeIdentifier = normalizeIdentifier(data.type_identifier);
    const parentIdentifier = normalizeIdentifier(data.parent_identifier);

    if (!identifier) {
      addIssue(issues, "error", sheetName, rowNumber, "identifier", "identifier is required");
      return;
    }

    if (!name) {
      addIssue(issues, "error", sheetName, rowNumber, "name_en", "name_en is required");
      return;
    }

    if (!typeIdentifier) {
      addIssue(issues, "error", sheetName, rowNumber, "type_identifier", "type_identifier is required");
      return;
    }

    if (seenItemIdentifiers.has(identifier)) {
      addIssue(
        issues,
        "error",
        sheetName,
        rowNumber,
        "identifier",
        `Duplicate item identifier '${identifier}' across item sheets`,
      );
      return;
    }

    seenItemIdentifiers.add(identifier);

    if (!typeMap.has(typeIdentifier)) {
      addIssue(
        issues,
        "warning",
        sheetName,
        rowNumber,
        "type_identifier",
        `Type '${typeIdentifier}' is not in Types sheet. It must already exist in PIM.`,
      );
    }

    const typeMeta = typeMap.get(typeIdentifier);
    if (typeMeta?.parentIdentifier && !parentIdentifier) {
      addIssue(
        issues,
        "error",
        sheetName,
        rowNumber,
        "parent_identifier",
        `type '${typeIdentifier}' is child of '${typeMeta.parentIdentifier}', so parent_identifier is required`,
      );
      return;
    }

    if (parentIdentifier === identifier) {
      addIssue(
        issues,
        "error",
        sheetName,
        rowNumber,
        "parent_identifier",
        "parent_identifier cannot reference the same item identifier",
      );
      return;
    }

    if (
      parentIdentifier &&
      !seenItemIdentifiers.has(parentIdentifier) &&
      !declaredItemIdentifiers.has(parentIdentifier)
    ) {
      addIssue(
        issues,
        "warning",
        sheetName,
        rowNumber,
        "parent_identifier",
        `Parent item '${parentIdentifier}' is not declared in workbook. It must already exist in PIM.`,
      );
    }

    const valuesParsed = parseJsonObject(data.values_json, true);
    if (!valuesParsed.ok) {
      addIssue(issues, "error", sheetName, rowNumber, "values_json", valuesParsed.error);
      return;
    }

    const channelsParsed = parseJsonObject(data.channels_json, true);
    if (!channelsParsed.ok) {
      addIssue(issues, "error", sheetName, rowNumber, "channels_json", channelsParsed.error);
      return;
    }

    const values = { ...valuesParsed.value };

    for (const [valueKey, valueRaw] of Object.entries(valuesParsed.value)) {
      const attributeIdentifier = normalizeIdentifier(valueKey);
      const attribute = attributeMap.get(attributeIdentifier);
      if (!attribute) {
        addIssue(
          issues,
          "error",
          sheetName,
          rowNumber,
          "values_json",
          `values_json contains unknown attribute '${valueKey}'`,
        );
        continue;
      }

      const coerced = coerceAttributeValue(valueRaw, attribute, config.defaultLanguage);
      if (!coerced.ok) {
        addIssue(
          issues,
          "error",
          sheetName,
          rowNumber,
          `values_json.${valueKey}`,
          `Invalid value: ${coerced.error}`,
        );
        continue;
      }

      if (!coerced.skip) {
        values[attributeIdentifier] = coerced.value;
      }
    }

    for (const header of attrHeaders) {
      const attributeIdentifier = normalizeIdentifier(header.slice(5));
      if (!attributeIdentifier) {
        continue;
      }

      if (isBlank(data[header])) {
        continue;
      }

      const attribute = attributeMap.get(attributeIdentifier);
      if (!attribute) {
        addIssue(
          issues,
          "error",
          sheetName,
          rowNumber,
          header,
          `Unknown attribute '${attributeIdentifier}'`,
        );
        continue;
      }

      const coerced = coerceAttributeValue(data[header], attribute, config.defaultLanguage);
      if (!coerced.ok) {
        addIssue(
          issues,
          "error",
          sheetName,
          rowNumber,
          header,
          `Invalid value: ${coerced.error}`,
        );
        continue;
      }

      if (!coerced.skip) {
        values[attributeIdentifier] = coerced.value;
      }
    }

    const request = {
      identifier,
      typeIdentifier,
      name: { [config.defaultLanguage]: name },
    };

    if (parentIdentifier) {
      request.parentIdentifier = parentIdentifier;
    }

    if (Object.keys(values).length > 0) {
      request.values = values;
    }

    if (Object.keys(channelsParsed.value).length > 0) {
      request.channels = channelsParsed.value;
    }

    payload.push(request);
  });

  return payload;
}

function splitIssues(issues) {
  const errors = issues.filter((issue) => issue.severity === "error");
  const warnings = issues.filter((issue) => issue.severity === "warning");
  return { errors, warnings };
}

export function parseAndValidateImportWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const issues = [];

  ensureRequiredSheets(workbook, issues);

  const config = parseConfig(workbook, issues);

  const { payload: attrGroups, byIdentifier: groupMap } = parseAttributeGroups(
    workbook,
    config,
    issues,
  );

  const { payload: types, byIdentifier: typeMap } = parseTypes(workbook, config, issues);

  const bindingMap = parseTypeGroupBindings(
    workbook,
    new Set(groupMap.keys()),
    new Set(typeMap.keys()),
    issues,
  );

  const { payload: attributes, byIdentifier: attributeMap } = parseAttributes(
    workbook,
    config,
    groupMap,
    typeMap,
    bindingMap,
    issues,
  );

  const seenItemIdentifiers = new Set();
  const declaredItemIdentifiers = collectDeclaredItemIdentifiers(workbook);

  const parentItems = parseItemSheet(
    workbook,
    SHEETS.ITEM_PARENTS,
    config,
    attributeMap,
    typeMap,
    seenItemIdentifiers,
    declaredItemIdentifiers,
    issues,
  );

  const productItems = PRODUCT_SHEETS.flatMap((sheetName) =>
    parseItemSheet(
      workbook,
      sheetName,
      config,
      attributeMap,
      typeMap,
      seenItemIdentifiers,
      declaredItemIdentifiers,
      issues,
    ),
  );
  const items = [...parentItems, ...productItems];

  const { errors, warnings } = splitIssues(issues);

  const payload = {
    config: {
      mode: config.mode,
      errors: config.errors,
    },
    attrGroups,
    attributes,
    types,
    items,
  };

  const summary = {
    attrGroups: attrGroups.length,
    attributes: attributes.length,
    types: types.length,
    items: items.length,
    errors: errors.length,
    warnings: warnings.length,
  };

  return {
    payload,
    summary,
    errors,
    warnings,
    valid: errors.length === 0,
  };
}

export function formatIssue(issue) {
  const location = `${issue.sheet}${issue.row ? `:row ${issue.row}` : ""}`;
  const fieldPart = issue.field ? ` [${issue.field}]` : "";
  return `${location}${fieldPart} - ${issue.message}`;
}
