import { useCallback, useEffect, useMemo, useState } from "react";
import "./App.css";

const SCALAR_TYPES = new Set([
  "ID",
  "String",
  "Int",
  "Float",
  "Boolean",
  "JSON",
  "JSONObject",
  "LanguageDependentString",
  "UTCDateTime",
]);

const API_BASE = import.meta.env.VITE_API_BASE || "";

function makeKey(operation) {
  return `${operation.kind}:${operation.name}`;
}

function defaultByType(arg) {
  if (arg.list) return [];

  const baseType = arg.baseType;
  if (baseType === "Int" || baseType === "Float") return 0;
  if (baseType === "Boolean") return false;
  if (baseType === "JSONObject" || baseType === "JSON") return {};
  if (baseType === "LanguageDependentString") return { en: "" };
  if (SCALAR_TYPES.has(baseType)) return "";

  return {};
}

function buildArgsTemplate(argumentsList) {
  const template = {};
  for (const arg of argumentsList) {
    template[arg.name] = defaultByType(arg);
  }
  return template;
}

function parseJsonObject(raw) {
  if (!raw || !raw.trim()) {
    return {};
  }

  const parsed = JSON.parse(raw);
  if (parsed === null || Array.isArray(parsed) || typeof parsed !== "object") {
    throw new Error('Expected JSON object like {"arg": value}');
  }

  return parsed;
}

function apiUrl(path) {
  return `${API_BASE}${path}`;
}

function flattenImportResults(data) {
  const sections = ["types", "attrGroups", "attributes", "items"];
  return sections.flatMap((section) =>
    Array.isArray(data?.[section]) ? data[section] : [],
  );
}

function formatValidationIssue(issue) {
  const location = `${issue.sheet}${issue.row ? `:row ${issue.row}` : ""}`;
  const fieldPart = issue.field ? ` [${issue.field}]` : "";
  return `${location}${fieldPart} - ${issue.message}`;
}

function App() {
  const [catalog, setCatalog] = useState({ queries: [], mutations: [] });
  const [kindFilter, setKindFilter] = useState("ALL");
  const [selectedKey, setSelectedKey] = useState("");

  const [selectedFile, setSelectedFile] = useState(null);
  const [uploadStatus, setUploadStatus] = useState({ text: "", tone: "" });
  const [uploadLoading, setUploadLoading] = useState(false);

  const [argsInput, setArgsInput] = useState("{}");
  const [selectionInput, setSelectionInput] = useState("");
  const [selectionDisabled, setSelectionDisabled] = useState(true);
  const [selectionPlaceholder, setSelectionPlaceholder] = useState("");

  const [executeStatus, setExecuteStatus] = useState({ text: "", tone: "" });
  const [resultOutput, setResultOutput] = useState("{}\n");

  const [importFileName, setImportFileName] = useState("");
  const [importSummary, setImportSummary] = useState(null);
  const [importErrors, setImportErrors] = useState([]);
  const [importWarnings, setImportWarnings] = useState([]);
  const [importPayload, setImportPayload] = useState(null);
  const [importStatus, setImportStatus] = useState({ text: "", tone: "" });
  const [importResult, setImportResult] = useState("");
  const [importLoading, setImportLoading] = useState(false);

  const allOperations = useMemo(
    () => [...(catalog.queries || []), ...(catalog.mutations || [])],
    [catalog],
  );

  const filteredOperations = useMemo(() => {
    if (kindFilter === "ALL") {
      return allOperations;
    }
    return allOperations.filter((operation) => operation.kind === kindFilter);
  }, [allOperations, kindFilter]);

  const selectedOperation = useMemo(
    () =>
      allOperations.find((operation) => makeKey(operation) === selectedKey) ||
      null,
    [allOperations, selectedKey],
  );

  const operationMeta = useMemo(() => {
    if (!selectedOperation) {
      return "Operation is not selected";
    }

    const argsText = (selectedOperation.arguments || [])
      .map(
        (arg) =>
          `${arg.name}: ${arg.type}${arg.required ? " (required)" : ""}`,
      )
      .join("\n");

    return [
      `Operation: ${selectedOperation.kind}.${selectedOperation.name}`,
      `Return: ${selectedOperation.returnType}`,
      `Source: ${selectedOperation.sourceFile}`,
      "Arguments:",
      argsText || "none",
    ].join("\n");
  }, [selectedOperation]);

  const loadOperations = useCallback(async () => {
    setExecuteStatus({ text: "Loading operations...", tone: "" });

    try {
      const response = await fetch(apiUrl("/api/graphql/operations"));
      const payload = await response.json();

      if (!response.ok) {
        throw new Error(payload.error || "Failed to fetch operation catalog");
      }

      setCatalog(payload);
      const totalCount =
        (payload.queries?.length || 0) + (payload.mutations?.length || 0);
      setExecuteStatus({
        text: `Operations loaded: ${totalCount}`,
        tone: "ok",
      });
    } catch (error) {
      setExecuteStatus({ text: error.message, tone: "error" });
    }
  }, []);

  useEffect(() => {
    loadOperations();
  }, [loadOperations]);

  useEffect(() => {
    if (filteredOperations.length === 0) {
      if (selectedKey) {
        setSelectedKey("");
      }
      return;
    }

    const hasCurrentSelection = filteredOperations.some(
      (operation) => makeKey(operation) === selectedKey,
    );

    if (!hasCurrentSelection) {
      setSelectedKey(makeKey(filteredOperations[0]));
    }
  }, [filteredOperations, selectedKey]);

  useEffect(() => {
    if (!selectedOperation) {
      setArgsInput("{}");
      setSelectionInput("");
      setSelectionDisabled(true);
      setSelectionPlaceholder("");
      return;
    }

    setArgsInput(
      JSON.stringify(buildArgsTemplate(selectedOperation.arguments || []), null, 2),
    );

    if (selectedOperation.returnScalar) {
      setSelectionInput("");
      setSelectionDisabled(true);
      setSelectionPlaceholder("Selection set is not needed for scalar return values");
      return;
    }

    setSelectionDisabled(false);
    setSelectionPlaceholder("Example: id name { en }");
    setSelectionInput(selectedOperation.suggestedSelection || "");
  }, [selectedOperation]);

  const handleExecute = async () => {
    if (!selectedOperation) {
      setExecuteStatus({
        text: "Select operation first",
        tone: "error",
      });
      return;
    }

    let args;
    try {
      args = parseJsonObject(argsInput);
    } catch (error) {
      setExecuteStatus({
        text: `Arguments JSON: ${error.message}`,
        tone: "error",
      });
      return;
    }

    setExecuteStatus({
      text: `Executing ${selectedOperation.kind}.${selectedOperation.name}...`,
      tone: "",
    });

    const payload = {
      operationName: selectedOperation.name,
      kind: selectedOperation.kind,
      arguments: args,
      selectionSet: selectionDisabled ? "" : selectionInput,
    };

    try {
      const response = await fetch(apiUrl("/api/graphql/execute"), {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      const result = await response.json();
      setResultOutput(JSON.stringify(result, null, 2));

      if (!response.ok || result.errors) {
        setExecuteStatus({
          text: result.error || "Operation failed",
          tone: "error",
        });
        return;
      }

      setExecuteStatus({ text: "Operation executed successfully", tone: "ok" });
    } catch (error) {
      setExecuteStatus({ text: error.message, tone: "error" });
    }
  };

  const handleUpload = async () => {
    if (!selectedFile) {
      setUploadStatus({ text: "Select file", tone: "error" });
      return;
    }

    setUploadLoading(true);
    setUploadStatus({ text: "Uploading file...", tone: "" });

    const formData = new FormData();
    formData.append("file", selectedFile);

    try {
      const response = await fetch(apiUrl("/api/upload"), {
        method: "POST",
        body: formData,
      });
      const result = await response.json();

      if (!response.ok) {
        throw new Error(result.error || "File upload failed");
      }

      setUploadStatus({
        text: `Uploaded: ${result.originalName} (${result.size} bytes)`,
        tone: "ok",
      });
    } catch (error) {
      setUploadStatus({ text: error.message, tone: "error" });
    } finally {
      setUploadLoading(false);
    }
  };

  const handleImportFile = async (event) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setImportFileName(file.name);
    setImportStatus({ text: "Parsing and validating workbook...", tone: "" });
    setImportResult("");

    try {
      const { parseAndValidateImportWorkbook } = await import("./excelImportParser.js");
      const buffer = await file.arrayBuffer();
      const validation = parseAndValidateImportWorkbook(buffer);

      setImportPayload(validation.payload);
      setImportSummary(validation.summary);
      setImportErrors(validation.errors);
      setImportWarnings(validation.warnings);

      if (validation.valid) {
        setImportStatus({
          text: `Validation passed. ${validation.summary.items} items ready for import.`,
          tone: "ok",
        });
      } else {
        setImportStatus({
          text: `Validation failed with ${validation.summary.errors} error(s).`,
          tone: "error",
        });
      }
    } catch (error) {
      setImportPayload(null);
      setImportSummary(null);
      setImportErrors([
        {
          sheet: "Workbook",
          row: null,
          field: "file",
          message: error.message,
        },
      ]);
      setImportWarnings([]);
      setImportStatus({ text: `Failed to parse workbook: ${error.message}`, tone: "error" });
    }
  };

  const handleDownloadTemplate = async () => {
    const { downloadImportTemplate } = await import("./excelImportTemplate.js");
    downloadImportTemplate();
  };

  const handlePushImport = async () => {
    if (!importPayload) {
      setImportStatus({ text: "Load and validate workbook first", tone: "error" });
      return;
    }

    if (importErrors.length > 0) {
      setImportStatus({ text: "Fix validation errors before import", tone: "error" });
      return;
    }

    setImportLoading(true);
    setImportStatus({ text: "Pushing data to PIM via import mutation...", tone: "" });

    try {
      const response = await fetch(apiUrl("/api/import/execute"), {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(importPayload),
      });

      const result = await response.json();
      setImportResult(JSON.stringify(result, null, 2));

      if (!response.ok || result.errors || result.error) {
        setImportStatus({
          text: result.error || "Import mutation failed",
          tone: "error",
        });
        return;
      }

      const rows = flattenImportResults(result.data);
      const rejectedCount = rows.filter((row) => row?.result === "REJECTED").length;
      const warningCount = rows.reduce(
        (total, row) => total + (Array.isArray(row?.warnings) ? row.warnings.length : 0),
        0,
      );

      if (rejectedCount > 0) {
        setImportStatus({
          text: `Import completed with ${rejectedCount} rejected row(s). Check response details below.`,
          tone: "error",
        });
        return;
      }

      setImportStatus({
        text: `Import completed successfully. Warnings: ${warningCount}.`,
        tone: "ok",
      });
    } catch (error) {
      setImportStatus({ text: error.message, tone: "error" });
    } finally {
      setImportLoading(false);
    }
  };

  const errorPreview = useMemo(
    () => importErrors.slice(0, 120).map(formatValidationIssue).join("\n"),
    [importErrors],
  );

  const warningPreview = useMemo(
    () => importWarnings.slice(0, 120).map(formatValidationIssue).join("\n"),
    [importWarnings],
  );

  return (
    <div className="page">
      <div className="header">
        <h1>PIM Console</h1>
        <p>
          Excel template import, file upload, and universal GraphQL explorer for
          schema operations.
        </p>
      </div>

      <div className="grid">
        <section className="panel wide">
          <h2>Excel Import Pipeline</h2>

          <div className="button-row">
            <button className="secondary" onClick={handleDownloadTemplate}>
              Download Excel Template
            </button>
          </div>

          <p className="import-note">
            Required product sheets inside template: <code>TCT_Router_Bit</code>,{" "}
            <code>Insert_Tool</code>, <code>Countersink</code>. Use metadata sheet{" "}
            <code>Item_Parents</code> to create parent items for child product types.
          </p>

          <div className="field">
            <label htmlFor="importFileInput">Excel file (.xlsx/.xlsm/.xls)</label>
            <input
              id="importFileInput"
              type="file"
              accept=".xlsx,.xlsm,.xls"
              onChange={handleImportFile}
            />
          </div>

          {importFileName && importSummary && (
            <div className="meta">
              {[
                `File: ${importFileName}`,
                `Attribute Groups: ${importSummary.attrGroups}`,
                `Attributes: ${importSummary.attributes}`,
                `Types: ${importSummary.types}`,
                `Items: ${importSummary.items}`,
                `Validation Errors: ${importSummary.errors}`,
                `Validation Warnings: ${importSummary.warnings}`,
              ].join("\n")}
            </div>
          )}

          <div className="button-row">
            <button
              onClick={handlePushImport}
              disabled={!importPayload || importErrors.length > 0 || importLoading}
            >
              {importLoading ? "Importing..." : "Push Validated Data to PIM"}
            </button>
          </div>

          <div className={`status ${importStatus.tone}`.trim()}>{importStatus.text}</div>

          {importErrors.length > 0 && (
            <div className="field">
              <label>Validation Errors</label>
              <pre className="issue-block error">{errorPreview}</pre>
              {importErrors.length > 120 && (
                <div className="import-note">
                  Showing first 120 errors out of {importErrors.length}.
                </div>
              )}
            </div>
          )}

          {importWarnings.length > 0 && (
            <div className="field">
              <label>Validation Warnings</label>
              <pre className="issue-block warning">{warningPreview}</pre>
              {importWarnings.length > 120 && (
                <div className="import-note">
                  Showing first 120 warnings out of {importWarnings.length}.
                </div>
              )}
            </div>
          )}

          {importResult && (
            <div className="field">
              <label>Import Mutation Response</label>
              <pre>{importResult}</pre>
            </div>
          )}
        </section>

        <section className="panel">
          <h2>Upload</h2>
          <div className="field">
            <label htmlFor="fileInput">File</label>
            <input
              id="fileInput"
              type="file"
              onChange={(event) => setSelectedFile(event.target.files?.[0] || null)}
            />
          </div>
          <div className="button-row">
            <button onClick={handleUpload} disabled={uploadLoading}>
              {uploadLoading ? "Uploading..." : "Upload"}
            </button>
          </div>
          <div className={`status ${uploadStatus.tone}`.trim()}>
            {uploadStatus.text}
          </div>
        </section>

        <section className="panel">
          <h2>GraphQL Explorer</h2>
          <div className="field">
            <label htmlFor="kindFilter">Operation Kind</label>
            <select
              id="kindFilter"
              value={kindFilter}
              onChange={(event) => setKindFilter(event.target.value)}
            >
              <option value="ALL">All</option>
              <option value="QUERY">Query</option>
              <option value="MUTATION">Mutation</option>
            </select>
          </div>

          <div className="field">
            <label htmlFor="operationSelect">Operation</label>
            <select
              id="operationSelect"
              value={selectedKey}
              onChange={(event) => setSelectedKey(event.target.value)}
              disabled={filteredOperations.length === 0}
            >
              {filteredOperations.length === 0 ? (
                <option value="">No operations found</option>
              ) : (
                filteredOperations.map((operation) => (
                  <option key={makeKey(operation)} value={makeKey(operation)}>
                    {operation.kind}.{operation.name}
                  </option>
                ))
              )}
            </select>
          </div>

          <div className="button-row">
            <button className="secondary" onClick={loadOperations}>
              Reload Operations
            </button>
          </div>

          <div className="meta">{operationMeta}</div>
        </section>

        <section className="panel wide">
          <h2>Execute</h2>
          <div className="field">
            <label htmlFor="argsInput">Arguments (JSON)</label>
            <textarea
              id="argsInput"
              spellCheck="false"
              value={argsInput}
              onChange={(event) => setArgsInput(event.target.value)}
            />
          </div>

          <div className="field">
            <label htmlFor="selectionInput">Selection Set</label>
            <textarea
              id="selectionInput"
              spellCheck="false"
              value={selectionInput}
              disabled={selectionDisabled}
              placeholder={selectionPlaceholder}
              onChange={(event) => setSelectionInput(event.target.value)}
            />
          </div>

          <div className="button-row">
            <button onClick={handleExecute}>Execute</button>
          </div>

          <div className={`status ${executeStatus.tone}`.trim()}>
            {executeStatus.text}
          </div>
        </section>

        <section className="panel wide">
          <h2>Result</h2>
          <pre>{resultOutput}</pre>
        </section>
      </div>
    </div>
  );
}

export default App;
