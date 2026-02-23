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
    throw new Error('Ожидался JSON-объект вида {"arg": value}');
  }

  return parsed;
}

function apiUrl(path) {
  return `${API_BASE}${path}`;
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
      return "Операция не выбрана";
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
      argsText || "нет",
    ].join("\n");
  }, [selectedOperation]);

  const loadOperations = useCallback(async () => {
    setExecuteStatus({ text: "Загрузка операций...", tone: "" });

    try {
      const response = await fetch(apiUrl("/api/graphql/operations"));
      const payload = await response.json();

      if (!response.ok) {
        throw new Error(payload.error || "Не удалось получить список операций");
      }

      setCatalog(payload);
      const totalCount =
        (payload.queries?.length || 0) + (payload.mutations?.length || 0);
      setExecuteStatus({
        text: `Операций загружено: ${totalCount}`,
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
      setSelectionPlaceholder("Для scalar-результата selection set не нужен");
      return;
    }

    setSelectionDisabled(false);
    setSelectionPlaceholder("Например: id name { en }");
    setSelectionInput(selectedOperation.suggestedSelection || "");
  }, [selectedOperation]);

  const handleExecute = async () => {
    if (!selectedOperation) {
      setExecuteStatus({
        text: "Сначала выберите операцию",
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
      text: `Выполняю ${selectedOperation.kind}.${selectedOperation.name}...`,
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
          text: result.error || "Операция выполнена с ошибкой",
          tone: "error",
        });
        return;
      }

      setExecuteStatus({ text: "Операция выполнена успешно", tone: "ok" });
    } catch (error) {
      setExecuteStatus({ text: error.message, tone: "error" });
    }
  };

  const handleUpload = async () => {
    if (!selectedFile) {
      setUploadStatus({ text: "Выберите файл", tone: "error" });
      return;
    }

    setUploadLoading(true);
    setUploadStatus({ text: "Загрузка файла...", tone: "" });

    const formData = new FormData();
    formData.append("file", selectedFile);

    try {
      const response = await fetch(apiUrl("/api/upload"), {
        method: "POST",
        body: formData,
      });
      const result = await response.json();

      if (!response.ok) {
        throw new Error(result.error || "Ошибка загрузки файла");
      }

      setUploadStatus({
        text: `Файл загружен: ${result.originalName} (${result.size} bytes)`,
        tone: "ok",
      });
    } catch (error) {
      setUploadStatus({ text: error.message, tone: "error" });
    } finally {
      setUploadLoading(false);
    }
  };

  return (
    <div className="page">
      <div className="header">
        <h1>PIM Console</h1>
        <p>
          Загрузка файлов + универсальный GraphQL Explorer для операций из
          schema/*.graphql
        </p>
      </div>

      <div className="grid">
        <section className="panel">
          <h2>Upload</h2>
          <div className="field">
            <label htmlFor="fileInput">Файл</label>
            <input
              id="fileInput"
              type="file"
              onChange={(event) => setSelectedFile(event.target.files?.[0] || null)}
            />
          </div>
          <div className="button-row">
            <button onClick={handleUpload} disabled={uploadLoading}>
              {uploadLoading ? "Загрузка..." : "Загрузить"}
            </button>
          </div>
          <div className={`status ${uploadStatus.tone}`.trim()}>
            {uploadStatus.text}
          </div>
        </section>

        <section className="panel">
          <h2>GraphQL Explorer</h2>
          <div className="field">
            <label htmlFor="kindFilter">Тип операции</label>
            <select
              id="kindFilter"
              value={kindFilter}
              onChange={(event) => setKindFilter(event.target.value)}
            >
              <option value="ALL">Все</option>
              <option value="QUERY">Query</option>
              <option value="MUTATION">Mutation</option>
            </select>
          </div>

          <div className="field">
            <label htmlFor="operationSelect">Операция</label>
            <select
              id="operationSelect"
              value={selectedKey}
              onChange={(event) => setSelectedKey(event.target.value)}
              disabled={filteredOperations.length === 0}
            >
              {filteredOperations.length === 0 ? (
                <option value="">Операции не найдены</option>
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
              Обновить операции
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
            <button onClick={handleExecute}>Выполнить</button>
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
