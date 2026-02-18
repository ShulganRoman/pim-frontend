import React, { useState, useRef } from "react";
import * as XLSX from "xlsx/dist/xlsx.full.min.js";
import "./App.css";

function App() {
  const [fileName, setFileName] = useState("");
  const [data, setData] = useState([]);
  const [status, setStatus] = useState("");
  const [loading, setLoading] = useState(false);
  const fileInputRef = useRef(null);
  const selectedFileRef = useRef(null);

  const handleFileSelect = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setFileName(file.name);
    selectedFileRef.current = file;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const workbook = XLSX.read(evt.target.result, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      setData(XLSX.utils.sheet_to_json(sheet));
    };
    reader.readAsBinaryString(file);
  };

  const handleUpload = async () => {
    if (!selectedFileRef.current) {
      setStatus("Выберите файл!");
      return;
    }

    setLoading(true);
    setStatus("");

    const formData = new FormData();
    formData.append("file", selectedFileRef.current);

    try {
      const response = await fetch("/api/upload", {
        method: "POST",
        body: formData,
      });

      const result = await response.json();

      if (response.ok) {
        setStatus(`✅ ${result.message}: ${result.fileName}`);
      } else {
        setStatus(`❌ ${result.error || "Ошибка загрузки"}`);
      }
    } catch (err) {
      console.error(err);
      setStatus("❌ Ошибка при отправке файла!");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="app-wrapper">
      <header className="header">
        <h1>PIM Project</h1>
      </header>

      <div className="app-container">
        <h1 className="title">PIM</h1>

        <div className="upload-section">
          <button
            className="custom-file-button"
            onClick={() => fileInputRef.current.click()}
          >
            Select file
          </button>
          <input
            type="file"
            accept=".xlsx,.xls"
            ref={fileInputRef}
            onChange={handleFileSelect}
            style={{ display: "none" }}
          />

          {fileName && <p>Selected file: {fileName}</p>}

          {fileName && (
            <button
              className="custom-file-button"
              onClick={handleUpload}
              disabled={loading}
            >
              {loading ? "Uploading..." : "Upload to server"}
            </button>
          )}

          {status && <p className="status">{status}</p>}
        </div>

        {data.length > 0 && (
          <div className="data-preview">
            <h2>Preview:</h2>
            <pre>{JSON.stringify(data, null, 2)}</pre>
          </div>
        )}
      </div>

      <footer className="footer">
        <p>© 2026 PIM Project. All rights reserved.</p>
      </footer>
    </div>
  );
}

export default App;
