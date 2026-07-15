import React, { useState, useRef, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import { Upload, X, Plus, Download, Search, Trash2, FileSpreadsheet, Pencil } from "lucide-react";

// ---------------------------------------------
// 合併簿 — Excel 合併與線上編輯工具
// 設計語言：帳冊紙感、印章紅、鋼筆藍墨
// ---------------------------------------------

let uid = 0;
const nextId = () => `f${++uid}-${Date.now()}`;

export default function ExcelMerger() {
  const [files, setFiles] = useState([]); // {id, name, sheetNames, selectedSheet, sheetsData: {name: rows[]}}
  const [columns, setColumns] = useState([]);
  const [rows, setRows] = useState([]);
  const [merged, setMerged] = useState(false);
  const [search, setSearch] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [editingHeader, setEditingHeader] = useState(null);
  const [toast, setToast] = useState(null);
  const inputRef = useRef(null);

  const showToast = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 2200);
  };

  // ---------- 讀取檔案 ----------
  const handleFiles = useCallback((fileList) => {
    const arr = Array.from(fileList).filter((f) =>
      /\.(xlsx|xls|csv)$/i.test(f.name)
    );
    if (arr.length === 0) {
      showToast("請上傳 .xlsx、.xls 或 .csv 檔案");
      return;
    }
    arr.forEach((file) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: "array" });
          const sheetsData = {};
          wb.SheetNames.forEach((sn) => {
            const ws = wb.Sheets[sn];
            sheetsData[sn] = XLSX.utils.sheet_to_json(ws, { defval: "" });
          });
          setFiles((prev) => [
            ...prev,
            {
              id: nextId(),
              name: file.name.replace(/\.(xlsx|xls|csv)$/i, ""),
              sheetNames: wb.SheetNames,
              selectedSheet: wb.SheetNames[0],
              sheetsData,
            },
          ]);
        } catch (err) {
          showToast(`讀取「${file.name}」失敗`);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }, []);

  const removeFile = (id) => setFiles((prev) => prev.filter((f) => f.id !== id));
  const changeSheet = (id, sheetName) =>
    setFiles((prev) => prev.map((f) => (f.id === id ? { ...f, selectedSheet: sheetName } : f)));

  // ---------- 合併 ----------
  const mergeAll = () => {
    if (files.length === 0) {
      showToast("請先加入至少一份 Excel");
      return;
    }
    const colOrder = ["來源檔案"];
    const seen = new Set(colOrder);
    const mergedRows = [];

    files.forEach((f) => {
      const data = f.sheetsData[f.selectedSheet] || [];
      data.forEach((row) => {
        Object.keys(row).forEach((k) => {
          if (!seen.has(k)) {
            seen.add(k);
            colOrder.push(k);
          }
        });
        mergedRows.push({ 來源檔案: f.name, ...row });
      });
    });

    // normalize: ensure every row has every column
    const normalized = mergedRows.map((r) => {
      const o = {};
      colOrder.forEach((c) => (o[c] = r[c] !== undefined ? r[c] : ""));
      return o;
    });

    setColumns(colOrder);
    setRows(normalized);
    setMerged(true);
    showToast(`已合併 ${files.length} 份檔案，共 ${normalized.length} 筆資料`);
  };

  // ---------- 編輯 ----------
  const updateCell = (rowIdx, col, value) => {
    setRows((prev) => {
      const next = [...prev];
      next[rowIdx] = { ...next[rowIdx], [col]: value };
      return next;
    });
  };

  const addRow = () => {
    const blank = {};
    columns.forEach((c) => (blank[c] = ""));
    setRows((prev) => [...prev, blank]);
  };

  const removeRow = (idx) => setRows((prev) => prev.filter((_, i) => i !== idx));

  const addColumn = () => {
    let name = "新欄位";
    let n = 1;
    while (columns.includes(name)) name = `新欄位${++n}`;
    setColumns((prev) => [...prev, name]);
    setRows((prev) => prev.map((r) => ({ ...r, [name]: "" })));
  };

  const removeColumn = (col) => {
    setColumns((prev) => prev.filter((c) => c !== col));
    setRows((prev) =>
      prev.map((r) => {
        const o = { ...r };
        delete o[col];
        return o;
      })
    );
  };

  const renameColumn = (oldName, newName) => {
    if (!newName || newName === oldName || columns.includes(newName)) {
      setEditingHeader(null);
      return;
    }
    setColumns((prev) => prev.map((c) => (c === oldName ? newName : c)));
    setRows((prev) =>
      prev.map((r) => {
        const o = { ...r };
        o[newName] = o[oldName];
        delete o[oldName];
        return o;
      })
    );
    setEditingHeader(null);
  };

  // ---------- 匯出 ----------
  const downloadXlsx = () => {
    const ws = XLSX.utils.json_to_sheet(rows, { header: columns });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "合併結果");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `合併結果_${new Date().toISOString().slice(0, 10)}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
    showToast("已下載合併結果");
  };

  const filteredRows = useMemo(() => {
    if (!search.trim()) return rows.map((r, i) => ({ r, i }));
    const q = search.toLowerCase();
    return rows
      .map((r, i) => ({ r, i }))
      .filter(({ r }) => columns.some((c) => String(r[c]).toLowerCase().includes(q)));
  }, [rows, columns, search]);

  return (
    <div style={styles.page}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Zilla+Slab:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Noto+Sans+TC:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        ::selection { background: #B23A2E33; }
        .file-chip { transition: transform .15s ease, box-shadow .15s ease; }
        .file-chip:hover { transform: translateY(-2px); box-shadow: 0 4px 0 #C9BFA6; }
        .stamp-btn { transition: transform .18s cubic-bezier(.34,1.56,.64,1), box-shadow .18s ease; }
        .stamp-btn:hover:not(:disabled) { transform: rotate(-3deg) scale(1.03); }
        .stamp-btn:active:not(:disabled) { transform: rotate(-1deg) scale(0.97); }
        .cell-input:focus { outline: 2px solid #B23A2E; outline-offset: -2px; background: #FFFDF7; }
        .ghost-btn { transition: background .15s ease, color .15s ease; }
        .ghost-btn:hover { background: #1F2A4410; }
        .drop-zone { transition: border-color .2s ease, background .2s ease; }
        tbody tr:hover { background: #F7F2E4; }
        ::-webkit-scrollbar { height: 10px; width: 10px; }
        ::-webkit-scrollbar-thumb { background: #C9BFA6; border-radius: 6px; }
        ::-webkit-scrollbar-track { background: #F2ECDF; }
        @keyframes toastIn { from { opacity: 0; transform: translate(-50%, 8px); } to { opacity: 1; transform: translate(-50%, 0); } }
      `}</style>

      {/* Header */}
      <header style={styles.header}>
        <div style={styles.stampBadge} aria-hidden="true">
          <span style={styles.stampChar}>合</span>
        </div>
        <div>
          <h1 style={styles.title}>合併簿</h1>
          <p style={styles.subtitle}>把好幾份 Excel 疊成一本，欄位直接在網頁上點選編輯</p>
        </div>
      </header>

      {/* Upload zone */}
      {!merged && (
        <section
          className="drop-zone"
          style={{
            ...styles.dropZone,
            borderColor: dragOver ? "#B23A2E" : "#C9BFA6",
            background: dragOver ? "#F7EFE0" : "#F9F5EA",
          }}
          onDragOver={(e) => {
            e.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => {
            e.preventDefault();
            setDragOver(false);
            handleFiles(e.dataTransfer.files);
          }}
        >
          <Upload size={28} color="#8A7F66" />
          <p style={styles.dropText}>
            拖曳檔案到這裡，或
            <button
              style={styles.linkBtn}
              onClick={() => inputRef.current?.click()}
            >
              點此選擇檔案
            </button>
          </p>
          <p style={styles.dropHint}>支援 .xlsx / .xls / .csv，可一次選擇多份</p>
          <input
            ref={inputRef}
            type="file"
            multiple
            accept=".xlsx,.xls,.csv"
            style={{ display: "none" }}
            onChange={(e) => handleFiles(e.target.files)}
          />
        </section>
      )}

      {/* File chips */}
      {files.length > 0 && !merged && (
        <section style={styles.chipRow}>
          {files.map((f) => (
            <div className="file-chip" key={f.id} style={styles.chip}>
              <FileSpreadsheet size={16} color="#1F2A44" />
              <span style={styles.chipName} title={f.name}>{f.name}</span>
              {f.sheetNames.length > 1 && (
                <select
                  value={f.selectedSheet}
                  onChange={(e) => changeSheet(f.id, e.target.value)}
                  style={styles.chipSelect}
                >
                  {f.sheetNames.map((sn) => (
                    <option key={sn} value={sn}>{sn}</option>
                  ))}
                </select>
              )}
              <span style={styles.chipCount}>
                {(f.sheetsData[f.selectedSheet] || []).length} 列
              </span>
              <button style={styles.chipX} onClick={() => removeFile(f.id)} aria-label="移除檔案">
                <X size={14} />
              </button>
            </div>
          ))}
        </section>
      )}

      {/* Merge action */}
      {!merged && (
        <div style={styles.mergeRow}>
          <button
            className="stamp-btn"
            style={{ ...styles.stampButton, opacity: files.length ? 1 : 0.45 }}
            onClick={mergeAll}
            disabled={!files.length}
          >
            蓋章合併
          </button>
          <span style={styles.mergeHint}>
            {files.length > 0
              ? `準備合併 ${files.length} 份檔案`
              : "先加入檔案才能合併"}
          </span>
        </div>
      )}

      {/* Merged editable table */}
      {merged && (
        <section style={styles.tableSection}>
          <div style={styles.toolbar}>
            <div style={styles.searchBox}>
              <Search size={15} color="#8A7F66" />
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="搜尋所有欄位…"
                style={styles.searchInput}
              />
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button className="ghost-btn" style={styles.toolBtn} onClick={addRow}>
                <Plus size={14} /> 新增列
              </button>
              <button className="ghost-btn" style={styles.toolBtn} onClick={addColumn}>
                <Plus size={14} /> 新增欄
              </button>
              <button
                className="ghost-btn"
                style={{ ...styles.toolBtn, color: "#8A2E24" }}
                onClick={() => {
                  setMerged(false);
                  setFiles([]);
                  setColumns([]);
                  setRows([]);
                }}
              >
                重新開始
              </button>
              <button style={styles.downloadBtn} onClick={downloadXlsx}>
                <Download size={15} /> 下載 Excel
              </button>
            </div>
          </div>

          <div style={styles.tableWrap}>
            <table style={styles.table}>
              <thead>
                <tr>
                  <th style={styles.thIdx}>#</th>
                  {columns.map((col) => (
                    <th key={col} style={styles.th}>
                      {editingHeader === col ? (
                        <input
                          autoFocus
                          defaultValue={col}
                          onBlur={(e) => renameColumn(col, e.target.value.trim())}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") e.target.blur();
                            if (e.key === "Escape") setEditingHeader(null);
                          }}
                          style={styles.thInput}
                        />
                      ) : (
                        <div style={styles.thContent}>
                          <span>{col}</span>
                          <span style={styles.thIcons}>
                            <Pencil
                              size={12}
                              style={{ cursor: "pointer" }}
                              onClick={() => setEditingHeader(col)}
                            />
                            {col !== "來源檔案" && (
                              <Trash2
                                size={12}
                                style={{ cursor: "pointer" }}
                                onClick={() => removeColumn(col)}
                              />
                            )}
                          </span>
                        </div>
                      )}
                    </th>
                  ))}
                  <th style={styles.thIdx}></th>
                </tr>
              </thead>
              <tbody>
                {filteredRows.map(({ r, i }) => (
                  <tr key={i}>
                    <td style={styles.tdIdx}>{i + 1}</td>
                    {columns.map((col) => (
                      <td key={col} style={styles.td}>
                        <input
                          className="cell-input"
                          value={r[col]}
                          onChange={(e) => updateCell(i, col, e.target.value)}
                          style={styles.cellInput}
                        />
                      </td>
                    ))}
                    <td style={styles.tdIdx}>
                      <button
                        style={styles.rowDelBtn}
                        onClick={() => removeRow(i)}
                        aria-label="刪除此列"
                      >
                        <Trash2 size={13} />
                      </button>
                    </td>
                  </tr>
                ))}
                {filteredRows.length === 0 && (
                  <tr>
                    <td colSpan={columns.length + 2} style={styles.emptyState}>
                      找不到符合的資料
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
          <p style={styles.footNote}>
            共 {rows.length} 筆資料・{columns.length} 個欄位
            {search.trim() && ` ・符合搜尋：${filteredRows.length} 筆`}
          </p>
        </section>
      )}

      {toast && <div style={styles.toast}>{toast}</div>}
    </div>
  );
}

const styles = {
  page: {
    minHeight: "100vh",
    background:
      "repeating-linear-gradient(#F2ECDF, #F2ECDF 27px, #E7DFC9 28px), #F2ECDF",
    fontFamily: "'Noto Sans TC', sans-serif",
    color: "#1F2A44",
    padding: "32px 20px 60px",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: 18,
    maxWidth: 1040,
    margin: "0 auto 28px",
  },
  stampBadge: {
    width: 62,
    height: 62,
    borderRadius: "50%",
    border: "2.5px dashed #B23A2E",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    transform: "rotate(-8deg)",
    flexShrink: 0,
    mixBlendMode: "multiply",
  },
  stampChar: {
    fontFamily: "'Zilla Slab', serif",
    fontWeight: 700,
    fontSize: 26,
    color: "#B23A2E",
  },
  title: {
    fontFamily: "'Zilla Slab', serif",
    fontSize: 30,
    fontWeight: 700,
    margin: 0,
    letterSpacing: "0.5px",
  },
  subtitle: {
    margin: "4px 0 0",
    color: "#6B6355",
    fontSize: 14.5,
  },
  dropZone: {
    maxWidth: 1040,
    margin: "0 auto 20px",
    border: "2px dashed #C9BFA6",
    borderRadius: 10,
    padding: "38px 20px",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: 6,
    textAlign: "center",
  },
  dropText: { margin: "6px 0 0", fontSize: 15, color: "#3A3323" },
  dropHint: { margin: 0, fontSize: 12.5, color: "#8A7F66" },
  linkBtn: {
    background: "none",
    border: "none",
    color: "#B23A2E",
    fontWeight: 600,
    fontSize: 15,
    cursor: "pointer",
    textDecoration: "underline",
    padding: "0 4px",
  },
  chipRow: {
    maxWidth: 1040,
    margin: "0 auto 18px",
    display: "flex",
    flexWrap: "wrap",
    gap: 10,
  },
  chip: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    background: "#FFFDF7",
    border: "1.5px solid #1F2A44",
    borderRadius: 8,
    padding: "8px 10px",
    boxShadow: "3px 3px 0 #C9BFA6",
    fontSize: 13,
    maxWidth: 280,
  },
  chipName: {
    fontFamily: "'IBM Plex Mono', monospace",
    fontWeight: 500,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    maxWidth: 110,
  },
  chipSelect: {
    fontSize: 12,
    border: "1px solid #C9BFA6",
    borderRadius: 4,
    background: "#F2ECDF",
    color: "#1F2A44",
  },
  chipCount: { color: "#8A7F66", fontSize: 11.5, whiteSpace: "nowrap" },
  chipX: {
    marginLeft: 2,
    border: "none",
    background: "none",
    color: "#8A7F66",
    cursor: "pointer",
    display: "flex",
  },
  mergeRow: {
    maxWidth: 1040,
    margin: "0 auto",
    display: "flex",
    alignItems: "center",
    gap: 16,
  },
  stampButton: {
    fontFamily: "'Zilla Slab', serif",
    fontWeight: 700,
    fontSize: 18,
    letterSpacing: "2px",
    color: "#FFFDF7",
    background: "#B23A2E",
    border: "2px solid #1F2A44",
    borderRadius: "50% 8% 50% 8% / 8% 50% 8% 50%",
    padding: "16px 34px",
    cursor: "pointer",
    boxShadow: "4px 4px 0 #1F2A44",
  },
  mergeHint: { color: "#6B6355", fontSize: 13.5 },
  tableSection: { maxWidth: 1240, margin: "0 auto" },
  toolbar: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    flexWrap: "wrap",
    gap: 10,
    marginBottom: 14,
  },
  searchBox: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    background: "#FFFDF7",
    border: "1.5px solid #C9BFA6",
    borderRadius: 8,
    padding: "7px 12px",
    minWidth: 220,
  },
  searchInput: {
    border: "none",
    outline: "none",
    background: "transparent",
    fontSize: 13.5,
    fontFamily: "'IBM Plex Mono', monospace",
    width: "100%",
    color: "#1F2A44",
  },
  toolBtn: {
    display: "flex",
    alignItems: "center",
    gap: 5,
    border: "1.5px solid #1F2A44",
    background: "#FFFDF7",
    borderRadius: 7,
    padding: "7px 12px",
    fontSize: 13,
    cursor: "pointer",
    color: "#1F2A44",
    fontWeight: 500,
  },
  downloadBtn: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    border: "2px solid #1F2A44",
    background: "#1F2A44",
    color: "#FFFDF7",
    borderRadius: 7,
    padding: "7px 16px",
    fontSize: 13.5,
    fontWeight: 600,
    cursor: "pointer",
    boxShadow: "3px 3px 0 #B23A2E",
  },
  tableWrap: {
    background: "#FFFDF7",
    border: "1.5px solid #1F2A44",
    borderRadius: 10,
    overflow: "auto",
    maxHeight: "62vh",
  },
  table: {
    borderCollapse: "collapse",
    width: "100%",
    fontSize: 13.5,
  },
  thIdx: {
    position: "sticky",
    top: 0,
    background: "#1F2A44",
    color: "#F2ECDF",
    padding: "8px 6px",
    fontSize: 11.5,
    minWidth: 34,
    zIndex: 2,
  },
  th: {
    position: "sticky",
    top: 0,
    background: "#1F2A44",
    color: "#F2ECDF",
    padding: "9px 10px",
    textAlign: "left",
    fontWeight: 600,
    fontFamily: "'IBM Plex Mono', monospace",
    minWidth: 140,
    borderRight: "1px solid #3A4666",
    zIndex: 2,
  },
  thContent: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 8,
  },
  thIcons: { display: "flex", gap: 6, opacity: 0.75 },
  thInput: {
    width: "100%",
    background: "#F2ECDF",
    color: "#1F2A44",
    border: "none",
    borderRadius: 4,
    padding: "3px 6px",
    fontFamily: "'IBM Plex Mono', monospace",
    fontSize: 13,
  },
  td: {
    padding: 0,
    borderBottom: "1px solid #E7DFC9",
    borderRight: "1px solid #EFE8D6",
  },
  tdIdx: {
    padding: "6px 8px",
    borderBottom: "1px solid #E7DFC9",
    color: "#8A7F66",
    fontSize: 11.5,
    textAlign: "center",
    fontFamily: "'IBM Plex Mono', monospace",
  },
  cellInput: {
    width: "100%",
    border: "none",
    background: "transparent",
    padding: "8px 10px",
    fontSize: 13.5,
    fontFamily: "'IBM Plex Mono', monospace",
    color: "#1F2A44",
  },
  rowDelBtn: {
    border: "none",
    background: "none",
    color: "#B23A2E",
    cursor: "pointer",
    display: "flex",
    margin: "0 auto",
    opacity: 0.7,
  },
  emptyState: {
    textAlign: "center",
    padding: "28px 0",
    color: "#8A7F66",
    fontSize: 13.5,
  },
  footNote: {
    marginTop: 10,
    fontSize: 12.5,
    color: "#8A7F66",
    fontFamily: "'IBM Plex Mono', monospace",
  },
  toast: {
    position: "fixed",
    bottom: 26,
    left: "50%",
    transform: "translateX(-50%)",
    background: "#1F2A44",
    color: "#F2ECDF",
    padding: "10px 20px",
    borderRadius: 8,
    fontSize: 13.5,
    boxShadow: "0 6px 18px rgba(0,0,0,0.25)",
    animation: "toastIn .25s ease",
    zIndex: 50,
  },
};
