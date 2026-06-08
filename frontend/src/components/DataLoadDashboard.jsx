import { useState } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
  CheckCircle2, AlertCircle, Loader2, FileText,
  Upload, Check, AlertTriangle, Database, Layers, Activity, X
} from "lucide-react";
import { useEntityStatus } from "../hooks/useEntityStatus";
import { useSSEAction } from "../hooks/useSSEAction";
import { API_ENDPOINTS } from "../config/api";
import { ENTITIES } from "../config/entities";

const statCardStyle = {
  background: "rgba(10,34,54,0.75)",
  backdropFilter: "blur(16px)",
  border: "1px solid rgba(34,211,238,0.12)",
  borderRadius: 12,
  padding: "16px 20px",
  boxShadow: "0 4px 24px rgba(0,0,0,0.35)",
};

export default function DataLoadDashboard() {
  const { status, loading: statusLoading, refresh } = useEntityStatus();
  const { run: runSSE } = useSSEAction();

  const [selected, setSelected]   = useState(new Set());
  const [validating, setValidating] = useState(false);
  const [loading, setLoading]     = useState(false);
  const [results, setResults]     = useState({});
  const [validated, setValidated] = useState(false);
  const [progress, setProgress]   = useState({});
  const [summary, setSummary]     = useState(null);
  const [hovered, setHovered]     = useState(null);

  const toggleEntity = (id) => {
    const n = new Set(selected);
    n.has(id) ? n.delete(id) : n.add(id);
    setSelected(n);
    setValidated(false);
    setResults({});
  };

  const handleValidate = async () => {
    if (selected.size === 0) return;
    setValidating(true); setValidated(false); setResults({}); setProgress({});
    await runSSE(API_ENDPOINTS.VALIDATE, { entities: Array.from(selected) }, {
      onEvent: (evt) => {
        if (evt.type === "progress")      setProgress(p => ({ ...p, [evt.entity]: evt.file }));
        else if (evt.type === "entity_result") setResults(r => ({ ...r, [evt.entity]: evt }));
      },
      onDone:  () => { setValidating(false); setValidated(true); refresh(); },
      onError: (err) => { console.error(err); setValidating(false); },
    });
  };

  const handleLoad = async () => {
    if (selected.size === 0) return;
    setLoading(true); setProgress({}); setSummary(null);
    const s = {};
    await runSSE(API_ENDPOINTS.LOAD, { entities: Array.from(selected) }, {
      onEvent: (evt) => {
        if (evt.type === "progress")    setProgress(p => ({ ...p, [evt.entity]: evt.file }));
        else if (evt.type === "file_result") { s[evt.entity] = s[evt.entity] || []; s[evt.entity].push(evt); }
      },
      onDone: () => {
        setLoading(false); setSummary(s); refresh();
        setSelected(new Set()); setValidated(false); setResults({});
      },
      onError: (err) => { console.error(err); setLoading(false); },
    });
  };

  const allGreen = validated
    && selected.size > 0
    && Array.from(selected).every(e => results[e]?.ok === true)
    && Array.from(selected).every(e => status[e]?.count > 0);

  const totalFiles = Object.values(status).reduce((s, e) => s + (e?.count || 0), 0);

  const EntityIcon = ({ iconClass }) => (
    <i className={`ti ${iconClass}`} style={{ fontSize: 17 }} />
  );

  return (
    <div style={{ maxWidth: 1100, margin: "0 auto", padding: "36px 24px 80px" }}>

      {/* ── Header ── */}
      <motion.div
        initial={{ opacity: 0, y: -16 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5 }}
        style={{ marginBottom: 36 }}
      >
        <p style={{ fontFamily: "'Space Mono',monospace", fontSize: "0.68rem", color: "#22d3ee", letterSpacing: "0.15em", textTransform: "uppercase", marginBottom: 6 }}>
          ◈ Data Management Console
        </p>
        <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: "clamp(1.6rem,3.5vw,2.4rem)", color: "#e2f4f8", lineHeight: 1.2, marginBottom: 8 }}>
          Load Your <span style={{ color: "#22d3ee", textShadow: "0 0 24px rgba(34,211,238,0.5)" }}>Business Data</span>
        </h1>
        <p style={{ color: "rgba(178,230,240,0.55)", fontSize: "0.9rem" }}>
          Select entities, validate integrity, then load into QAD with real-time feedback.
        </p>
      </motion.div>

      {/* ── Stats Row ── */}
      <motion.div
        initial={{ opacity: 0, y: 16 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5, delay: 0.1 }}
        style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12, marginBottom: 32 }}
      >
        {[
          { label: "Available Templates", value: ENTITIES.length, icon: <FileText size={16}/>, color: "#5eead4" },
          { label: "Selected",            value: selected.size,   icon: <Layers size={16}/>,   color: "#67e8f9" },
          {
            label: "Status",
            value: validated ? (allGreen ? "Ready" : "Review") : "Idle",
            icon: <Activity size={16}/>,
            color: validated ? (allGreen ? "#4ade80" : "#fbbf24") : "rgba(34,211,238,0.45)"
          },
        ].map((s, i) => (
          <motion.div
            key={s.label}
            initial={{ opacity: 0, y: 12 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.1 + i * 0.06 }}
            whileHover={{ y: -3, borderColor: "rgba(34,211,238,0.3)" }}
            style={{ ...statCardStyle, cursor: "default" }}
          >
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
              <span style={{ fontFamily: "'Space Mono',monospace", fontSize: "0.65rem", color: "rgba(178,230,240,0.5)", textTransform: "uppercase", letterSpacing: "0.08em" }}>{s.label}</span>
              <span style={{ color: s.color, opacity: 0.65 }}>{s.icon}</span>
            </div>
            <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: "1.7rem", color: s.color, lineHeight: 1 }}>
              {s.value}
            </div>
          </motion.div>
        ))}
      </motion.div>

      {/* ── Entity Grid ── */}
      <motion.div
        initial={{ opacity: 0, y: 16 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5, delay: 0.2 }}
        style={{ marginBottom: 28 }}
      >
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
          <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: "1rem", color: "#e2f4f8" }}>
            Select Templates
          </h2>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {selected.size > 0 && (
              <motion.button
                initial={{ opacity: 0, scale: 0.9 }}
                animate={{ opacity: 1, scale: 1 }}
                onClick={() => { setSelected(new Set()); setValidated(false); setResults({}); }}
                style={{ background: "transparent", border: "1px solid rgba(34,211,238,0.2)", color: "rgba(34,211,238,0.65)", borderRadius: 7, padding: "3px 10px", fontSize: "0.75rem", cursor: "pointer", fontFamily: "'Syne',sans-serif", display: "flex", alignItems: "center", gap: 4 }}
              >
                <X size={11} /> Clear all
              </motion.button>
            )}
            <span style={{ fontFamily: "'Space Mono',monospace", fontSize: "0.68rem", color: "rgba(178,230,240,0.4)" }}>
              {selected.size}/{ENTITIES.length} templates selected
            </span>
          </div>
        </div>

        {/* Compact 3-column grid */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 10 }}>
          {ENTITIES.map((entity, i) => {
            const isSelected = selected.has(entity.id);
            const result     = results[entity.id];
            const fileCount  = status[entity.id]?.count ?? 0;
            const inProg     = progress[entity.id];
            const isHovered  = hovered === entity.id;

            const hasFiles   = fileCount > 0;
            const fileBadgeColor = hasFiles
              ? isSelected ? "rgba(34,211,238,0.9)" : "rgba(34,211,238,0.65)"
              : "rgba(178,230,240,0.3)";
            const fileBadgeBg = hasFiles
              ? isSelected ? "rgba(34,211,238,0.15)" : "rgba(34,211,238,0.08)"
              : "rgba(255,255,255,0.04)";

            // Determine glow state after validate/load
            const noFiles = isSelected && validated && fileCount === 0;
            const hasError = result && !result.ok;
            const hasSuccess = result && result.ok && fileCount > 0;

            const glowBorder = hasError
              ? "1.5px solid rgba(239,68,68,0.75)"
              : noFiles
                ? "1.5px solid rgba(245,158,11,0.75)"
                : hasSuccess
                  ? "1.5px solid rgba(74,222,128,0.75)"
                  : isSelected
                    ? "1px solid rgba(34,211,238,0.45)"
                    : "1px solid rgba(34,211,238,0.1)";

            const glowShadow = hasError
              ? "0 0 18px rgba(239,68,68,0.4), 0 0 6px rgba(239,68,68,0.25), 0 4px 20px rgba(0,0,0,0.3)"
              : noFiles
                ? "0 0 18px rgba(245,158,11,0.4), 0 0 6px rgba(245,158,11,0.25), 0 4px 20px rgba(0,0,0,0.3)"
                : hasSuccess
                  ? "0 0 18px rgba(74,222,128,0.4), 0 0 6px rgba(74,222,128,0.25), 0 4px 20px rgba(0,0,0,0.3)"
                  : isSelected
                    ? "0 0 18px rgba(34,211,238,0.12), 0 4px 20px rgba(0,0,0,0.3)"
                    : isHovered
                      ? "0 4px 20px rgba(0,0,0,0.4), 0 0 10px rgba(34,211,238,0.06)"
                      : "0 2px 12px rgba(0,0,0,0.25)";

            const glowBg = hasError
              ? "rgba(239,68,68,0.06)"
              : noFiles
                ? "rgba(245,158,11,0.06)"
                : hasSuccess
                  ? "rgba(74,222,128,0.06)"
                  : isSelected
                    ? "rgba(34,211,238,0.07)"
                    : isHovered
                      ? "rgba(10,34,54,0.85)"
                      : "rgba(10,34,54,0.65)";

            return (
              <motion.button
                key={entity.id}
                onClick={() => toggleEntity(entity.id)}
                onHoverStart={() => setHovered(entity.id)}
                onHoverEnd={() => setHovered(null)}
                initial={{ opacity: 0, y: 12 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: 0.22 + i * 0.035 }}
                whileHover={{ y: -2, scale: 1.01 }}
                whileTap={{ scale: 0.98 }}
                style={{
                  background: glowBg,
                  backdropFilter: "blur(16px)",
                  border: glowBorder,
                  borderRadius: 12,
                  padding: "12px 14px",
                  textAlign: "left",
                  cursor: "pointer",
                  boxShadow: glowShadow,
                  transition: "background 0.2s, border 0.2s, box-shadow 0.2s",
                  position: "relative",
                  overflow: "hidden",
                }}
              >
                {/* Scan line when processing */}
                {inProg && <div className="scan-line" />}

                {/* Subtle left accent line when selected */}
                {isSelected && (
                  <div style={{
                    position: "absolute", left: 0, top: "15%", bottom: "15%",
                    width: 3, borderRadius: "0 3px 3px 0",
                    background: "linear-gradient(180deg, #22d3ee, #0d9488)",
                    boxShadow: "0 0 8px rgba(34,211,238,0.6)",
                  }} />
                )}

                <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
                  {/* Icon box */}
                  <div style={{
                    width: 36, height: 36, borderRadius: 9, flexShrink: 0,
                    background: isSelected ? "rgba(34,211,238,0.14)" : "rgba(34,211,238,0.06)",
                    border: `1px solid ${isSelected ? "rgba(34,211,238,0.35)" : "rgba(34,211,238,0.12)"}`,
                    display: "flex", alignItems: "center", justifyContent: "center",
                    color: isSelected ? "#22d3ee" : "rgba(34,211,238,0.6)",
                    transition: "all 0.2s",
                    boxShadow: isSelected ? "0 0 10px rgba(34,211,238,0.2)" : "none",
                  }}>
                    <EntityIcon iconClass={entity.icon} />
                  </div>

                  {/* Text */}
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 5, marginBottom: 2 }}>
                      <span style={{
                        fontFamily: "'Syne',sans-serif", fontWeight: 600, fontSize: "0.88rem",
                        color: isSelected ? "#22d3ee" : "#e2f4f8",
                        transition: "color 0.2s",
                        whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
                      }}>
                        {entity.id.replace(/_/g, " ")}
                      </span>
                      {/* Validation icon */}
                      {result && (
                        result.ok
                          ? <CheckCircle2 size={13} color="#4ade80" style={{ flexShrink: 0 }} />
                          : <AlertTriangle size={13} color="#f87171" style={{ flexShrink: 0 }} />
                      )}
                    </div>
                    <span style={{ fontSize: "0.72rem", color: "rgba(178,230,240,0.45)" }}>
                      {entity.desc}
                    </span>
                  </div>

                  {/* File count badge */}
                  <div style={{
                    flexShrink: 0,
                    background: fileBadgeBg,
                    border: `1px solid ${fileBadgeColor}`,
                    borderRadius: 20,
                    padding: "2px 9px",
                    fontFamily: "'Space Mono',monospace",
                    fontSize: "0.68rem",
                    color: fileBadgeColor,
                    transition: "all 0.2s",
                    whiteSpace: "nowrap",
                  }}>
                    {fileCount} {fileCount === 1 ? "file" : "files"}
                  </div>

                  {/* Checkbox */}
                  <div style={{
                    flexShrink: 0,
                    width: 18, height: 18, borderRadius: 5,
                    border: isSelected ? "none" : "1.5px solid rgba(34,211,238,0.25)",
                    background: isSelected ? "#22d3ee" : "transparent",
                    display: "flex", alignItems: "center", justifyContent: "center",
                    boxShadow: isSelected ? "0 0 8px rgba(34,211,238,0.5)" : "none",
                    transition: "all 0.2s",
                    marginLeft: 2,
                  }}>
                    {isSelected && <Check size={11} color="#020d1a" strokeWidth={3} />}
                  </div>
                </div>

                {/* Progress bar strip at bottom */}
                {inProg && (
                  <div style={{ marginTop: 8, marginLeft: 47 }}>
                    <div className="progress-bar-track">
                      <div className="progress-bar-fill" style={{ width: "60%" }} />
                    </div>
                    <p style={{ marginTop: 3, fontFamily: "'Space Mono',monospace", fontSize: "0.65rem", color: "rgba(34,211,238,0.55)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                      {inProg}
                    </p>
                  </div>
                )}
              </motion.button>
            );
          })}
        </div>
      </motion.div>

      {/* ── Action Buttons ── */}
      <motion.div
        initial={{ opacity: 0, y: 16 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5, delay: 0.35 }}
        style={{ display: "flex", gap: 12, marginBottom: 28 }}
      >
        <motion.button
          onClick={handleValidate}
          disabled={selected.size === 0 || validating}
          whileHover={selected.size > 0 && !validating ? { scale: 1.02 } : {}}
          whileTap={selected.size > 0 && !validating ? { scale: 0.98 } : {}}
          className="btn-cyan"
          style={{ flex: 1, padding: "13px 20px", fontSize: "0.88rem", display: "flex", alignItems: "center", justifyContent: "center", gap: 7 }}
        >
          {validating
            ? <><Loader2 size={15} className="spin" /> Validating…</>
            : <><FileText size={15} /> Validate Selected Templates</>}
        </motion.button>

        <motion.button
          onClick={handleLoad}
          disabled={!allGreen || loading}
          whileHover={allGreen && !loading ? { scale: 1.02 } : {}}
          whileTap={allGreen && !loading ? { scale: 0.98 } : {}}
          className="btn-green"
          style={{ flex: 1, padding: "13px 20px", fontSize: "0.88rem", display: "flex", alignItems: "center", justifyContent: "center", gap: 7 }}
        >
          {loading
            ? <><Loader2 size={15} className="spin" /> Loading Data…</>
            : <><Upload size={15} /> Load Data</>}
        </motion.button>
      </motion.div>

      {/* ── Global Progress Panel ── */}
      <AnimatePresence>
        {Object.keys(progress).length > 0 && (
          <motion.div
            initial={{ opacity: 0, height: 0 }}
            animate={{ opacity: 1, height: "auto" }}
            exit={{ opacity: 0, height: 0 }}
            style={{ marginBottom: 20, overflow: "hidden" }}
          >
            <div style={{
              background: "rgba(8,145,178,0.07)",
              border: "1px solid rgba(34,211,238,0.22)",
              borderRadius: 11, padding: "14px 18px",
            }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                <Loader2 size={14} color="#22d3ee" className="spin" />
                <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 600, fontSize: "0.82rem", color: "#22d3ee" }}>
                  Processing…
                </span>
              </div>
              <div className="progress-bar-track" style={{ marginBottom: 10 }}>
                <div className="progress-bar-fill" style={{ width: "60%" }} />
              </div>
              {Object.entries(progress).map(([entity, file]) => (
                <p key={entity} style={{ fontSize: "0.75rem", color: "rgba(178,230,240,0.6)", fontFamily: "'Space Mono',monospace", marginBottom: 2 }}>
                  <span style={{ color: "#22d3ee" }}>{entity}</span> › {file}
                </p>
              ))}
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* ── Validation Status ── */}
      <AnimatePresence>
        {validated && !allGreen && selected.size > 0 && (
          <motion.div
            initial={{ opacity: 0, y: -8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0 }}
            style={{ background: "rgba(245,158,11,0.07)", border: "1px solid rgba(245,158,11,0.28)", borderRadius: 11, padding: "12px 16px", display: "flex", alignItems: "flex-start", gap: 9, marginBottom: 20 }}
          >
            <AlertCircle size={16} color="#fbbf24" style={{ flexShrink: 0, marginTop: 1 }} />
            <div>
              <p style={{ fontWeight: 600, fontSize: "0.84rem", color: "#fcd34d", marginBottom: 2 }}>Validation Issues Found</p>
              <p style={{ fontSize: "0.76rem", color: "rgba(252,211,77,0.65)" }}>Fix errors in affected files and re-validate before loading.</p>
            </div>
          </motion.div>
        )}

        {validated && allGreen && (
          <motion.div
            initial={{ opacity: 0, y: -8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0 }}
            style={{ background: "rgba(16,185,129,0.07)", border: "1px solid rgba(16,185,129,0.28)", borderRadius: 11, padding: "12px 16px", display: "flex", alignItems: "flex-start", gap: 9, marginBottom: 20 }}
          >
            <CheckCircle2 size={16} color="#4ade80" style={{ flexShrink: 0, marginTop: 1 }} />
            <div>
              <p style={{ fontWeight: 600, fontSize: "0.84rem", color: "#4ade80", marginBottom: 2 }}>All Validations Passed</p>
              <p style={{ fontSize: "0.76rem", color: "rgba(74,222,128,0.65)" }}>Data is clean and ready. Click Load Data to proceed.</p>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* ── Load Summary ── */}
      <AnimatePresence>
        {summary && Object.keys(summary).length > 0 && (
          <motion.div initial={{ opacity: 0, y: 16 }} animate={{ opacity: 1, y: 0 }}>
            <h3 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: "1rem", color: "#e2f4f8", marginBottom: 16 }}>
              Load Summary
            </h3>
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              {Object.entries(summary).map(([entity, files], ei) => (
                <motion.div
                  key={entity}
                  initial={{ opacity: 0, y: 12 }}
                  animate={{ opacity: 1, y: 0 }}
                  transition={{ delay: ei * 0.07 }}
                  style={{
                    background: "rgba(10,34,54,0.75)",
                    backdropFilter: "blur(16px)",
                    border: "1px solid rgba(34,211,238,0.1)",
                    borderRadius: 12,
                    overflow: "hidden",
                    boxShadow: "0 4px 20px rgba(0,0,0,0.3)",
                  }}
                >
                  <div style={{ padding: "11px 18px", borderBottom: "1px solid rgba(34,211,238,0.08)", background: "rgba(34,211,238,0.03)", display: "flex", alignItems: "center", gap: 7 }}>
                    <Database size={14} color="#22d3ee" />
                    <span style={{ fontFamily: "'Syne',sans-serif", fontWeight: 600, fontSize: "0.86rem", color: "#22d3ee" }}>{entity}</span>
                  </div>
                  <div style={{ overflowX: "auto" }}>
                    <table className="data-table" style={{ width: "100%", borderCollapse: "collapse" }}>
                      <thead>
                        <tr>
                          <th style={{ textAlign: "left" }}>File</th>
                          <th style={{ textAlign: "center" }}>Added</th>
                          <th style={{ textAlign: "center" }}>Failed</th>
                          <th style={{ textAlign: "left" }}>Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {files.map((file, idx) => (
                          <tr key={idx}>
                            <td style={{ fontFamily: "'Space Mono',monospace", fontSize: "0.75rem", color: "#e2f4f8" }}>{file.file}</td>
                            <td style={{ textAlign: "center", color: "#4ade80", fontWeight: 600, fontFamily: "'Space Mono',monospace" }}>{file.ok}</td>
                            <td style={{ textAlign: "center", color: "#f87171", fontWeight: 600, fontFamily: "'Space Mono',monospace" }}>{file.fail}</td>
                            <td>
                              <span className={file.status === "archived" ? "badge-success" : "badge-warning"}>
                                {file.status === "archived"
                                  ? <><Check size={10} style={{ display:"inline", marginRight:3 }}/>{file.note}</>
                                  : <><AlertTriangle size={10} style={{ display:"inline", marginRight:3 }}/>{file.note}</>
                                }
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </motion.div>
              ))}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}