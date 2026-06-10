import { useState, useEffect, useRef } from "react";
import Header from "./components/Header";
import EntityGrid from "./components/EntityGrid";
import ActionButtons from "./components/ActionButtons";
import ResultsTable from "./components/ResultsTable";
import BackgroundLines from "./components/BackgroundLines";
import "./App.css";

const ENTITIES = [
  { id: "Supplier", label: "Supplier", sub: "Vendors, contracts, sites" },
  { id: "Customer", label: "Customer", sub: "Accounts, addresses" },
  { id: "GCM", label: "GCM", sub: "Global cost masters" },
  { id: "BR", label: "BR", sub: "Business relations" },
  { id: "Customer_Item", label: "Customer Item", sub: "Customer item cross-refs" },
  { id: "ProductionOrder", label: "ProductionOrder", sub: "Production order headers" },
  { id: "PurchaseOrder", label: "PurchaseOrder", sub: "PO headers & lines" },
  { id: "SalesOrder", label: "SalesOrder", sub: "SO headers & lines" },
  { id: "Supplier_Item", label: "Supplier Item", sub: "Supplier item cross-refs" },
  { id: "SupplierPriceList", label: "SupplierPriceList", sub: "Supplier price lists" },
];

export default function App() {
  const [selected, setSelected] = useState(new Set());
  const [fileCounts, setFileCounts] = useState({});
  const [cardStates, setCardStates] = useState({});   // { [entity]: 'idle'|'processing'|'success'|'failed' }
  const [processingFile, setProcessingFile] = useState({}); // { [entity]: filename }
  const [results, setResults] = useState([]);
  const [lineState, setLineState] = useState("idle"); // 'idle'|'burst'|'fail'
  const [opInProgress, setOpInProgress] = useState(false);
  const [showResults, setShowResults] = useState(false);
  const lineTimerRef = useRef(null);

  // Poll status every 5s
  useEffect(() => {
    fetchStatus();
    const id = setInterval(fetchStatus, 5000);
    return () => clearInterval(id);
  }, []);

  async function fetchStatus() {
    try {
      const res = await fetch("http://localhost:8000/api/status");
      const data = await res.json();
      const counts = {};
      for (const [k, v] of Object.entries(data)) counts[k] = v.count;
      setFileCounts(counts);
    } catch {
      // backend not running — silently ignore
    }
  }

  function toggleEntity(id) {
    if (opInProgress) return;
    setSelected(prev => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }

  function triggerLineBurst(fail = false) {
    clearTimeout(lineTimerRef.current);
    setLineState(fail ? "fail" : "burst");
    lineTimerRef.current = setTimeout(() => setLineState("idle"), fail ? 1500 : 800);
  }

  async function runOperation(type) {
    if (selected.size === 0 || opInProgress) return;
    setOpInProgress(true);
    setShowResults(false);
    setResults([]);
    const newCardStates = {};
    selected.forEach(e => { newCardStates[e] = "processing"; });
    setCardStates(newCardStates);
    triggerLineBurst(false);

    const endpoint = type === "validate" ? "/api/validate" : "/api/load";
    let anyFailed = false;
    const newResults = [];

    try {
      const res = await fetch(`http://localhost:8000${endpoint}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ entities: [...selected] }),
      });

      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      let buffer = "";

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop();

        for (const line of lines) {
          if (!line.startsWith("data: ")) continue;
          try {
            const event = JSON.parse(line.slice(6));

            if (event.type === "progress") {
              setProcessingFile(prev => ({ ...prev, [event.entity]: event.file }));
            }

            if (event.type === "entity_result") {
              const ok = event.ok;
              if (!ok) anyFailed = true;
              setCardStates(prev => ({ ...prev, [event.entity]: ok ? "success" : "failed" }));
              setProcessingFile(prev => { const n = { ...prev }; delete n[event.entity]; return n; });
            }

            if (event.type === "file_result") {
              const ok = event.fail === 0 || event.fail === "0";
              if (!ok) anyFailed = true;
              newResults.push(event);
              setResults([...newResults]);
              setCardStates(prev => ({
                ...prev,
                [event.entity]: (!ok && prev[event.entity] !== "failed") ? "failed" : (ok && prev[event.entity] !== "failed" ? "success" : prev[event.entity]),
              }));
            }

            if (event.type === "done") {
              triggerLineBurst(anyFailed);
              if (anyFailed) setLineState("fail");
              setShowResults(true);
            }
          } catch { /* malformed event */ }
        }
      }
    } catch (err) {
      console.error(err);
      selected.forEach(e => {
        setCardStates(prev => ({ ...prev, [e]: "failed" }));
      });
      anyFailed = true;
      triggerLineBurst(true);
    }

    setOpInProgress(false);
    fetchStatus();
  }

  return (
    <div className="app-root">
      <BackgroundLines state={lineState} />
      <div className="app-content">
        <Header />
        <EntityGrid
          entities={ENTITIES}
          selected={selected}
          fileCounts={fileCounts}
          cardStates={cardStates}
          processingFile={processingFile}
          onToggle={toggleEntity}
        />
        <ActionButtons
          hasSelection={selected.size > 0}
          inProgress={opInProgress}
          onValidate={() => runOperation("validate")}
          onLoad={() => runOperation("load")}
        />
        {showResults && results.length > 0 && <ResultsTable results={results} />}
      </div>
    </div>
  );
}