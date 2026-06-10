import "./EntityCard.css";

const ICONS = {
  Supplier:          "M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5",
  Customer:          "M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2M12 3a4 4 0 1 0 0 8 4 4 0 0 0 0-8z",
  GCM:               "M12 2a10 10 0 1 0 0 20A10 10 0 0 0 12 2zM2 12h20M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z",
  BR:                "M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2M9 7a4 4 0 1 0 0 8 4 4 0 0 0 0-8zM23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75",
  Customer_Item:     "M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4zM3 6h18M16 10a4 4 0 0 1-8 0",
  ProductionOrder:   "M2 20h.01M7 20v-4M12 20v-8M17 20V8M22 4L12 14.01l-4-4L2 16",
  PurchaseOrder:     "M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8zM14 2v6h6M16 13H8M16 17H8M10 9H8",
  SalesOrder:        "M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16zM3.27 6.96L12 12.01l8.73-5.05M12 22.08V12",
  Supplier_Item:     "M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82zM7 7h.01",
  SupplierPriceList: "M12 1v22M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6",
};

export default function EntityCard({ entity, fileCount, isSelected, state, processingFile, onToggle }) {
  const cardClass = [
    "entity-card",
    isSelected ? "selected" : "",
    state === "success" ? "success" : "",
    state === "failed"  ? "failed"  : "",
    state === "processing" ? "processing" : "",
  ].filter(Boolean).join(" ");

  const iconPath = ICONS[entity.id] || ICONS["Supplier"];

  return (
    <div className={cardClass} onClick={onToggle} role="button" tabIndex={0}
      onKeyDown={e => e.key === "Enter" && onToggle()}>

      <div className="card-top">
        <div className="card-icon-wrap">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"
            strokeLinecap="round" strokeLinejoin="round" className="card-icon">
            <path d={iconPath} />
          </svg>
        </div>

        <div className="card-info">
          <span className="card-name">{entity.label}</span>
          <span className="card-sub">{entity.sub}</span>
        </div>

        <div className="card-right">
          <div className={`card-checkbox ${isSelected ? "checked" : ""}`}>
            {isSelected && (
              <svg viewBox="0 0 10 8" fill="none">
                <path d="M1 4l3 3 5-6" stroke="currentColor" strokeWidth="1.5"
                  strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            )}
          </div>
        </div>
      </div>

      <div className="card-bottom">
        <span className="file-count">
          <span className="file-count-dot" />
          {fileCount} {fileCount === 1 ? "file" : "files"}
        </span>

        {processingFile && (
          <span className="processing-label">
            <span className="processing-blink" />
            {processingFile}
          </span>
        )}

        {(state === "success" || state === "failed") && !processingFile && (
          <span className={`status-badge ${state}`}>
            {state === "success" ? "✓ passed" : "✗ failed"}
          </span>
        )}
      </div>
    </div>
  );
}