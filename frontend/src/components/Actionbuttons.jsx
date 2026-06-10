import { useState } from "react";
import "./ActionButtons.css";

export default function ActionButtons({ hasSelection, inProgress, onValidate, onLoad }) {
  const [bursting, setBursting] = useState(null);

  function handleClick(type, fn) {
    if (!hasSelection || inProgress) return;
    setBursting(type);
    setTimeout(() => setBursting(null), 800);
    fn();
  }

  const disabled = !hasSelection || inProgress;

  return (
    <div className="action-area">
      <div className="action-line-left" />
      <div className="action-buttons">
        <button
          className={`action-btn ${disabled ? "disabled" : ""} ${bursting === "validate" ? "bursting" : ""}`}
          onClick={() => handleClick("validate", onValidate)}
          disabled={disabled}
        >
          {inProgress ? <span className="btn-spinner" /> : null}
          VALIDATE
        </button>
        <button
          className={`action-btn ${disabled ? "disabled" : ""} ${bursting === "load" ? "bursting" : ""}`}
          onClick={() => handleClick("load", onLoad)}
          disabled={disabled}
        >
          {inProgress ? <span className="btn-spinner" /> : null}
          LOAD
        </button>
      </div>
      <div className="action-line-right" />
    </div>
  );
}