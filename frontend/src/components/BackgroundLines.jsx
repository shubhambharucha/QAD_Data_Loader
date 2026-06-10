import { useEffect, useRef } from "react";
import "./BackgroundLines.css";

export default function BackgroundLines({ state }) {
  return (
    <div className={`bg-lines-wrap ${state}`} aria-hidden="true">
      {Array.from({ length: 18 }).map((_, i) => (
        <div key={i} className="bg-line" style={{ animationDelay: `${i * 0.17}s` }} />
      ))}
    </div>
  );
}