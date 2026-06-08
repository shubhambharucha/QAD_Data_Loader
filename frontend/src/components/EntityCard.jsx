import { motion } from "framer-motion";

/**
 * EntityCard
 * Props:
 *   entity       — { id, desc, icon }
 *   selected     — bool
 *   fileCount    — number
 *   valResult    — { ok, error_files } | null
 *   validated    — bool (global)
 *   onToggle     — fn()
 */
export function EntityCard({ entity, selected, fileCount, valResult, validated, onToggle }) {
  const { id, desc, icon } = entity;

  // ── Status indicator ───────────────────────────────────────────────────
  let dot = null;
  if (selected && validated && valResult) {
    if (fileCount === 0) {
      dot = <span className="w-2 h-2 rounded-full bg-warn shadow-[0_0_6px_#f0a500]" />;
    } else if (valResult.ok) {
      dot = <span className="w-2 h-2 rounded-full bg-teal shadow-[0_0_6px_#1db890] animate-pulse-dot" />;
    } else {
      dot = <span className="w-2 h-2 rounded-full bg-error shadow-[0_0_6px_#e05252]" />;
    }
  }

  // ── File count badge (shown when not in post-validate dot mode) ────────
  let badge = null;
  if (!dot) {
    badge = fileCount > 0
      ? (
        <span className="text-[0.62rem] font-bold px-2 py-0.5 rounded-full border
          bg-[#0e2e22] text-teal border-teal/20 whitespace-nowrap tracking-wide">
          {fileCount} file{fileCount !== 1 ? "s" : ""}
        </span>
      )
      : (
        <span className="text-[0.62rem] font-bold px-2 py-0.5 rounded-full border
          bg-[#1a1a2a] text-navy-400 border-navy-700/50 whitespace-nowrap">
          No files
        </span>
      );
  }

  return (
    <motion.button
      onClick={onToggle}
      whileHover={{ scale: 1.015 }}
      whileTap={{ scale: 0.985 }}
      transition={{ duration: 0.18, ease: [0.16, 1, 0.3, 1] }}
      className={[
        "relative w-full flex items-center gap-3 rounded-xl px-3 py-2.5 text-left",
        "border transition-all duration-200 cursor-pointer outline-none",
        "focus-visible:ring-2 focus-visible:ring-teal/50",
        selected
          ? "bg-[#0c2219] border-teal glow-teal"
          : "bg-navy-900 border-navy-700 hover:border-navy-500 hover:bg-navy-800",
      ].join(" ")}
      style={{ minHeight: 62 }}
    >
      {/* Icon */}
      <div className="w-9 h-9 min-w-[36px] rounded-lg bg-[#0e2e22] flex items-center justify-center">
        <i
          className={`ti ${icon} text-base text-teal`}
          style={{ filter: "drop-shadow(0 0 5px rgba(29,184,144,0.6))" }}
        />
      </div>

      {/* Text */}
      <div className="flex-1 pr-12 min-w-0">
        <p className="text-navy-100 text-[0.83rem] font-bold leading-tight truncate">{id}</p>
        <p className="text-navy-400 text-[0.68rem] leading-tight mt-0.5 truncate">{desc}</p>
      </div>

      {/* Right indicator */}
      <div className="absolute right-3 top-1/2 -translate-y-1/2 flex items-center">
        {dot ?? badge}
      </div>
    </motion.button>
  );
}
