import { motion } from "framer-motion";

/**
 * ActionButton
 * variant: "glow" | "disabled" | "default"
 */
export function ActionButton({ children, onClick, variant = "default", loading = false, className = "" }) {
  if (variant === "disabled") {
    return (
      <div className={[
        "w-full text-center py-3 rounded-xl text-[0.82rem] font-semibold",
        "bg-[#090f18] text-navy-600 border border-[#111d2a] cursor-not-allowed",
        "select-none tracking-wide",
        className,
      ].join(" ")}>
        {children}
      </div>
    );
  }

  const isGlow = variant === "glow";

  return (
    <motion.button
      onClick={onClick}
      disabled={loading}
      whileHover={loading ? {} : { scale: 1.02 }}
      whileTap={loading ? {} : { scale: 0.97 }}
      transition={{ duration: 0.18, ease: [0.16, 1, 0.3, 1] }}
      className={[
        "w-full py-3 rounded-xl text-[0.82rem] font-bold tracking-wide",
        "border transition-all duration-200 outline-none cursor-pointer",
        "focus-visible:ring-2 focus-visible:ring-teal/50",
        isGlow
          ? "bg-gradient-to-br from-[#0d2e22] to-[#0a2419] text-teal border-teal glow-teal hover:glow-teal-lg hover:text-[#2dffc0] disabled:opacity-60"
          : "bg-navy-800 text-navy-200 border-navy-600 hover:border-navy-500 hover:text-white disabled:opacity-60",
        className,
      ].join(" ")}
    >
      {loading ? (
        <span className="flex items-center justify-center gap-2">
          <LoadingDots />
          {children}
        </span>
      ) : children}
    </motion.button>
  );
}

function LoadingDots() {
  return (
    <span className="flex gap-1 items-center">
      {[0, 0.15, 0.3].map((d, i) => (
        <span
          key={i}
          className="w-1 h-1 rounded-full bg-teal animate-pulse-dot"
          style={{ animationDelay: `${d}s` }}
        />
      ))}
    </span>
  );
}
