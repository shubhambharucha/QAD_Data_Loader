import { AnimatePresence, motion } from "framer-motion";

/**
 * Toast — fixed bottom-center notification.
 * type: "error" | "warn" | "success"
 */
export function Toast({ message, type = "error", visible }) {
  const styles = {
    error:   "bg-[#1a0d0d] border-error/60 text-[#f08080]",
    warn:    "bg-[#1a1200] border-warn/60 text-warn",
    success: "bg-[#0c2219] border-teal/60 text-teal",
  };

  return (
    <AnimatePresence>
      {visible && (
        <motion.div
          initial={{ opacity: 0, y: 16, scale: 0.97 }}
          animate={{ opacity: 1, y: 0,  scale: 1 }}
          exit={{    opacity: 0, y: 16, scale: 0.97 }}
          transition={{ duration: 0.28, ease: [0.16, 1, 0.3, 1] }}
          className={[
            "fixed bottom-8 left-1/2 -translate-x-1/2 z-50",
            "px-6 py-3 rounded-xl border text-[0.8rem] font-medium",
            "shadow-[0_4px_24px_rgba(0,0,0,0.5)] whitespace-nowrap",
            styles[type],
          ].join(" ")}
        >
          {message}
        </motion.div>
      )}
    </AnimatePresence>
  );
}
