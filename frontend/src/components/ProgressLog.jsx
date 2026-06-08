import { useEffect, useRef } from "react";
import { motion, AnimatePresence } from "framer-motion";

/**
 * ProgressLog
 * Renders a scrollable, auto-scrolling live feed of SSE events.
 */
export function ProgressLog({ events }) {
  const bottomRef = useRef(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [events]);

  if (events.length === 0) return null;

  return (
    <motion.div
      initial={{ opacity: 0, height: 0 }}
      animate={{ opacity: 1, height: "auto" }}
      exit={{ opacity: 0, height: 0 }}
      transition={{ duration: 0.35, ease: [0.16, 1, 0.3, 1] }}
      className="mt-5 rounded-xl border border-navy-700 bg-navy-950 overflow-hidden"
    >
      <div className="px-4 py-2.5 border-b border-navy-700 flex items-center gap-2">
        <span className="w-1.5 h-1.5 rounded-full bg-teal animate-pulse-dot" />
        <span className="text-[0.72rem] font-semibold text-navy-300 uppercase tracking-widest">
          Live Progress
        </span>
      </div>

      <div className="max-h-52 overflow-y-auto px-4 py-3 space-y-1 font-mono text-[0.74rem]">
        <AnimatePresence initial={false}>
          {events.map((evt, i) => (
            <LogLine key={i} evt={evt} />
          ))}
        </AnimatePresence>
        <div ref={bottomRef} />
      </div>
    </motion.div>
  );
}

function LogLine({ evt }) {
  const { icon, color, text } = formatEvent(evt);

  return (
    <motion.div
      initial={{ opacity: 0, x: -8 }}
      animate={{ opacity: 1, x: 0 }}
      transition={{ duration: 0.2, ease: "easeOut" }}
      className="flex items-start gap-2 leading-snug"
    >
      <span className="mt-0.5 shrink-0" style={{ color }}>{icon}</span>
      <span className="text-navy-200">{text}</span>
    </motion.div>
  );
}

function formatEvent(evt) {
  switch (evt.type) {
    case "progress":
      return {
        icon:  "›",
        color: "#5a7a95",
        text:  `${evt.entity} › ${evt.file}`,
      };
    case "entity_result":
      return evt.ok
        ? { icon: "✔", color: "#1db890", text: `${evt.entity} — validated OK` }
        : { icon: "✘", color: "#e05252", text: `${evt.entity} — ${evt.error_files?.length ?? 0} file(s) with errors` };
    case "file_result":
      if (evt.status === "archived") {
        return { icon: "✔", color: "#1db890", text: `${evt.entity} › ${evt.file} — ${evt.ok} loaded, archived` };
      }
      if (evt.status === "partial") {
        return { icon: "⚠", color: "#f0a500", text: `${evt.entity} › ${evt.file} — ${evt.ok} ok, ${evt.fail} failed` };
      }
      return { icon: "✘", color: "#e05252", text: `${evt.entity} › ${evt.file} — ${evt.note}` };
    case "error":
      return { icon: "✘", color: "#e05252", text: `${evt.entity} › ${evt.file} — ${evt.msg}` };
    default:
      return { icon: "·", color: "#3e5870", text: JSON.stringify(evt) };
  }
}
