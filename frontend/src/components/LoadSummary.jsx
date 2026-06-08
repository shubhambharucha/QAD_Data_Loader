import { motion } from "framer-motion";

/**
 * LoadSummary — renders grouped results table after a load completes.
 * results: { [entity]: [ { file, ok, fail, status, note } ] }
 */
export function LoadSummary({ results }) {
  if (!results || Object.keys(results).length === 0) return null;

  const totalFail = Object.values(results)
    .flat()
    .reduce((acc, r) => acc + (typeof r.fail === "number" ? r.fail : 0), 0);

  return (
    <motion.div
      initial={{ opacity: 0, y: 12 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.4, ease: [0.16, 1, 0.3, 1] }}
      className="mt-6 space-y-5"
    >
      {/* Summary banner */}
      <div className={[
        "flex items-center gap-3 px-4 py-3 rounded-xl border text-sm font-semibold",
        totalFail === 0
          ? "bg-[#0c2219] border-teal/40 text-teal"
          : "bg-[#1a0d0d] border-error/40 text-error",
      ].join(" ")}>
        <i className={`ti ${totalFail === 0 ? "ti-circle-check" : "ti-alert-triangle"} text-base`} />
        {totalFail === 0
          ? "Load complete — all records loaded successfully."
          : `Load finished with ${totalFail} failure(s). Fix red rows and re-run.`}
      </div>

      {/* Per-entity tables */}
      {Object.entries(results).map(([entity, rows]) => (
        <div key={entity}>
          <p className="text-[0.72rem] font-bold text-teal uppercase tracking-widest mb-2">{entity}</p>
          <div className="rounded-xl border border-navy-700 overflow-hidden">
            <table className="w-full text-[0.78rem]">
              <thead>
                <tr className="bg-navy-950 border-b border-navy-700">
                  <th className="text-left px-4 py-2.5 text-navy-400 font-semibold uppercase text-[0.68rem] tracking-wider">File</th>
                  <th className="text-center px-3 py-2.5 text-navy-400 font-semibold uppercase text-[0.68rem] tracking-wider">Added</th>
                  <th className="text-center px-3 py-2.5 text-navy-400 font-semibold uppercase text-[0.68rem] tracking-wider">Failed</th>
                  <th className="text-left px-4 py-2.5 text-navy-400 font-semibold uppercase text-[0.68rem] tracking-wider">Status</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => (
                  <tr key={i} className="border-b border-navy-800 last:border-0 bg-navy-900">
                    <td className="px-4 py-2.5 font-mono text-navy-200 text-[0.73rem]">{r.file}</td>
                    <td className="px-3 py-2.5 text-center text-navy-200">{r.ok}</td>
                    <td className="px-3 py-2.5 text-center text-navy-200">{r.fail}</td>
                    <td className={[
                      "px-4 py-2.5 font-semibold",
                      r.status === "archived" ? "text-teal"  :
                      r.status === "partial"  ? "text-warn"  : "text-error",
                    ].join(" ")}>
                      {r.note}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ))}
    </motion.div>
  );
}
