import "./ResultsTable.css";

export default function ResultsTable({ results }) {
  if (!results || results.length === 0) return null;

  return (
    <div className="results-wrap">
      <div className="results-header">
        <span className="results-title">OPERATION RESULTS</span>
        <span className="results-count">{results.length} records</span>
      </div>
      <div className="results-table-wrap">
        <table className="results-table">
          <thead>
            <tr>
              <th>ENTITY</th>
              <th>FILE</th>
              <th>PASSED</th>
              <th>FAILED</th>
              <th>STATUS</th>
              <th>NOTE</th>
            </tr>
          </thead>
          <tbody>
            {results.map((r, i) => {
              const ok = r.fail === 0 || r.fail === "0" || r.fail === "?";
              return (
                <tr key={i} className={ok ? "row-ok" : "row-fail"}>
                  <td className="cell-entity">{r.entity}</td>
                  <td className="cell-file">{r.file}</td>
                  <td className="cell-ok">{r.ok}</td>
                  <td className="cell-fail">{r.fail}</td>
                  <td>
                    <span className={`status-pill ${r.status}`}>{r.status}</span>
                  </td>
                  <td className="cell-note">{r.note}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}