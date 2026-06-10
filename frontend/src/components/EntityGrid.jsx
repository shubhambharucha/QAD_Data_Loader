import EntityCard from "./EntityCard";
import "./EntityGrid.css";

export default function EntityGrid({ entities, selected, fileCounts, cardStates, processingFile, onToggle }) {
  return (
    <div className="entity-grid">
      {entities.map(e => (
        <EntityCard
          key={e.id}
          entity={e}
          fileCount={fileCounts[e.id] ?? 0}
          isSelected={selected.has(e.id)}
          state={cardStates[e.id] || "idle"}
          processingFile={processingFile[e.id] || null}
          onToggle={() => onToggle(e.id)}
        />
      ))}
    </div>
  );
}