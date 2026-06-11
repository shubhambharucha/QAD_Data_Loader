import "./BackgroundLines.css";

// Bipsync-style: two clusters of vertical bars, left and right sides
// Each bar has independent height + opacity animation (staggered)
const LEFT_BARS  = [0.3, 0.55, 0.75, 0.9, 1, 0.85, 0.65, 0.4, 0.25, 0.5, 0.7, 0.35];
const RIGHT_BARS = [0.25, 0.5, 0.8, 1, 0.9, 0.7, 0.45, 0.6, 0.85, 0.4, 0.3, 0.55];

export default function BackgroundLines({ state }) {
  return (
    <div className={`bg-lines-root ${state}`} aria-hidden="true">
      {/* Left cluster */}
      <div className="bar-cluster bar-cluster--left">
        {LEFT_BARS.map((h, i) => (
          <div
            key={i}
            className="bar"
            style={{
              "--bar-h": h,
              animationDelay: `${i * 0.22}s`,
            }}
          />
        ))}
      </div>

      {/* Right cluster */}
      <div className="bar-cluster bar-cluster--right">
        {RIGHT_BARS.map((h, i) => (
          <div
            key={i}
            className="bar"
            style={{
              "--bar-h": h,
              animationDelay: `${i * 0.19}s`,
            }}
          />
        ))}
      </div>
    </div>
  );
}