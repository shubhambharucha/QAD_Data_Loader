import { useEffect, useRef } from "react";
import { motion } from "framer-motion";
import DataLoadDashboard from "./components/DataLoadDashboard";
import "./index.css";

/* ── Constellation canvas ── */
function ConstellationBg() {
  const canvasRef = useRef(null);

  useEffect(() => {
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");
    let raf;

    const resize = () => {
      canvas.width = window.innerWidth;
      canvas.height = window.innerHeight;
    };
    resize();
    window.addEventListener("resize", resize);

    const NUM = 90;
    const dots = Array.from({ length: NUM }, () => ({
      x: Math.random() * canvas.width,
      y: Math.random() * canvas.height,
      vx: (Math.random() - 0.5) * 0.25,
      vy: (Math.random() - 0.5) * 0.25,
      r: Math.random() * 1.5 + 0.5,
      alpha: Math.random() * 0.5 + 0.3,
    }));

    const draw = () => {
      ctx.clearRect(0, 0, canvas.width, canvas.height);

      // move
      dots.forEach(d => {
        d.x += d.vx;
        d.y += d.vy;
        if (d.x < 0) d.x = canvas.width;
        if (d.x > canvas.width) d.x = 0;
        if (d.y < 0) d.y = canvas.height;
        if (d.y > canvas.height) d.y = 0;
      });

      // lines
      for (let i = 0; i < dots.length; i++) {
        for (let j = i + 1; j < dots.length; j++) {
          const dx = dots[i].x - dots[j].x;
          const dy = dots[i].y - dots[j].y;
          const dist = Math.sqrt(dx * dx + dy * dy);
          if (dist < 130) {
            ctx.beginPath();
            ctx.moveTo(dots[i].x, dots[i].y);
            ctx.lineTo(dots[j].x, dots[j].y);
            ctx.strokeStyle = `rgba(34,211,238,${(1 - dist / 130) * 0.18})`;
            ctx.lineWidth = 0.6;
            ctx.stroke();
          }
        }
      }

      // dots
      dots.forEach(d => {
        ctx.beginPath();
        ctx.arc(d.x, d.y, d.r, 0, Math.PI * 2);
        ctx.fillStyle = `rgba(34,211,238,${d.alpha})`;
        ctx.fill();
      });

      raf = requestAnimationFrame(draw);
    };
    draw();

    return () => {
      cancelAnimationFrame(raf);
      window.removeEventListener("resize", resize);
    };
  }, []);

  return <canvas ref={canvasRef} style={{ position: "fixed", inset: 0, zIndex: 0, pointerEvents: "none" }} />;
}

export default function App() {
  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #020d1a 0%, #041220 50%, #020d1a 100%)", position: "relative" }}>
      <ConstellationBg />

      {/* Ambient glowing orbs */}
      <div style={{ position: "fixed", inset: 0, zIndex: 0, pointerEvents: "none", overflow: "hidden" }}>
        <motion.div
          style={{ position: "absolute", top: "-10%", left: "15%", width: 500, height: 500, borderRadius: "50%", background: "radial-gradient(circle, rgba(8,145,178,0.12) 0%, transparent 70%)", filter: "blur(40px)" }}
          animate={{ scale: [1, 1.15, 1], opacity: [0.6, 0.9, 0.6] }}
          transition={{ duration: 18, repeat: Infinity, ease: "easeInOut" }}
        />
        <motion.div
          style={{ position: "absolute", bottom: "5%", right: "10%", width: 400, height: 400, borderRadius: "50%", background: "radial-gradient(circle, rgba(13,148,136,0.1) 0%, transparent 70%)", filter: "blur(40px)" }}
          animate={{ scale: [1, 1.2, 1], opacity: [0.5, 0.8, 0.5] }}
          transition={{ duration: 22, repeat: Infinity, ease: "easeInOut", delay: 4 }}
        />
        <motion.div
          style={{ position: "absolute", top: "40%", right: "25%", width: 300, height: 300, borderRadius: "50%", background: "radial-gradient(circle, rgba(34,211,238,0.07) 0%, transparent 70%)", filter: "blur(30px)" }}
          animate={{ scale: [1, 1.3, 1], opacity: [0.4, 0.7, 0.4] }}
          transition={{ duration: 14, repeat: Infinity, ease: "easeInOut", delay: 2 }}
        />
      </div>

      {/* Nav */}
      <nav style={{ position: "sticky", top: 0, zIndex: 50, borderBottom: "1px solid rgba(34,211,238,0.12)", background: "rgba(2,13,26,0.85)", backdropFilter: "blur(20px)" }}>
        <div style={{ maxWidth: 1280, margin: "0 auto", padding: "0 24px", height: 64, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            {/* Logo — replace src with your actual image path, e.g. "/logo.png" */}
            <div style={{ width: 38, height: 38, borderRadius: 10, overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 0 16px rgba(34,211,238,0.4)" }}>
              <img
                src="/yash-logo.png"
                alt="YASH Logo"
                style={{ width: "100%", height: "100%", objectFit: "contain" }}
                onError={e => {
                  // Fallback to initials if image not found
                  e.target.style.display = "none";
                  e.target.parentNode.style.background = "linear-gradient(135deg, #0891b2, #0d9488)";
                  e.target.parentNode.innerHTML = '<span style="font-family:Syne,sans-serif;font-weight:700;font-size:0.75rem;color:white;letter-spacing:0.05em">YT</span>';
                }}
              />
            </div>
            <div>
              <span style={{ fontFamily: "'Syne', sans-serif", fontWeight: 700, fontSize: "1.1rem", color: "#e2f4f8", letterSpacing: "0.03em" }}>
                YASH QAD <span style={{ color: "#22d3ee" }}>DataLoader</span>
              </span>
            </div>
          </div>

          {/* Nav right */}
          <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
            <span style={{ fontFamily: "'Space Mono', monospace", fontSize: "0.7rem", color: "rgba(34,211,238,0.6)", letterSpacing: "0.1em", textTransform: "uppercase" }}>
              Enterprise Edition
            </span>
            <div style={{ width: 8, height: 8, borderRadius: "50%", background: "#22d3ee", boxShadow: "0 0 8px rgba(34,211,238,0.8)", animation: "pulse 2s ease-in-out infinite" }} />
          </div>
        </div>
      </nav>

      {/* Main content */}
      <motion.div
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        transition={{ duration: 0.6 }}
        style={{ position: "relative", zIndex: 10 }}
      >
        <DataLoadDashboard />
      </motion.div>
    </div>
  );
}
