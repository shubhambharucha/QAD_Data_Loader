/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: {
    extend: {
      colors: {
        navy: {
          950: "#020d1a",
          900: "#041220",
          800: "#061828",
          700: "#0a2236",
          600: "#0e2d45",
        },
        cyan: {
          glow: "rgba(34,211,238,0.5)",
        },
      },
      fontFamily: {
        sans: ["'Syne'", "sans-serif"],
        mono: ["'Space Mono'", "monospace"],
      },
      boxShadow: {
        glow: "0 0 20px rgba(34,211,238,0.4)",
        "glow-lg": "0 0 40px rgba(34,211,238,0.6)",
        card: "0 4px 24px rgba(0,0,0,0.4)",
      },
      animation: {
        "fade-in": "fadeIn 0.5s ease-out forwards",
        "slide-up": "slideUp 0.5s ease-out forwards",
        "pulse-glow": "pulseGlow 2s ease-in-out infinite",
      },
      keyframes: {
        fadeIn: {
          "0%": { opacity: "0" },
          "100%": { opacity: "1" },
        },
        slideUp: {
          "0%": { opacity: "0", transform: "translateY(16px)" },
          "100%": { opacity: "1", transform: "translateY(0)" },
        },
        pulseGlow: {
          "0%, 100%": { boxShadow: "0 0 20px rgba(34,211,238,0.4)" },
          "50%": { boxShadow: "0 0 40px rgba(34,211,238,0.7)" },
        },
      },
    },
  },
  plugins: [],
};
