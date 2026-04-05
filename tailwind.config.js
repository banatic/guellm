/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        bg: "#1c1c1e",
        surface: "#2c2c2e",
        elevated: "#3a3a3c",
        border: "rgba(255,255,255,0.08)",
        accent: {
          DEFAULT: "#0a84ff",
          hover: "#409cff",
          subtle: "rgba(10,132,255,0.12)",
        },
        text: {
          DEFAULT: "#f5f5f7",
          secondary: "#98989d",
          tertiary: "#636366",
        },
        success: "#30d158",
        error: "#ff453a",
        warning: "#ff9f0a",
      },
      fontFamily: {
        sans: [
          "-apple-system", "BlinkMacSystemFont", "Pretendard",
          "Segoe UI", "system-ui", "sans-serif",
        ],
        mono: ["SF Mono", "D2Coding", "Consolas", "monospace"],
      },
      keyframes: {
        "fade-in": {
          from: { opacity: "0", transform: "translateY(6px)" },
          to: { opacity: "1", transform: "translateY(0)" },
        },
        spin: {
          from: { transform: "rotate(0deg)" },
          to: { transform: "rotate(360deg)" },
        },
        shimmer: {
          "0%": { backgroundPosition: "-200% 0" },
          "100%": { backgroundPosition: "200% 0" },
        },
        "glow-pulse": {
          "0%, 100%": { opacity: "0.4" },
          "50%": { opacity: "0.8" },
        },
      },
      animation: {
        "fade-in": "fade-in 0.2s ease-out",
        shimmer: "shimmer 2s linear infinite",
        "glow-pulse": "glow-pulse 2s ease-in-out infinite",
      },
      boxShadow: {
        glow: "0 0 20px rgba(10, 132, 255, 0.15)",
        "glow-lg": "0 0 40px rgba(10, 132, 255, 0.2)",
        card: "0 2px 8px rgba(0, 0, 0, 0.2)",
        "card-hover": "0 4px 16px rgba(0, 0, 0, 0.3)",
      },
    },
  },
  plugins: [],
};
