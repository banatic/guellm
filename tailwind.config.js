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
      },
      animation: {
        "fade-in": "fade-in 0.2s ease-out",
      },
    },
  },
  plugins: [],
};
