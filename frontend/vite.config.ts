import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

export default defineConfig({
  plugins: [react(), tailwindcss()],
  server: {
    // During `npm run dev`, proxy API calls to the FastAPI backend on :8000.
    proxy: {
      "/api": "http://localhost:8000",
    },
  },
});
