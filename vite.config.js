import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      // The plain "xlsx" entry point isn't browser-safe; point Vite at
      // SheetJS's bundled UMD build instead, which has no Node dependencies.
      xlsx: "xlsx/dist/xlsx.full.min.js",
    },
  },
  build: {
    rollupOptions: {
      onwarn(warning, warn) {
        if (warning.code === "MODULE_LEVEL_DIRECTIVE" || warning.message.includes("externalized")) return;
        warn(warning);
      },
    },
  },
});
