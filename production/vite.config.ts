import { defineConfig } from "vite";
import { resolve } from "node:path";

export default defineConfig({
  build: {
    rollupOptions: {
      input: {
        main: resolve(__dirname, "index.html"),
        p42: resolve(__dirname, "p4-2.html"),
        p43: resolve(__dirname, "p4-3.html"),
      },
    },
  },
});
