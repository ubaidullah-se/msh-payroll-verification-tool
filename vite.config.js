import { defineConfig } from "vite";
import { viteSingleFile } from "vite-plugin-singlefile";
import ViteObfuscator from "vite-plugin-javascript-obfuscator";

export default defineConfig({
  root: "src",
  plugins: [
    viteSingleFile(),
    ViteObfuscator({
      compact: true,
      stringArray: true,
      stringArrayEncoding: ["rc4"],
      controlFlowFlattening: false, // safer
      disableConsoleOutput: true,
      deadCodeInjection: true,
      debugProtection: true,
    }),
  ],
  build: {
    outDir: "../dist",
    emptyOutDir: true,
  },
});
