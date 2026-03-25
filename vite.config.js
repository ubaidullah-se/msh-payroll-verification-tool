import { defineConfig } from "vite";
import { viteSingleFile } from "vite-plugin-singlefile";
import ViteObfuscator from "vite-plugin-javascript-obfuscator";

export default defineConfig({
  root: "src",
  plugins: [
    viteSingleFile(),
    ViteObfuscator({
      // obfuscate JS
      options: {
        compact: true,
        controlFlowFlattening: true,
        deadCodeInjection: true,
        stringArrayEncoding: ["rc4"],
        disableConsoleOutput: true,
        debugProtection: true,
      },
    }),
  ],
  build: {
    outDir: "../dist",
    emptyOutDir: true,
  },
});
