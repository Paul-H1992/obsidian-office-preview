import * as esbuild from "esbuild";
import builtinModules from "builtin-modules";

const isProd = process.argv[2] === "production";

const ctx = await esbuild.context({
  entryPoints: ["main.ts"],
  bundle: true,
  write: true,
  logLevel: "info",
  plugins: [
    {
      name: "obsidian-import",
      setup(build) {
        build.onResolve({ filter: /^obsidian$/ }, () => ({
          path: "obsidian",
          namespace: "obsidian-shim",
        }));
        build.onLoad({ filter: /.*/, namespace: "obsidian-shim" }, async () => ({
          contents: `module.exports = require("obsidian");`,
          loader: "js",
        }));
      },
    },
  ],
  external: [
    "obsidian",
    "electron",
    ...builtinModules,
    ...builtinModules.map((m) => `node:${m}`),
  ],
  format: "cjs",
  target: "es2020",
  platform: "node",
  sourcemap: !isProd,
  minify: isProd,
  outfile: "main.js",
});

if (isProd) {
  await ctx.rebuild();
  await ctx.dispose();
} else {
  await ctx.watch();
  console.log("Watching for changes...");
}
