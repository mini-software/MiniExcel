import { cpSync, mkdirSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const here = dirname(fileURLToPath(import.meta.url));
const demoRoot = resolve(here, "..");
const repoRoot = resolve(demoRoot, "..", "..");
const dist = join(demoRoot, "dist");
const publicDir = join(demoRoot, "public");
const wasmInput = join(
  repoRoot,
  "rust",
  "target",
  "wasm32-unknown-unknown",
  "release",
  "miniexcel_wasm.wasm",
);
const toolchain = process.env.MINIEXCEL_RUST_TOOLCHAIN ?? "+1.85.0";

run("cargo", [
  toolchain,
  "build",
  "--manifest-path",
  "rust/Cargo.toml",
  "-p",
  "miniexcel-wasm",
  "--target",
  "wasm32-unknown-unknown",
  "--release",
  "--locked",
]);

mkdirSync(dist, { recursive: true });
for (const entry of readdirSync(dist)) {
  rmSync(join(dist, entry), {
    recursive: true,
    force: true,
    maxRetries: 3,
    retryDelay: 100,
  });
}
cpSync(publicDir, dist, { recursive: true });

run("wasm-bindgen", [
  wasmInput,
  "--out-dir",
  join(dist, "pkg"),
  "--target",
  "web",
  "--no-typescript",
]);

writeFileSync(join(dist, ".nojekyll"), "");
console.log(`Built browser demo: ${dist}`);

function run(command, args) {
  const result = spawnSync(command, args, {
    cwd: repoRoot,
    stdio: "inherit",
    shell: process.platform === "win32",
  });
  if (result.error?.code === "ENOENT") {
    throw new Error(
      `${command} was not found. Install wasm-bindgen-cli 0.2.127 and ensure it is on PATH.`,
    );
  }
  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}
