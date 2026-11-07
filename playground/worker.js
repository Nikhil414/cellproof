// cellproof in a Web Worker: Pyodide (Python in WebAssembly) + DuckDB + openpyxl + cellproof itself.
// The workbook is read from memory and converted here; nothing is sent anywhere.
importScripts("https://cdn.jsdelivr.net/pyodide/v0.27.7/full/pyodide.js");

// The newest DuckDB built for Pyodide (Python 3.12, ABI 2024_0, the one Pyodide 0.27 uses). build.sh also puts a
// copy next to this file, checked against the same SHA-256.
const DUCKDB = "duckdb-1.2.0-cp312-cp312-pyodide_2024_0_wasm32.whl";

// jsDelivr first: some networks crawl on github.io (measured 260 s for this 10 MB file, against seconds from jsDelivr).
// The bytes must match build.sh's SHA-256, else the copy next to this file is used.
const DUCKDB_MIRROR = "https://cdn.jsdelivr.net/gh/duckdb/duckdb-pyodide@main/wheels/" + DUCKDB;
const DUCKDB_SHA256 = "48ace5ee2bf26d13943e2260c3143ae3eb87c727c268f6020d5424e2fbe6d2ea";

async function duckdbWheel() {
  try {
    const bytes = new Uint8Array(await (await fetch(DUCKDB_MIRROR, { signal: AbortSignal.timeout(60000) })).arrayBuffer());
    const hex = [...new Uint8Array(await crypto.subtle.digest("SHA-256", bytes))].map((b) => b.toString(16).padStart(2, "0")).join("");
    if (hex === DUCKDB_SHA256) return bytes;
  } catch {}
  return new Uint8Array(await (await fetch(DUCKDB)).arrayBuffer());
}

const post = (type, data) => self.postMessage({ type, data });
let py;

async function boot() {
  post("status", "Loading Python (first visit only, about 15 MB)…");
  py = await loadPyodide();
  post("status", "Loading DuckDB, openpyxl and cellproof…");
  // openpyxl and cellproof are served next to this file (build.sh)
  const local = (await (await fetch("wheels.txt")).text()).trim().split("\n");
  // unpacked, not micropip-installed: micropip preloads DuckDB's native library in a way that crashed it in Chromium
  // ("null function" on the first read_csv); unpacked, Python loads it on import, as in Node, where it works
  py.unpackArchive(await duckdbWheel(), "wheel");
  for (const w of local) py.unpackArchive(new Uint8Array(await (await fetch(w)).arrayBuffer()), "wheel");
  py.FS.writeFile("/home/pyodide/cellproof_web.py", await (await fetch("cellproof_web.py")).text());
  py.runPython("import cellproof_web, duckdb, cellproof");
  post("ready", py.runPython("f'cellproof {cellproof.__version__ if hasattr(cellproof, \"__version__\") else \"\"} on DuckDB {duckdb.__version__}'"));
}

const booted = boot().catch((err) => post("fatal", String(err)));

self.onmessage = async ({ data: { name, bytes } }) => {
  await booted;
  try {
    const t0 = performance.now();
    py.FS.writeFile("/tmp/in.bin", new Uint8Array(bytes));
    py.globals.set("in_name", name);
    const json = py.runPython(`
res, z = cellproof_web.run(open('/tmp/in.bin', 'rb').read(), in_name)
open('/tmp/out.zip', 'wb').write(z)
res`);
    const zip = py.FS.readFile("/tmp/out.zip");
    post("result", { json, zip, ms: performance.now() - t0 });
  } catch (err) {
    post("error", String(err));
  }
};
