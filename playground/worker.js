// cellproof in a Web Worker: Pyodide (Python in WebAssembly) + DuckDB + openpyxl + cellproof itself.
// The workbook is read from memory and converted here; nothing is sent anywhere.
importScripts("https://cdn.jsdelivr.net/pyodide/v0.27.7/full/pyodide.js");

// The newest DuckDB built for Pyodide (Python 3.12, ABI 2024_0, the one Pyodide 0.27 uses). build.sh puts it next
// to this file, checked against its SHA-256: the page then depends on one host for its Python, not two.
const DUCKDB = "duckdb-1.2.0-cp312-cp312-pyodide_2024_0_wasm32.whl";

const post = (type, data) => self.postMessage({ type, data });
let py;

async function boot() {
  post("status", "Loading Python (first visit only, about 15 MB)…");
  py = await loadPyodide();
  post("status", "Loading DuckDB, openpyxl and cellproof…");
  // every wheel is served next to this file (build.sh), so nothing but Pyodide itself comes from elsewhere
  const local = (await (await fetch("wheels.txt")).text()).trim().split("\n");
  // unpacked, not micropip-installed: micropip preloads DuckDB's native library in a way that crashed it in Chromium
  // ("null function" on the first read_csv); unpacked, Python loads it on import, as in Node, where it works
  for (const w of [DUCKDB, ...local]) py.unpackArchive(new Uint8Array(await (await fetch(w)).arrayBuffer()), "wheel");
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
