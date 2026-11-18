// 로컬 개발 서버: public/ 정적 파일 + api/review.js
import http from "node:http";
import { readFile } from "node:fs/promises";
import path from "node:path";
import * as review from "../api/review.js";
const root = path.resolve(import.meta.dirname, "../public");
const types = { ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".css": "text/css", ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" };
http.createServer(async (req, res) => {
  const url = new URL(req.url, "http://x");
  if (url.pathname === "/api/review") {
    const body = req.method === "POST" ? await new Promise((r) => { const c = []; req.on("data", (d) => c.push(d)); req.on("end", () => r(Buffer.concat(c))); }) : undefined;
    const r = await review[req.method](new Request(url, { method: req.method, headers: req.headers, body }));
    res.writeHead(r.status, Object.fromEntries(r.headers));
    if (r.body) for await (const chunk of r.body) res.write(chunk);
    return res.end();
  }
  const p = path.join(root, url.pathname === "/" ? "index.html" : decodeURIComponent(url.pathname));
  if (!p.startsWith(root)) { res.writeHead(403); return res.end(); }
  try { const d = await readFile(p); res.writeHead(200, { "Content-Type": types[path.extname(p)] || "application/octet-stream" }); res.end(d); }
  catch { res.writeHead(404); res.end("not found"); }
}).listen(process.env.PORT || 3000, () => console.log("http://localhost:" + (process.env.PORT || 3000)));
