// 로컬 개발 서버: public/ 정적 파일
import http from "node:http";
import { readFile } from "node:fs/promises";
import path from "node:path";
const root = path.resolve(import.meta.dirname, "../public");
const types = { ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".css": "text/css", ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" };
http.createServer(async (req, res) => {
  const url = new URL(req.url, "http://x");
  const p = path.join(root, url.pathname === "/" ? "index.html" : decodeURIComponent(url.pathname));
  if (!p.startsWith(root)) { res.writeHead(403); return res.end(); }
  try { const d = await readFile(p); res.writeHead(200, { "Content-Type": types[path.extname(p)] || "application/octet-stream" }); res.end(d); }
  catch { res.writeHead(404); res.end("not found"); }
}).listen(process.env.PORT || 3000, () => console.log("http://localhost:" + (process.env.PORT || 3000)));
