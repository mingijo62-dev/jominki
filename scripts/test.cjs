const XLSX = require("xlsx");
require("vm").runInThisContext(require("fs").readFileSync(__dirname + "/../public/analyzer.js", "utf8")); const A = globalThis.RiskAnalyzer;
const wb = XLSX.readFile(process.argv[2], { cellDates: true });
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
const res = A.analyze(rows, { merges: ws["!merges"], today: "2026-11-18" });
console.log(JSON.stringify({ columns: res.columns, settings: res.settings, meta: res.meta, score: res.score, verdict: res.verdict, counts: res.counts, n: res.items.length }, null, 1));
for (const f of res.findings) console.log(f.level, f.row ?? "-", f.code, f.msg);
