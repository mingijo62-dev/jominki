// 사용법: node scripts/test-any.cjs <파일> [시트명]
const XLSX = require("xlsx");
require("vm").runInThisContext(require("fs").readFileSync(__dirname + "/../public/analyzer.js", "utf8"));
const wb = XLSX.readFile(process.argv[2], { cellDates: true });
const name = process.argv[3] || wb.SheetNames[0];
const ws = wb.Sheets[name];
const rg = XLSX.utils.decode_range(ws["!ref"]); rg.s.r = 0; rg.s.c = 0;
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "", blankrows: true, range: rg });
const res = globalThis.RiskAnalyzer.analyze(rows, { merges: ws["!merges"], today: "2026-11-18" });
if (!res.ok) { console.log(res.error); process.exit(); }
console.log(Object.entries(res.columns).map(([k, v]) => `${k}=${v.label}`).join(" | "));
console.log(res.settings, res.meta, "items", res.items.length, "score", res.score, res.verdict);
for (const f of res.findings) console.log(f.level, f.row ?? "-", f.code, f.msg);
