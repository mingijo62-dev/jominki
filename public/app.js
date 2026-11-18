(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const ROLE_NAMES = {
    no: "번호", process: "공정/작업", category: "분류", hazard: "유해·위험요인", current: "현재 안전조치", law: "관련근거",
    freq: "가능성", sev: "중대성", risk: "위험성", measure: "감소대책",
    after_freq: "개선후 가능성", after_sev: "개선후 중대성", after_risk: "개선후 위험성",
    owner: "담당자", due: "예정일", done: "완료일", status: "상태",
  };
  const KEY_ROLES = ["hazard", "freq", "sev", "risk", "measure", "after_risk", "owner", "due"];
  const LEVEL_KO = { error: "오류", warn: "경고", info: "참고" };

  const state = { wb: null, fileName: "", sheet: "", result: null, filter: "all", opts: {}, aiEnabled: null };

  // ---------- 파일 로드 ----------
  const drop = $("drop"), input = $("file");
  input.addEventListener("change", () => input.files[0] && loadFile(input.files[0]));
  drop.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } });
  ["dragenter", "dragover"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f) loadFile(f); });

  function loadFile(file) {
    const err = $("loadError");
    err.hidden = true;
    if (!/\.(xlsx|xls|xlsm|csv)$/i.test(file.name)) {
      err.textContent = "엑셀(.xlsx, .xls, .xlsm) 또는 CSV 파일만 지원합니다. 한글(.hwp)·PDF 파일은 엑셀로 변환 후 올려주세요.";
      err.hidden = false;
      return;
    }
    if (file.size > 20 * 1024 * 1024) {
      err.textContent = "파일이 너무 큽니다(최대 20MB).";
      err.hidden = false;
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const isCsv = /\.csv$/i.test(file.name);
        let wb;
        if (isCsv) {
          let txt;
          try { txt = new TextDecoder("utf-8", { fatal: true }).decode(data); }
          catch { txt = new TextDecoder("euc-kr").decode(data); }
          wb = XLSX.read(txt, { type: "string", cellDates: true });
        } else {
          wb = XLSX.read(data, { type: "array", cellDates: true });
        }
        state.wb = wb;
        state.fileBuf = e.target.result;
        state.fileName = file.name;
        state.opts = {};
        setupSheets();
      } catch (ex) {
        console.error(ex);
        err.textContent = "파일을 읽지 못했습니다. 암호가 걸려 있거나 손상된 파일일 수 있습니다.";
        err.hidden = false;
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function sheetRows(name) {
    const ws = state.wb.Sheets[name];
    if (!ws["!ref"]) return { ws, rows: [] };
    const rg = XLSX.utils.decode_range(ws["!ref"]);
    rg.s.r = 0; rg.s.c = 0; // A1부터 읽어 행·열 번호를 엑셀과 일치시킴
    return { ws, rows: XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "", blankrows: true, range: rg }) };
  }

  function setupSheets() {
    const sel = $("sheetSel");
    sel.innerHTML = "";
    // 위험성평가표로 보이는 시트를 자동 선택
    let best = state.wb.SheetNames[0], bestScore = -1;
    for (const n of state.wb.SheetNames) {
      const opt = document.createElement("option");
      opt.value = opt.textContent = n;
      sel.appendChild(opt);
      const { rows } = sheetRows(n);
      const hr = RiskAnalyzer.findHeaderRow(rows);
      const score = (hr >= 0 ? 100 : 0) + (/위험성|평가/.test(n) ? 10 : 0) + Math.min(rows.length, 50) / 10;
      if (score > bestScore) { bestScore = score; best = n; }
    }
    sel.value = best;
    $("fileName").textContent = state.fileName;
    $("intro").hidden = true;
    $("result").hidden = false;
    runAnalysis();
    window.scrollTo({ top: 0 });
  }
  $("sheetSel").addEventListener("change", () => { state.opts = {}; runAnalysis(); });
  $("resetBtn").addEventListener("click", () => {
    $("result").hidden = true; $("intro").hidden = false; input.value = ""; state.result = null;
    $("aiOut").hidden = true; $("aiOut").innerHTML = "";
  });

  // ---------- 분석 ----------
  function runAnalysis() {
    state.sheet = $("sheetSel").value;
    const { ws, rows } = sheetRows(state.sheet);
    const res = RiskAnalyzer.analyze(rows, { merges: ws["!merges"], ...state.opts });
    state.result = res;
    $("aiOut").hidden = true; $("aiOut").innerHTML = "";
    if (!res.ok) {
      $("sheetError").textContent = res.error;
      $("sheetError").hidden = false;
      $("report").hidden = true;
      return;
    }
    $("sheetError").hidden = true;
    $("report").hidden = false;
    renderSummary(res);
    renderSettings(res);
    renderFindings();
    renderTable(res);
  }

  function renderSummary(res) {
    $("score").textContent = res.score;
    const color = res.verdict === "적정" ? "var(--ok)" : res.verdict === "보완 필요" ? "var(--warn)" : "var(--err)";
    const ring = $("ring");
    ring.style.setProperty("--p", res.score);
    ring.style.setProperty("--c", color);
    $("verdict").textContent = res.verdict;
    $("verdict").style.color = color;
    $("itemCount").textContent = `평가 항목 ${res.items.length}개 · 헤더 ${res.headerRow}행`;
    $("cErr").textContent = res.counts.error;
    $("cWarn").textContent = res.counts.warn;
    $("cInfo").textContent = res.counts.info;
  }

  function renderSettings(res) {
    const s = res.settings;
    $("optMethod").value = s.method;
    $("optF").value = String(s.fScale);
    $("optS").value = String(s.sScale);
    $("optAcc").value = s.acceptable;
    $("settingsBrief").textContent = `· ${s.fScale}×${s.sScale} 척도, ${s.method === "mul" ? "곱셈" : "덧셈"}, 허용 ≤ ${s.acceptable}`;
    const cols = res.columns;
    $("colsInfo").innerHTML = Object.keys(ROLE_NAMES)
      .filter((k) => cols[k] || KEY_ROLES.includes(k))
      .map((k) => cols[k]
        ? `<span class="chip">${ROLE_NAMES[k]}: <b>${esc(cols[k].label || "열 " + (cols[k].col + 1))}</b></span>`
        : `<span class="chip miss">${ROLE_NAMES[k]}: 없음</span>`)
      .join("");
    const m = res.meta;
    $("metaInfo").innerHTML = [
      `평가일자: <b>${esc(m.evalDate || "미확인")}</b>`,
      m.typeText ? `평가종류: <b>${esc(m.typeText)}</b>` : "",
      `근로자 참여 기록: <b>${m.workerParticipation ? "있음" : "미확인"}</b>`,
    ].filter(Boolean).map((t) => `<span class="chip">${t}</span>`).join("");
  }
  $("rerunBtn").addEventListener("click", () => {
    state.opts = {
      method: $("optMethod").value,
      freqScale: Number($("optF").value),
      sevScale: Number($("optS").value),
      acceptable: Number($("optAcc").value),
    };
    runAnalysis();
  });

  // ---------- 지적사항 ----------
  function renderFindings() {
    const res = state.result;
    const list = res.findings.filter((f) => state.filter === "all" || f.level === state.filter);
    const ul = $("findings");
    if (!list.length) {
      ul.innerHTML = `<li class="empty">${state.filter === "all" ? "지적 사항이 없습니다. 👍" : "해당 수준의 지적 사항이 없습니다."}</li>`;
      return;
    }
    ul.innerHTML = list.map((f) => `
      <li ${f.row ? `data-row="${f.row}"` : ""}>
        <span class="tag ${f.level}">${LEVEL_KO[f.level]}</span>
        <span class="rowref">${f.row ? f.row + "행" : "전체"}</span>
        <span class="msg">${esc(f.msg)}</span>
      </li>`).join("");
  }
  document.querySelectorAll("[data-filter]").forEach((b) => b.addEventListener("click", () => {
    state.filter = state.filter === b.dataset.filter && b.classList.contains("count") ? "all" : b.dataset.filter;
    document.querySelectorAll(".seg button").forEach((x) => x.classList.toggle("on", x.dataset.filter === state.filter));
    renderFindings();
    if (b.classList.contains("count")) $("findings").scrollIntoView({ behavior: "smooth", block: "center" });
  }));
  $("findings").addEventListener("click", (e) => {
    const li = e.target.closest("li[data-row]");
    if (!li) return;
    const tr = document.querySelector(`#itemsTable tr[data-row="${li.dataset.row}"]`);
    if (!tr) return;
    tr.scrollIntoView({ behavior: "smooth", block: "center" });
    tr.classList.add("flash");
    setTimeout(() => tr.classList.remove("flash"), 1600);
  });

  // ---------- 표 ----------
  function renderTable(res) {
    const has = (k) => !!res.columns[k];
    const byRow = new Map();
    for (const f of res.findings) if (f.row) { if (!byRow.has(f.row)) byRow.set(f.row, []); byRow.get(f.row).push(f); }
    const fields = [
      ["rowNum", "행", "num"],
      has("process") && ["process", "공정/작업", "proc"],
      ["hazard", "유해·위험요인", "wide"],
      has("freq") && ["freq", "가능성", "num"],
      has("sev") && ["sev", "중대성", "num"],
      has("risk") && ["riskRaw", "위험성", "num", "risk"],
      ["measure", "감소대책", "wide"],
      has("after_freq") && ["afterFreq", "후 가능성", "num", "after_freq"],
      has("after_sev") && ["afterSev", "후 중대성", "num", "after_sev"],
      has("after_risk") && ["afterRiskRaw", "후 위험성", "num", "after_risk"],
      has("owner") && ["owner", "담당자", ""],
      has("due") && ["due", "예정일", "num"],
      has("done") && ["done", "완료일", "num"],
    ].filter(Boolean);
    const head = `<thead><tr><th>${fields[0][1]}</th><th>검토 결과</th>${fields.slice(1).map((f) => `<th>${f[1]}</th>`).join("")}</tr></thead>`;
    const body = res.items.map((it) => {
      const fs = byRow.get(it.rowNum) || [];
      const lvl = fs.some((f) => f.level === "error") ? "r-error" : fs.some((f) => f.level === "warn") ? "r-warn" : "";
      const badFields = new Set(fs.filter((f) => f.level !== "info").map((f) => f.field));
      const tdList = fields.map(([key, , cls, field]) => {
        let v = it[key];
        if (v instanceof Date) v = v.toISOString().slice(0, 10);
        const fieldName = field || key;
        const bad = badFields.has(fieldName) ? " bad" : "";
        return `<td class="${cls}${bad}">${esc(v ?? "")}</td>`;
      });
      const iss = fs.map((f) => `<div class="${f.level[0]}">• ${esc(f.msg)}</div>`).join("") || `<span class="muted">이상 없음</span>`;
      return `<tr data-row="${it.rowNum}" class="${lvl}">${tdList[0]}<td class="issues">${iss}</td>${tdList.slice(1).join("")}</tr>`;
    }).join("");
    $("itemsTable").innerHTML = head + `<tbody>${body}</tbody>`;
  }

  // ---------- 엑셀 다운로드 (원본 파일에 직접 표시) ----------
  const FILL = { error: "FFFFC7CE", warn: "FFFFEB9C", info: "FFDDEBF7" };
  const FONT = { error: "FF9C0006", warn: "FF9C5700", info: "FF1F4E79" };
  const RANK = { error: 3, warn: 2, info: 1 };
  const solid = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
  const thin = { style: "thin", color: { argb: "FFBFBFBF" } };
  // ExcelJS는 같은 서식의 칸끼리 스타일 객체를 공유하므로, 칸마다 복사본을 만든 뒤 수정
  const own = (cell) => { cell.style = JSON.parse(JSON.stringify(cell.style || {})); return cell; };

  $("dlBtn").addEventListener("click", async () => {
    const res = state.result;
    if (!res || !res.ok) return;
    const btn = $("dlBtn");
    const label = btn.textContent;
    btn.disabled = true; btn.textContent = "만드는 중…";
    try {
      const { wb, ws, preserved } = await openTargetWorkbook();
      markSheet(ws, res);
      addSummarySheets(wb, res, preserved);
      const buf = await wb.xlsx.writeBuffer();
      const base = state.fileName.replace(/\.[^.]+$/, "");
      const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      const a = document.createElement("a");
      a.href = url; a.download = `${base}_검토결과.xlsx`;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    } catch (ex) {
      console.error(ex);
      alert("결과 파일을 만들지 못했습니다: " + ex.message);
    } finally {
      btn.disabled = false; btn.textContent = label;
    }
  });

  // .xlsx/.xlsm 은 원본 그대로 열고, .xls/.csv 는 값으로 새 통합문서를 만든다
  async function openTargetWorkbook() {
    const wb = new ExcelJS.Workbook();
    if (/\.(xlsx|xlsm)$/i.test(state.fileName)) {
      try {
        await wb.xlsx.load(state.fileBuf.slice(0));
        const ws = wb.getWorksheet(state.sheet);
        if (ws) return { wb, ws, preserved: true };
      } catch (ex) {
        console.warn("원본 열기 실패, 값으로 대체", ex);
      }
    }
    const wb2 = new ExcelJS.Workbook();
    const ws = wb2.addWorksheet(state.sheet.slice(0, 31) || "Sheet1");
    const { ws: src, rows } = sheetRows(state.sheet);
    rows.forEach((r, i) => r.forEach((v, c) => {
      if (v !== "" && v != null) ws.getCell(i + 1, c + 1).value = v instanceof Date ? v : v;
    }));
    for (const m of src["!merges"] || []) {
      try { ws.mergeCells(m.s.r + 1, m.s.c + 1, m.e.r + 1, m.e.c + 1); } catch {}
    }
    const width = Math.max(0, ...rows.map((r) => r.length));
    for (let c = 1; c <= width; c++) ws.getColumn(c).width = 14;
    return { wb: wb2, ws, preserved: false };
  }

  function markSheet(ws, res) {
    const byRow = new Map();
    for (const f of res.findings) if (f.row) { if (!byRow.has(f.row)) byRow.set(f.row, []); byRow.get(f.row).push(f); }

    // 검토결과 열 위치: 기존 사용 열의 바로 오른쪽
    const { rows } = sheetRows(state.sheet);
    const lastCol = Math.max(ws.columnCount || 0, ...rows.map((r) => r.length));
    const outCol = lastCol + 1;
    ws.getColumn(outCol).width = 60;

    // 헤더
    const hr = res.headerRow;
    const head = own(ws.getCell(hr, outCol));
    head.value = "검토결과 (자동)";
    head.font = { bold: true, color: { argb: "FFFFFFFF" } };
    head.fill = solid("FFC00000");
    head.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    head.border = { top: thin, left: thin, bottom: thin, right: thin };
    if (res.dataStartRow > hr + 1) {
      try { ws.mergeCells(hr, outCol, res.dataStartRow - 1, outCol); } catch {}
    }

    // 항목별 결과
    for (const it of res.items) {
      const fs = byRow.get(it.rowNum) || [];
      const cell = own(ws.getCell(it.rowNum, outCol));
      cell.border = { top: thin, left: thin, bottom: thin, right: thin };
      cell.alignment = { vertical: "top", wrapText: true };
      if (!fs.length) {
        cell.value = "이상 없음";
        cell.font = { color: { argb: "FF006100" } };
        continue;
      }
      const worst = fs.reduce((a, f) => (RANK[f.level] > RANK[a] ? f.level : a), "info");
      cell.value = fs.map((f) => `[${LEVEL_KO[f.level]}] ${f.msg}`).join("\n");
      cell.font = { color: { argb: FONT[worst] } };
      cell.fill = solid(FILL[worst]);

      // 문제 칸에 색칠 + 메모
      const byField = new Map();
      for (const f of fs) {
        if (!f.field || !res.columns[f.field]) continue;
        if (!byField.has(f.field)) byField.set(f.field, []);
        byField.get(f.field).push(f);
      }
      for (const [field, list] of byField) {
        const target = own(ws.getCell(it.rowNum, res.columns[field].col + 1));
        const lv = list.reduce((a, f) => (RANK[f.level] > RANK[a] ? f.level : a), "info");
        target.fill = solid(FILL[lv]);
        const msg = "[검토] " + list.map((f) => `${LEVEL_KO[f.level]}: ${f.msg}`).join("\n");
        const prev = target.note ? (typeof target.note === "string" ? target.note : (target.note.texts || []).map((t) => t.text).join("")) : "";
        target.note = prev ? `${prev}\n\n${msg}` : msg;
      }
    }
  }

  function addSummarySheets(wb, res, preserved) {
    const s = wb.addWorksheet("검토요약");
    s.getColumn(1).width = 16; s.getColumn(2).width = 90;
    const put = (r, a, b, opt = {}) => {
      s.getCell(r, 1).value = a; s.getCell(r, 2).value = b;
      s.getCell(r, 1).font = { bold: true };
      s.getCell(r, 2).alignment = { wrapText: true, vertical: "top" };
      if (opt.fill) s.getCell(r, 2).fill = solid(opt.fill);
    };
    s.getCell(1, 1).value = "위험성평가 재검토 결과";
    s.getCell(1, 1).font = { bold: true, size: 16 };
    let r = 3;
    put(r++, "파일", state.fileName);
    put(r++, "검토 시트", `${state.sheet} (이 시트의 오른쪽 끝에 ‘검토결과’ 열이 추가됨)`);
    put(r++, "검토일", new Date().toISOString().slice(0, 10));
    put(r++, "점수 / 판정", `${res.score}점 / ${res.verdict}`);
    put(r++, "지적 건수", `오류 ${res.counts.error} · 경고 ${res.counts.warn} · 참고 ${res.counts.info}`);
    put(r++, "평가 기준", `가능성 1~${res.settings.fScale} ${res.settings.method === "mul" ? "×" : "+"} 중대성 1~${res.settings.sScale}, 허용 가능 위험성 ≤ ${res.settings.acceptable}`);
    put(r++, "색 표시", "빨강 = 오류", { fill: FILL.error });
    put(r++, "", "노랑 = 경고", { fill: FILL.warn });
    put(r++, "", "파랑 = 참고  (색칠된 칸에 마우스를 올리면 메모로 이유가 보입니다)", { fill: FILL.info });
    if (!preserved) put(r++, "안내", "원본 형식(.xls/.csv 등)을 직접 수정할 수 없어 값만 옮긴 새 파일로 만들었습니다.");
    r++;
    const general = res.findings.filter((f) => !f.row);
    if (general.length) {
      s.getCell(r++, 1).value = "문서 전체 지적사항";
      s.getCell(r - 1, 1).font = { bold: true, size: 13 };
      for (const f of general) put(r++, LEVEL_KO[f.level], f.msg, { fill: FILL[f.level] });
      r++;
    }
    s.getCell(r++, 1).value = "※ 자동 검토 결과는 참고용이며, 최종 판단은 평가 실시 주체와 안전보건 전문가가 해야 합니다.";

    const d = wb.addWorksheet("지적사항목록");
    d.columns = [{ header: "수준", width: 8 }, { header: "행", width: 7 }, { header: "항목", width: 16 }, { header: "내용", width: 100 }];
    d.getRow(1).font = { bold: true };
    for (const f of res.findings) {
      const row = d.addRow([LEVEL_KO[f.level], f.row || "전체", f.field ? (res.columns[f.field] ? res.columns[f.field].label : f.field) : "", f.msg]);
      row.getCell(1).fill = solid(FILL[f.level]);
      row.getCell(4).alignment = { wrapText: true };
    }
    d.views = [{ state: "frozen", ySplit: 1 }];

    // 요약 시트를 맨 앞으로, 검토 시트를 활성화
    const order = [s, ...wb.worksheets.filter((w) => w !== s)];
    order.forEach((w, i) => { w.orderNo = i; });
    wb.views = [{ activeTab: 0, firstSheet: 0 }];
  }

  // ---------- AI 검토 ----------
  fetch("/api/review").then((r) => r.ok ? r.json() : null).then((j) => {
    state.aiEnabled = !!(j && j.aiEnabled);
    if (!state.aiEnabled) {
      $("aiBtn").disabled = true;
      $("aiNote").textContent = "AI 정밀 검토는 서버에 API 키가 설정된 경우에만 사용할 수 있습니다. (규칙 기반 검토는 정상 동작합니다)";
    }
  }).catch(() => { state.aiEnabled = false; $("aiBtn").disabled = true; });

  $("aiBtn").addEventListener("click", async () => {
    const res = state.result;
    if (!res || !res.ok) return;
    const btn = $("aiBtn"), out = $("aiOut");
    btn.disabled = true;
    out.hidden = false;
    out.innerHTML = `<p class="muted"><span class="spinner"></span>AI가 평가표를 검토하고 있습니다… (1~3분 소요)</p>`;
    const rows = res.items.slice(0, 400).map((i) => ({
      행: i.rowNum, 공정: i.process, 유해위험요인: i.hazard, 현재조치: i.current || undefined,
      가능성: i.freq, 중대성: i.sev, 위험성: i.riskRaw === "" ? undefined : i.riskRaw,
      감소대책: i.measure, 개선후가능성: i.afterFreq ?? undefined, 개선후중대성: i.afterSev ?? undefined,
      개선후위험성: i.afterRiskRaw === "" ? undefined : i.afterRiskRaw, 담당자: i.owner || undefined, 예정일: i.due || undefined, 완료일: i.done || undefined,
    }));
    const body = {
      fileName: state.fileName, sheetName: state.sheet,
      meta: { ...res.meta, 평가기준: res.settings, 전체항목수: res.items.length },
      rows,
      ruleFindings: res.findings.slice(0, 200).map((f) => `${f.row ? f.row + "행" : "전체"} [${LEVEL_KO[f.level]}] ${f.msg}`),
    };
    let acc = "";
    try {
      const r = await fetch("/api/review", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
      if (!r.ok) {
        let msg = `요청 실패 (${r.status})`;
        try { msg = (await r.json()).error || msg; } catch {}
        throw new Error(msg);
      }
      const reader = r.body.getReader();
      const dec = new TextDecoder();
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        acc += dec.decode(value, { stream: true });
        out.innerHTML = md(acc) + `<p class="muted small"><span class="spinner"></span>작성 중…</p>`;
      }
      out.innerHTML = md(acc);
    } catch (ex) {
      out.innerHTML = (acc ? md(acc) : "") + `<p class="err">${esc(ex.message)}</p>`;
    } finally {
      btn.disabled = false;
      btn.textContent = "AI 검토 다시 실행";
    }
  });

  // 간단 Markdown 렌더러 (제목, 목록, 굵게, 인용, 코드)
  function md(src) {
    const inline = (s) => esc(s)
      .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
      .replace(/`([^`]+)`/g, "<code>$1</code>");
    const lines = src.split(/\r?\n/);
    let html = "", list = null;
    const closeList = () => { if (list) { html += `</${list}>`; list = null; } };
    for (const line of lines) {
      let m;
      if ((m = line.match(/^(#{1,4})\s+(.*)/))) { closeList(); const lv = Math.min(3, m[1].length + 1); html += `<h${lv}>${inline(m[2])}</h${lv}>`; }
      else if ((m = line.match(/^\s*[-*•]\s+(.*)/))) { if (list !== "ul") { closeList(); html += "<ul>"; list = "ul"; } html += `<li>${inline(m[1])}</li>`; }
      else if ((m = line.match(/^\s*\d+[.)]\s+(.*)/))) { if (list !== "ol") { closeList(); html += "<ol>"; list = "ol"; } html += `<li>${inline(m[1])}</li>`; }
      else if ((m = line.match(/^>\s?(.*)/))) { closeList(); html += `<blockquote>${inline(m[1])}</blockquote>`; }
      else if (!line.trim()) { closeList(); }
      else { closeList(); html += `<p>${inline(line)}</p>`; }
    }
    closeList();
    return html;
  }
})();
