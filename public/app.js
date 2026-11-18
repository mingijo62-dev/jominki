(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const state = { wb: null, fileBuf: null, fileName: "", sheet: "", result: null };

  // ---------- 파일 로드 ----------
  const drop = $("drop"), input = $("file");
  input.addEventListener("change", () => input.files[0] && loadFile(input.files[0]));
  drop.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); } });
  ["dragenter", "dragover"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f) loadFile(f); });

  function showLoadError(msg) { const err = $("loadError"); err.textContent = msg; err.hidden = false; }

  function loadFile(file) {
    $("loadError").hidden = true;
    if (!/\.(xlsx|xls|xlsm|csv)$/i.test(file.name)) return showLoadError("엑셀(.xlsx, .xls, .xlsm) 또는 CSV 파일만 지원합니다. 한글(.hwp)·PDF 파일은 엑셀로 변환 후 올려주세요.");
    if (file.size > 20 * 1024 * 1024) return showLoadError("파일이 너무 큽니다(최대 20MB).");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        let wb;
        if (/\.csv$/i.test(file.name)) {
          let txt;
          try { txt = new TextDecoder("utf-8", { fatal: true }).decode(data); } catch { txt = new TextDecoder("euc-kr").decode(data); }
          wb = XLSX.read(txt, { type: "string", cellDates: true });
        } else {
          wb = XLSX.read(data, { type: "array", cellDates: true });
        }
        state.wb = wb;
        state.fileBuf = e.target.result;
        state.fileName = file.name;
        setupSheets();
      } catch (ex) {
        console.error(ex);
        showLoadError("파일을 읽지 못했습니다. 암호가 걸려 있거나 손상된 파일일 수 있습니다.");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function sheetRows(name) {
    const ws = state.wb.Sheets[name];
    if (!ws || !ws["!ref"]) return { ws: ws || {}, rows: [] };
    // 실제 값이 있는 칸 기준으로 범위 계산 (서식만 시트 전체에 걸린 파일 대비), A1부터 읽어 행 번호를 엑셀과 일치
    let maxR = 0, maxC = 0;
    for (const k of Object.keys(ws)) {
      if (k[0] === "!") continue;
      const cell = ws[k];
      if (!cell || cell.v == null || cell.v === "") continue;
      const a = XLSX.utils.decode_cell(k);
      if (a.r > maxR) maxR = a.r;
      if (a.c > maxC) maxC = a.c;
    }
    const rg = { s: { r: 0, c: 0 }, e: { r: Math.min(maxR, 20000), c: Math.min(maxC, 200) } };
    return { ws, rows: XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "", blankrows: true, range: rg }) };
  }

  function parseSheet(name) {
    const { ws, rows } = sheetRows(name);
    const merges = (ws["!merges"] || []).filter((m) => m.s.r <= 20000);
    return RiskAnalyzer.analyze(rows, { merges });
  }

  function setupSheets() {
    const sel = $("sheetSel");
    sel.innerHTML = "";
    let best = state.wb.SheetNames[0], bestScore = -1;
    for (const n of state.wb.SheetNames) {
      const opt = document.createElement("option");
      opt.value = opt.textContent = n;
      sel.appendChild(opt);
      let score = 0;
      try {
        const r = parseSheet(n);
        if (r.ok) score = 100 + Object.keys(r.columns).length * 10 + Math.min(r.items.length, 300) / 10;
      } catch (ex) { console.warn(n, ex); }
      if (/위험성|평가/.test(n)) score += 5;
      if (score > bestScore) { bestScore = score; best = n; }
    }
    sel.value = best;
    $("fileName").textContent = state.fileName;
    $("intro").hidden = true;
    $("result").hidden = false;
    refresh();
    window.scrollTo({ top: 0 });
  }
  $("sheetSel").addEventListener("change", refresh);
  $("resetBtn").addEventListener("click", () => {
    $("result").hidden = true; $("intro").hidden = false; input.value = ""; state.result = null;
  });

  // ---------- 화면 ----------
  // 보완목록에 옮겨 적을 원본 열 (있는 것만)
  const COPY_FIELDS = [
    ["process", "작업내용", 14],
    ["division", "평가구분", 8],
    ["hazard", "위험요인 및 재해형태", 40],
    ["category", "재해형태", 9],
    ["current", "현재안전조치", 24],
    ["freq", "빈도", 5],
    ["sev", "강도", 5],
    ["risk", "위험도", 6, "riskRaw"],
  ];
  const copyFieldsFor = (res) => COPY_FIELDS.filter(([k]) => res.columns[k]).map(([k, title, w, prop]) => ({
    key: k, prop: prop || k, width: w,
    // 원본 열 제목을 그대로 사용 (빈도·강도·위험도는 짧은 이름 유지)
    title: ["freq", "sev", "risk"].includes(k) ? title : (res.columns[k].label || title).replace(/\s+/g, " ").trim(),
  }));

  function refresh() {
    state.sheet = $("sheetSel").value;
    let res;
    try { res = parseSheet(state.sheet); } catch (ex) { console.error(ex); res = { ok: false, error: "이 시트를 읽는 중 오류가 발생했습니다: " + ex.message }; }
    state.result = res;
    if (!res.ok || !res.items.length) {
      $("sheetError").textContent = res.ok ? "이 시트에서 평가 항목을 찾지 못했습니다. 위의 ‘평가표 시트’에서 다른 시트를 골라 보세요." : res.error;
      $("sheetError").hidden = false;
      $("report").hidden = true;
      return;
    }
    $("sheetError").hidden = true;
    $("report").hidden = false;
    const fields = copyFieldsFor(res);
    $("itemCount").textContent = `평가 항목 ${res.items.length}개를 찾았습니다`;
    $("colsBrief").textContent = `헤더 ${res.headerRow}행 · 데이터 ${res.items[0].rowNum}~${res.items[res.items.length - 1].rowNum}행`;
    const head = `<thead><tr><th>행</th>${fields.map((f) => `<th>${esc(f.title)}</th>`).join("")}</tr></thead>`;
    const body = res.items.map((it) => `<tr><td class="num">${it.rowNum}</td>${fields.map((f) =>
      `<td class="${["freq", "sev", "risk"].includes(f.key) ? "num" : f.key === "hazard" || f.key === "current" ? "wide" : ""}">${esc(it[f.prop])}</td>`).join("")}</tr>`).join("");
    $("itemsTable").innerHTML = head + `<tbody>${body}</tbody>`;
  }

  // ---------- 보완목록 시트 만들기 ----------
  const solid = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
  const thin = { style: "thin", color: { argb: "FF7F7F7F" } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const colLetter = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const C = { head: "FF404040", copyHead: "FF595959", inputHead: "FFC65911", autoHead: "FF2F5597", input: "FFFFF9E6", copy: "FFF2F2F2" };

  $("dlBtn").addEventListener("click", async () => {
    const res = state.result;
    if (!res || !res.ok) return;
    const btn = $("dlBtn"), label = btn.textContent;
    btn.disabled = true; btn.textContent = "만드는 중…";
    try {
      const { wb, preserved } = await openWorkbook();
      const ws = wb.getWorksheet(state.sheet.slice(0, 31));
      if (!ws) throw new Error("평가표 시트를 찾지 못했습니다.");
      const layout = addInputArea(ws, res);
      const resultCreated = addResultSheet(wb, ws, res, layout);
      // 파일을 열면 평가표 시트가 보이게
      wb.worksheets.forEach((w) => { if (w.views && w.views[0]) w.views[0].tabSelected = w === ws; });
      wb.views = [{ activeTab: wb.worksheets.indexOf(ws), firstSheet: 0, visibility: "visible" }];
      const buf = await wb.xlsx.writeBuffer();
      const base = state.fileName.replace(/\.[^.]+$/, "").replace(/_보완목록$/, "");
      download(buf, `${base}_보완목록.xlsx`);
      if (layout.existing && !resultCreated) alert("이 파일에는 이미 작성란과 ‘보완결과’ 시트가 있어 바꾼 내용이 없습니다.");
      else if (layout.existing) alert("이미 있던 작성란(입력한 내용 포함)은 그대로 두고 ‘보완결과’ 시트를 추가했습니다.");
      if (!preserved) alert("원본 형식(.xls/.csv)은 그대로 수정할 수 없어, 값만 옮긴 새 .xlsx 파일로 만들었습니다.");
    } catch (ex) {
      console.error(ex);
      alert("파일을 만들지 못했습니다: " + ex.message);
    } finally {
      btn.disabled = false; btn.textContent = label;
    }
  });

  // 원본 평가표 시트의 표 오른쪽(마지막 사용 열 다음)에 보완 작성란을 붙인다
  const INPUT_COLS = [
    { key: "judge", title: "유효성 검토", width: 11, kind: "input" },
    { key: "issue", title: "보완 필요 사항", width: 30, kind: "input" },
    { key: "measure", title: "보완 대책", width: 30, kind: "input" },
    { key: "af", title: "보완후 빈도", width: 9, kind: "input" },
    { key: "as", title: "보완후 강도", width: 9, kind: "input" },
    { key: "ar", title: "보완후 위험도", width: 10, kind: "auto" },
    { key: "owner", title: "담당자", width: 9, kind: "input" },
    { key: "due", title: "완료 예정일", width: 11, kind: "input" },
    { key: "done", title: "완료 확인", width: 9, kind: "input" },
    { key: "memo", title: "비고", width: 14, kind: "input" },
  ];
  // 작성란 뒤에 숨겨 두는 보조 열 (보완결과 시트가 참조)
  const HELPER_COLS = ["보완순번", "작업내용(보조)", "평가구분(보조)"];

  // 이미 작성란이 있는 파일(다시 올린 경우)인지 확인
  function findExistingArea(ws, res) {
    const hr = res.headerRow, last = Math.max(ws.columnCount || 0, 1);
    for (let r = hr; r < res.dataStartRow; r++)
      for (let c = 1; c <= last; c++) {
        const v = ws.getCell(r, c).value;
        if (typeof v === "string" && v.replace(/\s+/g, "").startsWith("유효성검토")) return c;
      }
    return 0;
  }

  function addInputArea(ws, res) {
    const s = res.settings;
    const items = res.items;
    const hr = res.headerRow;
    const r1 = items[0].rowNum, r2 = items[items.length - 1].rowNum;

    const existing = findExistingArea(ws, res);
    if (existing) {
      const ci = Object.fromEntries(INPUT_COLS.map((c, i) => [c.key, existing + i]));
      const layout = { existing: true, ci, r1, r2, c0: existing, cEnd: existing + INPUT_COLS.length - 1 };
      addHelperCols(ws, res, layout);
      return layout;
    }

    // 시작 열: 헤더~마지막 항목 행에서 값이 있는 가장 오른쪽 열의 다음 열
    const { rows } = sheetRows(state.sheet);
    let lastCol = 0;
    for (let r = Math.max(0, hr - 2); r < Math.min(rows.length, r2); r++) {
      const row = rows[r] || [];
      for (let c = row.length - 1; c >= 0; c--) {
        if (row[c] !== "" && row[c] != null) { if (c + 1 > lastCol) lastCol = c + 1; break; }
      }
    }
    const c0 = lastCol + 1;
    const ci = Object.fromEntries(INPUT_COLS.map((c, i) => [c.key, c0 + i]));
    const cEnd = c0 + INPUT_COLS.length - 1;
    INPUT_COLS.forEach((c, i) => { ws.getColumn(c0 + i).width = c.width; });

    // 원본 글자 크기에 맞춤
    const hz = res.columns.hazard ? ws.getCell(r1, res.columns.hazard.col + 1) : null;
    const size = (hz && hz.font && hz.font.size) || 9;
    const fontName = hz && hz.font && hz.font.name;
    const font = (extra = {}) => ({ size, ...(fontName ? { name: fontName } : {}), ...extra });

    // 머리글: 원본이 2단 헤더면 위 칸에 묶음 제목, 아래 칸에 항목명
    const headRows = Math.max(1, res.dataStartRow - hr);
    const styleHead = (cell, argb) => {
      cell.style = {};
      cell.font = font({ bold: true, color: { argb: "FFFFFFFF" } });
      cell.fill = solid(argb);
      // 원본 머리글 행이 낮아도 잘리지 않게 한 줄 + 칸에 맞춰 축소
      cell.alignment = { horizontal: "center", vertical: "middle", shrinkToFit: true };
      cell.border = border;
    };
    let subRow = hr;
    if (headRows >= 2) {
      const g = ws.getCell(hr, c0);
      try { ws.mergeCells(hr, c0, hr, cEnd); } catch {}
      g.value = "보완목록 (유효성 검토 · 직접 작성)";
      styleHead(g, "FFC65911");
      subRow = hr + 1;
    }
    INPUT_COLS.forEach((col, i) => {
      const cell = ws.getCell(subRow, c0 + i);
      cell.value = col.title;
      styleHead(cell, col.kind === "auto" ? C.autoHead : "FFE26B0A");
      if (res.dataStartRow - 1 > subRow) { try { ws.mergeCells(subRow, c0 + i, res.dataStartRow - 1, c0 + i); } catch {} }
    });

    // 진행 현황: 머리글 바로 위 칸 (작성하면 자동 갱신)
    if (hr > 1) {
      const rng = (k) => `${colLetter(ci[k])}${r1}:${colLetter(ci[k])}${r2}`;
      const total = items.length;
      const stat = [
        ["전체", String(total)],
        ["검토", `COUNTA(${rng("judge")})`],
        ["보완필요", `COUNTIF(${rng("judge")},"보완필요")`],
        ["보완완료", `COUNTIF(${rng("judge")},"보완완료")`],
        ["미검토", `${total}-COUNTA(${rng("judge")})`],
      ];
      const pr = hr - 1;
      let free = true;
      for (let c = c0; c <= cEnd; c++) { const v = ws.getCell(pr, c).value; if (v !== null && v !== undefined && v !== "") free = false; }
      if (free) {
        try { ws.mergeCells(pr, c0, pr, cEnd); } catch {}
        const p = ws.getCell(pr, c0);
        p.value = { formula: '"진행 현황 ▶ " & ' + stat.map(([k, f]) => `"${k} "&${f}`).join(' & "  ·  " & ') };
        p.style = {};
        p.font = font({ bold: true, color: { argb: "FF7F3F00" } });
        p.fill = solid("FFFFF2CC");
        p.alignment = { horizontal: "center", vertical: "middle", shrinkToFit: true };
        p.border = border;
      }
    }

    // 각 평가 항목 행에 작성 칸 (원본 행 높이는 그대로 둠)
    const plus = s.method === "add" ? "+" : "*";
    const itemRows = new Set(items.map((it) => it.rowNum));
    for (let r = r1; r <= r2; r++) {
      if (!itemRows.has(r)) continue;
      INPUT_COLS.forEach((col, i) => {
        const cell = ws.getCell(r, c0 + i);
        cell.style = {};
        cell.border = border;
        cell.font = font();
        const left = col.key === "issue" || col.key === "measure" || col.key === "memo";
        cell.alignment = { vertical: "middle", horizontal: left ? "left" : "center", wrapText: true };
        if (col.kind === "input") cell.fill = solid(C.input);
      });
      const fL = colLetter(ci.af), sL = colLetter(ci.as);
      ws.getCell(r, ci.ar).value = { formula: `IF(AND(ISNUMBER(${fL}${r}),ISNUMBER(${sL}${r})),${fL}${r}${plus}${sL}${r},"")` };
      ws.getCell(r, ci.judge).dataValidation = { type: "list", allowBlank: true, formulae: ['"적정,보완필요,해당없음,보완완료"'], showErrorMessage: true, errorTitle: "유효성 검토", error: "목록에서 선택하세요." };
      ws.getCell(r, ci.af).dataValidation = { type: "whole", operator: "between", allowBlank: true, formulae: [1, s.fScale], showErrorMessage: true, error: `1~${s.fScale} 사이 숫자를 입력하세요.` };
      ws.getCell(r, ci.as).dataValidation = { type: "whole", operator: "between", allowBlank: true, formulae: [1, s.sScale], showErrorMessage: true, error: `1~${s.sScale} 사이 숫자를 입력하세요.` };
      ws.getCell(r, ci.due).numFmt = "yyyy-mm-dd";
      ws.getCell(r, ci.done).dataValidation = { type: "list", allowBlank: true, formulae: ['"완료,진행중,미착수"'] };
    }

    // 판정에 따라 색이 바뀌는 조건부 서식
    const L = colLetter(ci.judge);
    let pri = 1;
    const rule = (text, argb, fontArgb) => ({ type: "expression", formulae: ["$" + L + r1 + '="' + text + '"'], priority: pri++,
      style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb }, bgColor: { argb } }, font: { color: { argb: fontArgb }, bold: true } } });
    ws.addConditionalFormatting({ ref: `${L}${r1}:${L}${r2}`, rules: [
      rule("보완필요", "FFFFC7CE", "FF9C0006"), rule("보완완료", "FFDDEBF7", "FF1F4E79"),
      rule("적정", "FFC6EFCE", "FF006100"), rule("해당없음", "FFEDEDED", "FF595959"),
    ] });

    // 인쇄 영역이 지정돼 있으면 작성란까지 넓힘
    const pa = ws.pageSetup && ws.pageSetup.printArea;
    if (pa && /:\$?[A-Z]+\$?\d+$/.test(pa)) {
      ws.pageSetup.printArea = pa.replace(/:(\$?)([A-Z]+)(\$?\d+)$/, (m, d, col, rest) => `:${d}${colLetter(cEnd)}${rest}`);
    }
    const layout = { existing: false, ci, r1, r2, c0, cEnd };
    addHelperCols(ws, res, layout);
    return layout;
  }

  // 숨김 보조 열: 보완필요/보완완료 항목 순번, 병합 셀을 채운 작업내용·평가구분
  function addHelperCols(ws, res, layout) {
    const { ci, r1, r2, cEnd } = layout;
    const h0 = cEnd + 1;
    layout.helper = { seq: h0, proc: h0 + 1, div: h0 + 2 };
    const already = ws.getCell(res.dataStartRow - 1, h0).value === HELPER_COLS[0];
    HELPER_COLS.forEach((t, i) => {
      ws.getCell(res.dataStartRow - 1, h0 + i).value = t;
      ws.getColumn(h0 + i).hidden = true;
    });
    if (already) return;
    const J = colLetter(ci.judge);
    for (const it of res.items) {
      const r = it.rowNum;
      ws.getCell(r, h0).value = { formula: `IF(OR($${J}${r}="보완필요",$${J}${r}="보완완료"),COUNTIF($${J}$${r1}:$${J}${r},"보완필요")+COUNTIF($${J}$${r1}:$${J}${r},"보완완료"),"")` };
      ws.getCell(r, h0 + 1).value = it.process || null;
      ws.getCell(r, h0 + 2).value = it.division || null;
    }
    void r2;
  }

  // '보완결과' 시트: 평가표에서 보완필요(또는 보완완료)로 고른 항목이 자동으로 올라오고, 조치 결과를 입력
  function addResultSheet(wb, ws, res, layout) {
    if (wb.getWorksheet("보완결과")) return false;
    const rs = wb.addWorksheet("보완결과", { views: [{ state: "frozen", ySplit: 5 }] });
    const q = "'" + ws.name.replace(/'/g, "''") + "'!";
    const colRef = (c) => `${q}$${colLetter(c)}:$${colLetter(c)}`;
    const { ci, r1, r2, helper } = layout;
    const seqRange = `${q}$${colLetter(helper.seq)}$${r1}:$${colLetter(helper.seq)}$${r2}`;
    const src = (k) => (res.columns[k] ? res.columns[k].col + 1 : 0);
    const n = res.items.length;
    const s = res.settings;

    const cols = [
      { key: "no", title: "No", width: 5, kind: "pull" },
      { key: "row", title: "원본 행", width: 7, kind: "pull" },
      { key: "proc", title: "작업내용", width: 14, kind: "pull", from: helper.proc, text: true },
      { key: "div", title: "평가구분", width: 8, kind: "pull", from: helper.div, text: true },
      { key: "hazard", title: "위험요인 및 재해형태", width: 36, kind: "pull", from: src("hazard"), text: true },
      { key: "risk", title: "현재 위험도", width: 8, kind: "pull", from: src("risk") },
      { key: "judge", title: "판정", width: 9, kind: "pull", from: ci.judge, text: true },
      { key: "issue", title: "보완 필요 사항", width: 26, kind: "pull", from: ci.issue, text: true },
      { key: "measure", title: "보완 대책", width: 26, kind: "pull", from: ci.measure, text: true },
      { key: "owner", title: "담당자", width: 9, kind: "pull", from: ci.owner, text: true },
      { key: "due", title: "완료 예정일", width: 11, kind: "pull", from: ci.due, date: true },
      { key: "result", title: "조치 결과 (실시 내용)", width: 34, kind: "input" },
      { key: "rdate", title: "조치 일자", width: 11, kind: "input" },
      { key: "rf", title: "조치후 빈도", width: 9, kind: "input" },
      { key: "rsv", title: "조치후 강도", width: 9, kind: "input" },
      { key: "rr", title: "조치후 위험도", width: 10, kind: "auto" },
      { key: "checker", title: "확인자", width: 9, kind: "input" },
      { key: "complete", title: "완료 여부", width: 9, kind: "input" },
      { key: "memo", title: "비고", width: 14, kind: "input" },
    ].filter((c) => c.kind !== "pull" || c.key === "no" || c.key === "row" || c.from);
    const k = Object.fromEntries(cols.map((c, i) => [c.key, i + 1]));
    const last = cols.length;
    cols.forEach((c, i) => { rs.getColumn(i + 1).width = c.width; });

    // 제목·안내·진행 현황
    rs.mergeCells(1, 1, 1, last);
    Object.assign(rs.getCell(1, 1), { value: "위험성평가 보완결과" });
    rs.getCell(1, 1).font = { bold: true, size: 16 };
    rs.getCell(1, 1).alignment = { horizontal: "center", vertical: "middle" };
    rs.getRow(1).height = 28;
    rs.mergeCells(2, 1, 2, last);
    rs.getCell(2, 1).value = `※ ‘${ws.name}’ 시트의 ‘유효성 검토’에서 보완필요(또는 보완완료)를 고르면 이 시트에 자동으로 나타납니다. 노란 칸에 조치 결과를 입력하세요. ` +
      "판정을 먼저 정한 뒤 결과를 입력하는 것을 권장합니다 (나중에 위쪽 행을 보완필요로 바꾸면 목록 순서가 한 칸씩 밀립니다).";
    rs.getCell(2, 1).font = { size: 9, color: { argb: "FF7F7F7F" } };
    rs.getCell(2, 1).alignment = { wrapText: true, vertical: "middle" };
    rs.getRow(2).height = 28;

    const R1 = 6, R2 = 5 + n;
    const L = (key) => colLetter(k[key]);
    const rr = (key) => `${L(key)}${R1}:${L(key)}${R2}`;
    rs.mergeCells(3, 1, 3, last);
    const p = rs.getCell(3, 1);
    p.value = { formula: `"보완 대상 "&COUNT(${rr("row")})&"건   ·   조치 완료 "&COUNTIF(${rr("complete")},"완료")&"건   ·   미완료 "&(COUNT(${rr("row")})-COUNTIF(${rr("complete")},"완료"))&"건"` };
    p.font = { bold: true, size: 11, color: { argb: "FF7F3F00" } };
    p.fill = solid("FFFFF2CC");
    p.alignment = { horizontal: "center", vertical: "middle" };
    rs.getRow(3).height = 22;

    // 머리글
    const firstInput = k.result;
    const groups = [
      { from: 1, to: firstInput - 1, title: "평가표에서 자동으로 가져옴", argb: C.copyHead },
      { from: firstInput, to: last, title: "보완 결과 (직접 입력)", argb: C.inputHead },
    ];
    for (const g of groups) {
      rs.mergeCells(4, g.from, 4, g.to);
      const c = rs.getCell(4, g.from);
      c.value = g.title; c.fill = solid(g.argb);
      c.font = { bold: true, color: { argb: "FFFFFFFF" } };
      c.alignment = { horizontal: "center", vertical: "middle" };
      c.border = border;
    }
    cols.forEach((col, i) => {
      const c = rs.getCell(5, i + 1);
      c.value = col.title;
      c.fill = solid(col.kind === "pull" ? C.copyHead : col.kind === "auto" ? C.autoHead : "FFE26B0A");
      c.font = { bold: true, size: 10, color: { argb: "FFFFFFFF" } };
      c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      c.border = border;
    });
    rs.getRow(5).height = 30;

    // 데이터 행 (평가 항목 수만큼 미리 수식을 넣어 둠)
    const light = { style: "thin", color: { argb: "FFD0D0D0" } };
    const lightBorder = { top: light, left: light, bottom: light, right: light };
    const plus = s.method === "add" ? "+" : "*";
    for (let i = 0; i < n; i++) {
      const r = R1 + i;
      const B = `$${L("row")}${r}`;
      rs.getCell(r, k.row).value = { formula: `IFERROR(MATCH(${i + 1},${seqRange},0)+${r1 - 1},"")` };
      rs.getCell(r, k.no).value = { formula: `IF(${B}="","",${i + 1})` };
      for (const col of cols) {
        if (col.kind !== "pull" || !col.from) continue;
        const ref = `INDEX(${colRef(col.from)},${B})`;
        rs.getCell(r, k[col.key]).value = { formula: col.text ? `IF(${B}="","",${ref}&"")` : `IF(${B}="","",IF(${ref}="","",${ref}))` };
        if (col.date) rs.getCell(r, k[col.key]).numFmt = "yyyy-mm-dd";
      }
      rs.getCell(r, k.rr).value = { formula: `IF(AND(ISNUMBER(${L("rf")}${r}),ISNUMBER(${L("rsv")}${r})),${L("rf")}${r}${plus}${L("rsv")}${r},"")` };
      rs.getCell(r, k.rdate).numFmt = "yyyy-mm-dd";
      rs.getCell(r, k.rf).dataValidation = { type: "whole", operator: "between", allowBlank: true, formulae: [1, s.fScale], showErrorMessage: true, error: `1~${s.fScale} 사이 숫자를 입력하세요.` };
      rs.getCell(r, k.rsv).dataValidation = { type: "whole", operator: "between", allowBlank: true, formulae: [1, s.sScale], showErrorMessage: true, error: `1~${s.sScale} 사이 숫자를 입력하세요.` };
      rs.getCell(r, k.complete).dataValidation = { type: "list", allowBlank: true, formulae: ['"완료,진행중,미착수"'] };
      cols.forEach((col, j) => {
        const c = rs.getCell(r, j + 1);
        c.border = lightBorder;
        c.font = { size: 10 };
        const left = ["hazard", "issue", "measure", "result", "memo", "proc"].includes(col.key);
        c.alignment = { vertical: "middle", horizontal: left ? "left" : "center", wrapText: true };
      });
      rs.getRow(r).height = 30;
    }

    // 항목이 올라온 행만 색칠 (평가표에서 가져온 칸 = 회색, 입력 칸 = 노랑), 완료 여부 색
    const has = `$${L("row")}${R1}<>""`;
    let pri = 1;
    const cf = (ref, formula, argb, extra = {}) => rs.addConditionalFormatting({ ref, rules: [{ type: "expression", formulae: [formula], priority: pri++,
      style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb }, bgColor: { argb } }, ...extra } }] });
    cf(`${L("complete")}${R1}:${L("complete")}${R2}`, `$${L("complete")}${R1}="완료"`, "FFC6EFCE", { font: { color: { argb: "FF006100" }, bold: true } });
    cf(`${L("judge")}${R1}:${L("judge")}${R2}`, `$${L("judge")}${R1}="보완필요"`, "FFFFC7CE", { font: { color: { argb: "FF9C0006" }, bold: true } });
    cf(`${L("judge")}${R1}:${L("judge")}${R2}`, `$${L("judge")}${R1}="보완완료"`, "FFDDEBF7", { font: { color: { argb: "FF1F4E79" }, bold: true } });
    cf(`A${R1}:${colLetter(firstInput - 1)}${R2}`, has, "FFF2F2F2");
    cf(`${colLetter(firstInput)}${R1}:${colLetter(last)}${R2}`, has, C.input);

    rs.pageSetup = { paperSize: 8, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 }, printTitlesRow: "4:5" };

    // 평가표 시트 바로 뒤에 배치
    const others = wb.worksheets.filter((w) => w !== rs);
    const idx = others.indexOf(ws) + 1;
    [...others.slice(0, idx), rs, ...others.slice(idx)].forEach((w, i) => { w.orderNo = i; });
    return true;
  }

  function download(buf, filename) {
    const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  }

  // .xlsx/.xlsm 은 원본 그대로 열고, .xls/.csv 는 값으로 새 통합문서를 만든다
  async function openWorkbook() {
    if (/\.(xlsx|xlsm)$/i.test(state.fileName)) {
      try {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(state.fileBuf.slice(0));
        return { wb, preserved: true };
      } catch (ex) {
        console.warn("원본 열기 실패, 값으로 대체", ex);
      }
    }
    const wb = new ExcelJS.Workbook();
    for (const n of state.wb.SheetNames) {
      const ws = wb.addWorksheet(n.slice(0, 31));
      const { ws: src, rows } = sheetRows(n);
      rows.forEach((r, i) => r.forEach((v, c) => { if (v !== "" && v != null) ws.getCell(i + 1, c + 1).value = v; }));
      for (const m of src["!merges"] || []) { try { ws.mergeCells(m.s.r + 1, m.s.c + 1, m.e.r + 1, m.e.c + 1); } catch {} }
    }
    return { wb, preserved: false };
  }
})();
