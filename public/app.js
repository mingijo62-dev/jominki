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
    $("copiedCols").textContent = fields.map((f) => f.title).join(", ");
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
      const name = uniqueName(wb, "보완목록");
      const ws = wb.addWorksheet(name, { views: [{ state: "frozen", ySplit: 6, xSplit: 0 }] });
      buildList(ws, res);
      // 보완목록을 평가표 시트 바로 뒤에 두고, 열었을 때 보완목록이 보이게
      const target = wb.getWorksheet(state.sheet);
      const others = wb.worksheets.filter((w) => w !== ws);
      const idx = target ? others.indexOf(target) + 1 : others.length;
      const order = [...others.slice(0, idx), ws, ...others.slice(idx)];
      order.forEach((w, i) => { w.orderNo = i; if (w.views && w.views[0]) w.views[0].tabSelected = w === ws; });
      wb.views = [{ activeTab: order.indexOf(ws), firstSheet: 0, visibility: "visible" }];
      const buf = await wb.xlsx.writeBuffer();
      const base = state.fileName.replace(/\.[^.]+$/, "");
      download(buf, `${base}_보완목록.xlsx`);
      if (!preserved) alert("원본 형식(.xls/.csv)은 그대로 수정할 수 없어, 값만 옮긴 새 .xlsx 파일로 만들었습니다.");
    } catch (ex) {
      console.error(ex);
      alert("파일을 만들지 못했습니다: " + ex.message);
    } finally {
      btn.disabled = false; btn.textContent = label;
    }
  });

  function uniqueName(wb, base) {
    let n = base, i = 2;
    while (wb.getWorksheet(n)) n = `${base}(${i++})`;
    return n;
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

  function buildList(ws, res) {
    const fields = copyFieldsFor(res);
    const s = res.settings;
    const cols = [
      { key: "no", title: "No", width: 5, kind: "copy" },
      { key: "row", title: "원본\n행", width: 6, kind: "copy" },
      ...fields.map((f) => ({ ...f, kind: "copy" })),
      { key: "judge", title: "유효성 검토\n(판정)", width: 11, kind: "input" },
      { key: "issue", title: "보완 필요 사항\n(검토 의견)", width: 34, kind: "input" },
      { key: "measure", title: "보완 대책", width: 34, kind: "input" },
      { key: "af", title: "보완후\n빈도", width: 7, kind: "input" },
      { key: "as", title: "보완후\n강도", width: 7, kind: "input" },
      { key: "ar", title: "보완후\n위험도", width: 7, kind: "auto" },
      { key: "owner", title: "담당자", width: 10, kind: "input" },
      { key: "due", title: "완료\n예정일", width: 11, kind: "input" },
      { key: "done", title: "완료\n확인", width: 8, kind: "input" },
      { key: "memo", title: "비고", width: 16, kind: "input" },
    ];
    const ci = Object.fromEntries(cols.map((c, i) => [c.key, i + 1]));
    const last = cols.length;
    cols.forEach((c, i) => { ws.getColumn(i + 1).width = c.width; });

    // 1행: 제목, 2~3행: 정보·결재, 4행: 진행 현황, 5~6행: 머리글
    ws.mergeCells(1, 1, 1, last);
    const t = ws.getCell(1, 1);
    t.value = "위험성평가 보완목록 (유효성 검토)";
    t.font = { bold: true, size: 16 };
    t.alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(1).height = 30;

    const info = [
      ["대상 파일", state.fileName], ["평가표 시트", state.sheet],
      ["평가일자", res.meta.evalDate || ""], ["항목 수", res.items.length],
    ];
    ws.getCell(2, 1).value = info.map(([k, v]) => `${k}: ${v}`).join("    ");
    ws.mergeCells(2, 1, 2, Math.max(3, ci.judge - 1));
    ws.getCell(2, 1).font = { size: 10, color: { argb: "FF404040" } };
    ws.getCell(3, 1).value = "※ 노란 칸에 직접 작성합니다. ‘유효성 검토’는 목록에서 선택(적정/보완필요/해당없음/보완완료), 보완후 위험도와 진행 현황은 자동 계산됩니다.";
    ws.mergeCells(3, 1, 3, Math.max(3, ci.judge - 1));
    ws.getCell(3, 1).font = { size: 9, color: { argb: "FF7F7F7F" } };

    // 결재란 (오른쪽 위)
    const signs = ["작성", "검토", "승인"];
    const signStart = last - signs.length + 1;
    signs.forEach((sg, i) => {
      const h = ws.getCell(2, signStart + i), b = ws.getCell(3, signStart + i);
      h.value = sg; h.alignment = { horizontal: "center", vertical: "middle" }; h.fill = solid("FFE7E6E6"); h.font = { bold: true, size: 9 };
      h.border = border; b.border = border;
    });
    ws.getRow(3).height = 34;

    const r1 = 7, r2 = 6 + res.items.length;
    const rng = (k) => `${colLetter(ci[k])}${r1}:${colLetter(ci[k])}${r2}`;
    const stat = [
      ["전체", `COUNTA(${rng("row")})`],
      ["검토완료", `COUNTA(${rng("judge")})`],
      ["보완필요", `COUNTIF(${rng("judge")},"보완필요")`],
      ["보완완료", `COUNTIF(${rng("judge")},"보완완료")`],
      ["미검토", `COUNTA(${rng("row")})-COUNTA(${rng("judge")})`],
    ];
    // 진행 현황: 한 칸에 모아서 표시 (작성하면 자동 갱신)
    ws.mergeCells(4, ci.judge, 4, last);
    const prog = ws.getCell(4, ci.judge);
    prog.value = { formula: '"진행 현황  ▶  " & ' + stat.map(([k, f]) => `"${k} "&${f}`).join(' & "   ·   " & ') };
    prog.font = { bold: true, size: 10, color: { argb: "FF7F3F00" } };
    prog.fill = solid("FFFFF2CC");
    prog.alignment = { horizontal: "center", vertical: "middle" };
    prog.border = border;
    ws.getRow(4).height = 20;

    // 머리글 (5행: 묶음, 6행: 항목)
    const groups = [
      { from: 1, to: ci.judge - 1, title: "원본 평가 내용", argb: C.copyHead },
      { from: ci.judge, to: last, title: "유효성 검토 및 보완 (직접 작성)", argb: C.inputHead },
    ];
    for (const g of groups) {
      ws.mergeCells(5, g.from, 5, g.to);
      const cell = ws.getCell(5, g.from);
      cell.value = g.title;
      cell.fill = solid(g.argb);
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = border;
    }
    cols.forEach((col, i) => {
      const cell = ws.getCell(6, i + 1);
      cell.value = col.title;
      cell.fill = solid(col.kind === "copy" ? C.copyHead : col.kind === "auto" ? C.autoHead : C.inputHead);
      cell.font = { bold: true, size: 10, color: { argb: "FFFFFFFF" } };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.border = border;
    });
    ws.getRow(5).height = 20;
    ws.getRow(6).height = 32;

    // 항목
    const plus = s.method === "add" ? "+" : "*";
    res.items.forEach((it, idx) => {
      const r = r1 + idx;
      const row = ws.getRow(r);
      cols.forEach((col, i) => {
        const cell = row.getCell(i + 1);
        cell.border = border;
        cell.font = { size: 10 };
        const left = ["hazard", "current", "issue", "measure", "memo", "process"].includes(col.key);
        cell.alignment = { vertical: "middle", horizontal: left ? "left" : "center", wrapText: true };
        if (col.kind === "input") cell.fill = solid(C.input);
        else if (col.kind === "copy") cell.fill = solid(C.copy);
      });
      row.getCell(ci.no).value = idx + 1;
      row.getCell(ci.row).value = it.rowNum;
      for (const f of fields) {
        const v = it[f.prop];
        row.getCell(ci[f.key]).value = v === "" || v == null ? null : v;
      }
      const fL = colLetter(ci.af), sL = colLetter(ci.as);
      row.getCell(ci.ar).value = { formula: `IF(AND(ISNUMBER(${fL}${r}),ISNUMBER(${sL}${r})),${fL}${r}${plus}${sL}${r},"")` };
      row.getCell(ci.judge).dataValidation = { type: "list", allowBlank: true, formulae: ['"적정,보완필요,해당없음,보완완료"'], showErrorMessage: true, errorTitle: "유효성 검토", error: "목록에서 선택하세요." };
      row.getCell(ci.af).dataValidation = { type: "whole", operator: "between", allowBlank: true, formulae: [1, s.fScale], showErrorMessage: true, error: `1~${s.fScale} 사이 숫자를 입력하세요.` };
      row.getCell(ci.as).dataValidation = { type: "whole", operator: "between", allowBlank: true, formulae: [1, s.sScale], showErrorMessage: true, error: `1~${s.sScale} 사이 숫자를 입력하세요.` };
      row.getCell(ci.due).numFmt = "yyyy-mm-dd";
      row.getCell(ci.done).dataValidation = { type: "list", allowBlank: true, formulae: ['"완료,진행중,미착수"'] };
      // 줄바꿈되는 글자 수를 대략 계산해 행 높이 지정
      const lines = (v, w) => String(v || "").split("\n").reduce((n, part) => n + Math.max(1, Math.ceil(part.length / w)), 0);
      const widthOf = (k) => (cols[ci[k] - 1] ? cols[ci[k] - 1].width * 0.55 : 20);
      const n = Math.max(
        ci.hazard ? lines(it.hazard, widthOf("hazard")) : 1,
        ci.current ? lines(it.current, widthOf("current")) : 1,
        ci.process ? lines(it.process, widthOf("process")) : 1,
      );
      row.height = Math.max(20, n * 14 + 4);
    });

    // 판정에 따라 색이 바뀌는 조건부 서식
    const L = colLetter(ci.judge);
    let pri = 1;
    const rule = (text, argb, fontArgb) => ({ type: "expression", formulae: ["$" + L + r1 + '="' + text + '"'], priority: pri++,
      style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb }, bgColor: { argb } }, font: { color: { argb: fontArgb }, bold: true } } });
    ws.addConditionalFormatting({ ref: `${L}${r1}:${L}${r2}`, rules: [
      rule("보완필요", "FFFFC7CE", "FF9C0006"), rule("보완완료", "FFDDEBF7", "FF1F4E79"),
      rule("적정", "FFC6EFCE", "FF006100"), rule("해당없음", "FFEDEDED", "FF595959"),
    ] });

    // 인쇄 설정: A3 가로, 너비 한 페이지, 머리글 반복
    ws.pageSetup = { paperSize: 8, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 }, printTitlesRow: "5:6" };
    ws.autoFilter = { from: { row: 6, column: 1 }, to: { row: r2, column: last } };
  }
})();
