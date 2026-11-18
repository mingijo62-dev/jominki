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
    return { ws, rows: XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "", blankrows: true }) };
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

  // ---------- 엑셀 다운로드 ----------
  $("dlBtn").addEventListener("click", () => {
    const res = state.result;
    if (!res || !res.ok) return;
    const byRow = new Map();
    for (const f of res.findings) if (f.row) { if (!byRow.has(f.row)) byRow.set(f.row, []); byRow.get(f.row).push(f); }

    const summary = [
      ["위험성평가 재검토 결과"],
      [],
      ["파일", state.fileName], ["시트", state.sheet],
      ["검토일", new Date().toISOString().slice(0, 10)],
      ["점수", res.score], ["종합 판정", res.verdict],
      ["오류", res.counts.error], ["경고", res.counts.warn], ["참고", res.counts.info],
      ["평가 기준", `${res.settings.fScale}×${res.settings.sScale}, ${res.settings.method === "mul" ? "곱셈" : "덧셈"}, 허용 가능 ≤ ${res.settings.acceptable}`],
      ["평가일자", res.meta.evalDate || "미확인"],
      [],
      ["※ 본 결과는 자동 검토 결과로 참고용입니다."],
    ];
    const findings = [["수준", "행", "코드", "내용"], ...res.findings.map((f) => [LEVEL_KO[f.level], f.row || "전체", f.code, f.msg])];

    // 원본 시트 + 검토결과 열 추가
    const { rows } = sheetRows(state.sheet);
    const width = Math.max(...rows.map((r) => r.length));
    const annotated = rows.map((r, i) => {
      const out = r.slice();
      while (out.length < width) out.push("");
      if (i === res.headerRow - 1) out.push("검토결과");
      else {
        const fs = byRow.get(i + 1);
        out.push(fs ? fs.map((f) => `[${LEVEL_KO[f.level]}] ${f.msg}`).join("\n") : "");
      }
      return out;
    });

    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.aoa_to_sheet(summary); ws1["!cols"] = [{ wch: 14 }, { wch: 60 }];
    const ws2 = XLSX.utils.aoa_to_sheet(findings); ws2["!cols"] = [{ wch: 6 }, { wch: 6 }, { wch: 16 }, { wch: 90 }];
    const ws3 = XLSX.utils.aoa_to_sheet(annotated, { cellDates: true });
    const origMerges = state.wb.Sheets[state.sheet]["!merges"];
    if (origMerges) ws3["!merges"] = origMerges;
    XLSX.utils.book_append_sheet(wb, ws1, "요약");
    XLSX.utils.book_append_sheet(wb, ws2, "지적사항");
    XLSX.utils.book_append_sheet(wb, ws3, "원본+검토결과");
    const base = state.fileName.replace(/\.[^.]+$/, "");
    const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const a = document.createElement("a");
    a.href = url; a.download = `${base}_재검토결과.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
  });

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
