/* 위험성평가표 규칙 기반 유효성 검사기
 * 입력: 2차원 배열(시트 셀 값), 출력: 구조화된 행 + 지적사항
 */
(function (root) {
  "use strict";

  const norm = (v) => String(v ?? "").replace(/\s+/g, "").replace(/[()（）\[\]·.:]/g, "").toLowerCase();
  const text = (v) => String(v ?? "").trim();

  // 컬럼 역할 정의 (우선순위 순)
  const ROLE_PATTERNS = [
    ["no", /^(no|번호|순번|연번|일련번호)$/],
    ["process", /(공정|단위작업|작업명|작업내용|세부작업|작업단계|작업공정|설비명|장소)/],
    ["category", /(분류|재해형태|재해유형|위험분류|유형|원인)/],
    ["hazard", /(유해위험요인|유해위험요인파악|위험요인|유해요인|위험내용|위험상황|잠재위험|유해위험)/],
    ["current", /(현재안전|현재의안전|현재조치|현안전|기존대책|현재대책|현재상태|현행)/],
    ["law", /(관련근거|법적기준|관련법규|법규)/],
    ["freq", /(가능성|빈도)/],
    ["sev", /(중대성|강도|심각도)/],
    ["risk", /(위험성|위험도|위험등급|위험수준|위험성수준|위험성크기|평가점수|위험성결정)/],
    ["measure", /(감소대책|개선대책|개선방안|저감대책|감소방안|대책|개선내용|조치내용|관리방안)/],
    ["owner", /(담당자|담당|책임자|관리자|조치자)/],
    ["due", /(예정일|완료예정|개선일정|조치기한|조치예정|기한|일정|개선예정)/],
    ["done", /(완료일|조치일|조치완료|개선완료|실시일)/],
    ["status", /(진행|상태|완료여부|이행여부|결과)/],
  ];
  const AFTER_RE = /(개선후|감소후|대책후|조치후|저감후|잔류|잔여|개선결과|사후)/;
  const BEFORE_RE = /(개선전|현재위험|감소전|조치전|사전)/;
  const HEADER_HINT = /(유해|위험요인|가능성|빈도|중대성|강도|위험성|위험도|감소대책|개선대책|공정|작업)/;

  function cellStr(rows, r, c) {
    return rows[r] && rows[r][c] != null ? String(rows[r][c]) : "";
  }

  // 헤더 행 찾기: 키워드가 가장 많이 매칭되는 행
  function findHeaderRow(rows) {
    let best = { r: -1, score: 0 };
    const lim = Math.min(rows.length, 40);
    for (let r = 0; r < lim; r++) {
      const row = rows[r] || [];
      let score = 0;
      const seen = new Set();
      for (const v of row) {
        const n = norm(v);
        if (!n || n.length > 30) continue;
        for (const [role, re] of ROLE_PATTERNS) {
          if (re.test(n) && !seen.has(role)) { seen.add(role); score++; break; }
        }
      }
      if (score > best.score) best = { r, score };
    }
    return best.score >= 2 ? best.r : -1;
  }

  // 병합 셀(가로) 상위 라벨을 오른쪽으로 전파
  function fillRight(arr) {
    const out = arr.slice();
    let last = "";
    for (let i = 0; i < out.length; i++) {
      if (text(out[i])) last = out[i]; else out[i] = last;
    }
    return out;
  }

  function buildColumns(rows, hr, merges) {
    const width = Math.max(...rows.slice(hr, hr + 3).map((r) => (r ? r.length : 0)), 0);
    // 서브헤더 여부: 다음 행에 가능성/중대성 같은 짧은 라벨이 다수 있고 숫자가 거의 없으면 헤더로 간주
    const next = rows[hr + 1] || [];
    let subHeader = false;
    {
      let kw = 0, num = 0;
      for (const v of next) {
        const n = norm(v);
        if (!n) continue;
        if (/^-?\d+(\.\d+)?$/.test(n)) num++;
        else if (n.length <= 20 && HEADER_HINT.test(n)) kw++;
      }
      subHeader = kw >= 2 && num <= 1;
    }
    const parentRaw = [];
    for (let c = 0; c < width; c++) parentRaw.push(cellStr(rows, hr, c));
    // merges 기반 상위 라벨 전파 (merges가 없으면 빈칸을 상위로 보고 오른쪽 전파)
    let parent = parentRaw.slice();
    if (merges && merges.length) {
      for (const m of merges) {
        if (m.s.r <= hr && m.e.r >= hr) {
          const v = cellStr(rows, m.s.r, m.s.c);
          for (let c = m.s.c; c <= m.e.c; c++) if (!text(parent[c])) parent[c] = v;
        }
      }
    } else if (subHeader) {
      parent = fillRight(parentRaw);
    }
    const labels = [];
    for (let c = 0; c < width; c++) {
      const p = text(parent[c]);
      const s = subHeader ? text(cellStr(rows, hr + 1, c)) : "";
      labels.push({ c, parent: p, sub: s, full: s && s !== p ? `${p} ${s}`.trim() : p });
    }
    return { labels, dataStart: hr + (subHeader ? 2 : 1) };
  }

  function classify(labels) {
    const map = {}; // role -> col
    const seenCount = {};
    for (const L of labels) {
      const nSub = norm(L.sub);
      const nParent = norm(L.parent);
      const nFull = norm(L.full);
      if (!nFull) continue;
      const target = nSub || nParent;
      let role = null;
      for (const [r, re] of ROLE_PATTERNS) {
        if (re.test(target)) { role = r; break; }
      }
      // 서브 라벨로 못 찾으면 상위 라벨로
      if (!role && nSub) {
        for (const [r, re] of ROLE_PATTERNS) if (re.test(nParent)) { role = r; break; }
      }
      if (!role) continue;
      const isAfter = AFTER_RE.test(nFull) && !BEFORE_RE.test(nSub);
      const scorable = role === "freq" || role === "sev" || role === "risk";
      let key = role;
      if (scorable) {
        seenCount[role] = (seenCount[role] || 0) + 1;
        if (isAfter || seenCount[role] >= 2) key = "after_" + role;
      } else if (role === "measure" && map.measure != null) {
        // 두번째 대책 열은 무시(첫번째 유지) — 단, 상위 라벨이 '감소대책'인 경우 우선
        if (/감소|개선/.test(nFull) && !/감소|개선/.test(norm(labels[map.measure].full))) map.measure = L.c;
        continue;
      } else if (role === "hazard" && map.hazard != null) {
        continue;
      }
      if (map[key] == null) map[key] = L.c;
    }
    return map;
  }

  function toNum(v) {
    if (v == null || v === "") return null;
    if (typeof v === "number") return v;
    const s = String(v).trim();
    const m = s.match(/^-?\d+(\.\d+)?/);
    return m ? Number(m[0]) : null;
  }

  // 위험성 값이 등급(상/중/하, 높음/보통/낮음 등)인 경우 레벨로 변환
  const GRADE_MAP = [
    [/(매우높음|매우큼|최상|very\s*high|^5$)/i, 5],
    [/(높음|높다|상|high|^h$|큼|위험|허용불가)/i, 3],
    [/(보통|중|medium|^m$|중간)/i, 2],
    [/(낮음|낮다|하|low|^l$|작음|허용가능|경미)/i, 1],
  ];
  function gradeOf(v) {
    const n = norm(v);
    if (!n) return null;
    for (const [re, g] of GRADE_MAP) if (re.test(n)) return g;
    return null;
  }

  function toDate(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date && !isNaN(v)) return v;
    if (typeof v === "number" && v > 20000 && v < 80000) {
      return new Date(Math.round((v - 25569) * 86400 * 1000));
    }
    const s = String(v).trim();
    let m = s.match(/(20\d{2}|19\d{2})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    m = s.match(/^(\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})$/);
    if (m) return new Date(2000 + +m[1], +m[2] - 1, +m[3]);
    m = s.match(/(20\d{2})[.\-/년\s]+(\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, 1);
    return null;
  }
  const fmtDate = (d) => (d ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}` : "");

  const WEAK_MEASURE = /(주의|철저|준수|유의|숙지|당부|강조|경각심|조심|안전의식|생활화|습관화|확인할것|교육실시|교육강화|안전교육|tbm|주지)/;
  const STRONG_MEASURE = /(설치|교체|제거|대체|방호|덮개|울|난간|가드|인터록|인터락|센서|자동|격리|차단|국소배기|환기|개선|보강|고정|개조|변경|잠금|loto|접지|누전|안전장치|비상정지|방책|칸막이|도입|구매|보수|수리|배치|정비|표지)/;
  const PPE_ONLY = /(보호구|안전모|안전화|장갑|보안경|마스크|귀마개|안전대|방독|방진|착용)/;

  /**
   * @param {any[][]} rows 시트 데이터(header:1)
   * @param {object} opts { merges, acceptable, today }
   */
  function analyze(rows, opts = {}) {
    const today = opts.today ? new Date(opts.today) : new Date();
    const findings = []; // {level:'error'|'warn'|'info', row, col, code, msg}
    const add = (level, row, code, msg, field) => findings.push({ level, row, code, msg, field });

    const hr = findHeaderRow(rows);
    if (hr < 0) {
      return {
        ok: false,
        error: "위험성평가표의 헤더(유해·위험요인, 가능성, 중대성, 위험성, 감소대책 등)를 찾지 못했습니다. 다른 시트를 선택하거나 양식을 확인해 주세요.",
      };
    }
    const { labels, dataStart } = buildColumns(rows, hr, opts.merges);
    const cols = classify(labels);

    // 메타(평가일, 근로자 참여 등) — 헤더 위/아래 전체 텍스트 검색
    const allText = rows.slice(0, 200).map((r) => (r || []).map(text).join(" ")).join("\n");
    const preText = rows.slice(0, hr).map((r) => (r || []).map(text).join(" ")).join("\n");
    const meta = {};
    {
      const m = preText.match(/(평가\s*일자?|평가\s*일시|실시\s*일자?|작성\s*일자?|평가\s*기간)[^\d]{0,10}((?:19|20)\d{2}[.\-/년\s]+\d{1,2}[.\-/월\s]+\d{1,2})/);
      let d = m ? toDate(m[2]) : null;
      if (!d) {
        // 헤더 위 영역 숫자 날짜 셀
        for (const r of rows.slice(0, hr)) for (const v of r || []) {
          if (!d && (v instanceof Date || (typeof v === "number" && v > 36000 && v < 60000))) d = toDate(v);
        }
      }
      meta.evalDate = d ? fmtDate(d) : "";
      meta.workerParticipation = /(근로자|작업자|현장)\s*(참여|의견|서명|대표)|참여자|참석자|청취/.test(allText);
      meta.typeText = (preText.match(/(최초|정기|수시|상시)\s*(평가|위험성평가)?/) || [""])[0];
      meta.methodText = (allText.match(/(빈도\s*[·x×*]\s*강도|3\s*[x×]\s*3|4\s*[x×]\s*4|5\s*[x×]\s*5|체크리스트|3단계\s*판단|핵심요인\s*기술|OPS|위험성\s*수준\s*3단계)/i) || [""])[0];
    }

    // 데이터 행 수집
    const items = [];
    let lastProcess = "";
    for (let r = dataStart; r < rows.length; r++) {
      const row = rows[r] || [];
      const get = (k) => (cols[k] != null ? row[cols[k]] : undefined);
      const hazard = text(get("hazard"));
      const measure = text(get("measure"));
      const nonEmpty = row.filter((v) => text(v)).length;
      if (nonEmpty === 0) continue;
      // 서명란/비고 등 꼬리 영역 판별: 핵심 컬럼이 모두 비어있고 숫자도 없음
      const f = toNum(get("freq")), s = toNum(get("sev")), rk = get("risk");
      const noScore = f == null && s == null && !text(rk);
      if (noScore && !measure && (!hazard || nonEmpty <= 3)) continue;
      // 서명/결재/비고 꼬리 영역
      if (noScore && /(작성자|검토자|승인|결재|서명|확인자|참여자|비고|※)/.test(row.map(text).join(" "))) continue;
      // 반복 헤더 행 건너뛰기
      if (HEADER_HINT.test(norm(hazard)) && /(유해|위험요인)/.test(norm(hazard)) && f == null) continue;
      let process = text(get("process"));
      if (process) lastProcess = process; else process = lastProcess;
      items.push({
        rowNum: r + 1, // 엑셀 행 번호
        no: text(get("no")),
        process,
        category: text(get("category")),
        hazard,
        current: text(get("current")),
        freq: f, sev: s,
        riskRaw: get("risk"), risk: toNum(get("risk")), riskGrade: gradeOf(get("risk")),
        measure,
        afterFreq: toNum(get("after_freq")), afterSev: toNum(get("after_sev")),
        afterRiskRaw: get("after_risk"), afterRisk: toNum(get("after_risk")), afterGrade: gradeOf(get("after_risk")),
        owner: text(get("owner")),
        dueRaw: get("due"), due: toDate(get("due")),
        doneRaw: get("done"), done: toDate(get("done")),
        status: text(get("status")),
      });
    }

    // --- 구조 점검 ---
    const has = (k) => cols[k] != null;
    if (!has("hazard")) add("error", null, "COL_HAZARD", "‘유해·위험요인’ 열을 찾지 못했습니다. 위험성평가의 핵심 항목입니다.");
    if (!has("freq") && !has("sev") && !has("risk")) add("error", null, "COL_RISK", "위험성 추정 열(가능성·중대성·위험성)을 찾지 못했습니다.");
    else {
      if (!has("freq")) add("warn", null, "COL_FREQ", "‘가능성(빈도)’ 열이 없습니다. 위험성 추정 근거를 확인할 수 없습니다.");
      if (!has("sev")) add("warn", null, "COL_SEV", "‘중대성(강도)’ 열이 없습니다. 위험성 추정 근거를 확인할 수 없습니다.");
      if (!has("risk")) add("info", null, "COL_RISKV", "‘위험성(위험도)’ 결과 열이 없습니다. 위험성 결정 결과를 명시하는 것이 좋습니다.");
    }
    if (!has("measure")) add("error", null, "COL_MEASURE", "‘감소대책(개선대책)’ 열을 찾지 못했습니다.");
    if (!has("after_risk") && !has("after_freq")) add("warn", null, "COL_AFTER", "‘개선 후 위험성’ 열이 없습니다. 감소대책의 효과(잔류 위험성)를 확인할 수 없습니다.");
    if (!has("owner")) add("warn", null, "COL_OWNER", "‘담당자’ 열이 없습니다. 감소대책 이행 책임자를 지정해야 합니다.");
    if (!has("due") && !has("done")) add("warn", null, "COL_DUE", "‘개선 예정일/완료일’ 열이 없습니다. 이행 일정을 관리할 수 없습니다.");

    if (!meta.evalDate) add("info", null, "META_DATE", "평가 실시일자를 찾지 못했습니다. 평가일자를 기록해 두어야 정기평가 주기(매년)를 관리할 수 있습니다.");
    else {
      const d = toDate(meta.evalDate);
      const days = (today - d) / 86400000;
      if (days > 365) add("warn", null, "META_OLD", `평가일(${meta.evalDate})로부터 1년이 지났습니다. 정기 위험성평가(매년)를 다시 실시해야 합니다.`);
    }
    if (!meta.workerParticipation) add("info", null, "META_WORKER", "근로자 참여(의견 청취·참여자 서명 등) 기록을 찾지 못했습니다. 위험성평가 전 과정에 근로자 참여가 필요합니다.");

    if (items.length === 0) {
      add("error", null, "NO_ROWS", "평가 항목(데이터 행)이 없습니다.");
      return finalize();
    }

    // --- 척도/계산 방식 추정 ---
    const maxOf = (arr) => arr.filter((x) => x != null).reduce((a, b) => Math.max(a, b), 0);
    const fMax = maxOf(items.map((i) => i.freq).concat(items.map((i) => i.afterFreq)));
    const sMax = maxOf(items.map((i) => i.sev).concat(items.map((i) => i.afterSev)));
    const scaleOf = (m) => (m <= 3 ? 3 : m <= 4 ? 4 : m <= 5 ? 5 : m);
    const fScale = opts.freqScale || scaleOf(fMax || 3);
    const sScale = opts.sevScale || scaleOf(sMax || 3);
    let mul = 0, addc = 0;
    for (const i of items) {
      if (i.freq != null && i.sev != null && i.risk != null) {
        if (Math.abs(i.freq * i.sev - i.risk) < 1e-9) mul++;
        if (Math.abs(i.freq + i.sev - i.risk) < 1e-9) addc++;
      }
    }
    const method = opts.method || (addc > mul ? "add" : "mul");
    const calc = (f, s) => (f == null || s == null ? null : method === "add" ? f + s : f * s);
    const maxRisk = calc(fScale, sScale);
    const defaultAcceptable = method === "add"
      ? Math.round(maxRisk * 0.5)
      : ({ 9: 2, 16: 4, 25: 6, 12: 3, 15: 4, 20: 5 }[maxRisk] ?? Math.round(maxRisk * 0.25));
    const acceptable = opts.acceptable != null ? Number(opts.acceptable) : defaultAcceptable;

    // 위험성 수준 결정 (숫자 또는 등급)
    const levelOf = (num, grade) => {
      if (num != null) return num > acceptable ? "high" : "ok";
      if (grade != null) return grade >= 2 ? "high" : "ok";
      return null;
    };

    // --- 행 단위 점검 ---
    const dupKey = new Map();
    for (const i of items) {
      const R = i.rowNum;
      const iss = (level, code, msg, field) => add(level, R, code, msg, field);

      if (!i.hazard) iss("error", "HAZ_EMPTY", "유해·위험요인이 비어 있습니다.", "hazard");
      else if (i.hazard.replace(/\s/g, "").length < 6) iss("info", "HAZ_SHORT", `유해·위험요인 서술이 너무 짧습니다(“${i.hazard}”). ‘무엇 때문에 → 어떤 재해’가 발생하는지 구체적으로 적어야 합니다.`, "hazard");

      if (has("freq") && i.freq == null) iss("error", "FREQ_EMPTY", "가능성(빈도) 값이 비어 있거나 숫자가 아닙니다.", "freq");
      if (has("sev") && i.sev == null) iss("error", "SEV_EMPTY", "중대성(강도) 값이 비어 있거나 숫자가 아닙니다.", "sev");
      if (i.freq != null && (i.freq < 1 || i.freq > fScale || !Number.isInteger(i.freq))) iss("error", "FREQ_RANGE", `가능성 값 ${i.freq}이(가) 척도(1~${fScale}) 범위를 벗어났습니다.`, "freq");
      if (i.sev != null && (i.sev < 1 || i.sev > sScale || !Number.isInteger(i.sev))) iss("error", "SEV_RANGE", `중대성 값 ${i.sev}이(가) 척도(1~${sScale}) 범위를 벗어났습니다.`, "sev");

      const expected = calc(i.freq, i.sev);
      if (has("risk")) {
        if (i.risk == null && i.riskGrade == null) {
          if (expected != null) iss("warn", "RISK_EMPTY", `위험성(위험도) 결과가 비어 있습니다. (가능성×중대성 = ${expected}) 수식이라면 엑셀에서 한 번 저장한 뒤 다시 올려주세요.`, "risk");
          else iss("error", "RISK_EMPTY", "위험성(위험도) 결과가 비어 있습니다.", "risk");
        }
        else if (expected != null && i.risk != null && Math.abs(expected - i.risk) > 1e-9)
          iss("error", "RISK_CALC", `위험성 계산 오류: 가능성 ${i.freq} ${method === "add" ? "+" : "×"} 중대성 ${i.sev} = ${expected}인데 ${i.risk}(으)로 기재되어 있습니다.`, "risk");
      }

      const curNum = i.risk != null ? i.risk : expected;
      const curLevel = levelOf(curNum, i.riskGrade);
      const needsMeasure = curLevel === "high";

      if (needsMeasure && !i.measure)
        iss("error", "MEAS_MISSING", `위험성(${curNum ?? text(i.riskRaw)})이 허용 기준(${acceptable}) 초과인데 감소대책이 없습니다.`, "measure");

      if (i.measure) {
        const nm = norm(i.measure);
        const weak = WEAK_MEASURE.test(nm);
        const strong = STRONG_MEASURE.test(nm);
        const ppe = PPE_ONLY.test(nm);
        if (weak && !strong && !ppe) iss("warn", "MEAS_WEAK", "감소대책이 ‘주의·교육·준수’ 등 형식적 관리대책에 그칩니다. 제거·대체·공학적 대책을 우선 검토하세요.", "measure");
        else if (ppe && !strong && needsMeasure) iss("warn", "MEAS_PPE", "고위험 항목의 감소대책이 개인보호구 착용 위주입니다. 보호구는 최후 수단이므로 공학적 대책을 함께 마련하세요.", "measure");
        if (nm.length < 4) iss("info", "MEAS_SHORT", `감소대책 서술이 너무 짧습니다(“${i.measure}”). 구체적인 조치 내용을 적어주세요.`, "measure");
        if (i.current && norm(i.current) === nm) iss("warn", "MEAS_SAME", "감소대책이 ‘현재 안전조치’와 동일합니다. 추가 대책 없이 위험성이 줄어든다고 볼 수 없습니다.", "measure");
      }

      // 개선 후
      const afterExpected = calc(i.afterFreq, i.afterSev);
      const afterNum = i.afterRisk != null ? i.afterRisk : afterExpected;
      const hasAfterCols = has("after_risk") || has("after_freq");
      if (hasAfterCols && i.measure && afterNum == null && i.afterGrade == null)
        iss("warn", "AFTER_EMPTY", "감소대책이 있으나 개선 후 위험성이 기재되지 않았습니다.", "after_risk");
      if (i.afterFreq != null && (i.afterFreq < 1 || i.afterFreq > fScale)) iss("error", "AFREQ_RANGE", `개선 후 가능성 ${i.afterFreq}이(가) 척도 범위를 벗어났습니다.`, "after_freq");
      if (i.afterSev != null && (i.afterSev < 1 || i.afterSev > sScale)) iss("error", "ASEV_RANGE", `개선 후 중대성 ${i.afterSev}이(가) 척도 범위를 벗어났습니다.`, "after_sev");
      if (afterExpected != null && i.afterRisk != null && Math.abs(afterExpected - i.afterRisk) > 1e-9)
        iss("error", "AFTER_CALC", `개선 후 위험성 계산 오류: ${i.afterFreq} ${method === "add" ? "+" : "×"} ${i.afterSev} = ${afterExpected}인데 ${i.afterRisk}(으)로 기재되어 있습니다.`, "after_risk");
      if (afterNum != null && curNum != null) {
        if (afterNum > curNum) iss("error", "AFTER_UP", `개선 후 위험성(${afterNum})이 개선 전(${curNum})보다 높습니다.`, "after_risk");
        else if (afterNum === curNum && i.measure) iss("warn", "AFTER_SAME", `감소대책 후에도 위험성(${afterNum})이 줄지 않았습니다. 대책의 실효성을 재검토하세요.`, "after_risk");
        if (afterNum > acceptable) iss("warn", "AFTER_HIGH", `개선 후 위험성(${afterNum})이 여전히 허용 기준(${acceptable})을 초과합니다. 추가 대책이 필요합니다.`, "after_risk");
        if (i.afterSev != null && i.sev != null && i.afterSev < i.sev && i.measure && !STRONG_MEASURE.test(norm(i.measure)))
          iss("info", "AFTER_SEV_DROP", `관리적 대책만으로 중대성이 ${i.sev}→${i.afterSev}로 낮아졌습니다. 중대성은 보통 공학적 대책(방호·격리 등)으로만 낮아집니다.`, "after_sev");
        if (curNum - afterNum >= Math.max(4, maxRisk * 0.5) && i.measure && !STRONG_MEASURE.test(norm(i.measure)))
          iss("warn", "AFTER_TOO_LOW", `대책 내용에 비해 위험성 감소폭(${curNum}→${afterNum})이 과도합니다. 개선 후 위험성이 과소평가되었을 수 있습니다.`, "after_risk");
      }

      // 이행관리
      if (i.measure && needsMeasure) {
        if (has("owner") && !i.owner) iss("warn", "OWNER_EMPTY", "감소대책 담당자가 지정되지 않았습니다.", "owner");
        if ((has("due") || has("done")) && !i.due && !i.done && !text(i.dueRaw)) iss("warn", "DUE_EMPTY", "개선 예정일(완료일)이 없습니다.", "due");
      }
      const doneLike = i.done || /(완료|조치완료|이행|done|o$|^o)/i.test(i.status);
      if (i.due && !doneLike && i.due < today) iss("warn", "OVERDUE", `개선 예정일(${fmtDate(i.due)})이 지났으나 완료 기록이 없습니다.`, "due");
      if (i.due && i.done && i.done > i.due) iss("info", "LATE_DONE", `예정일(${fmtDate(i.due)})보다 늦게 완료(${fmtDate(i.done)})되었습니다.`, "done");

      // 중복
      if (i.hazard) {
        const k = norm(i.process) + "|" + norm(i.hazard);
        if (dupKey.has(k)) iss("info", "DUP", `${dupKey.get(k)}행과 공정·유해위험요인이 중복됩니다.`, "hazard");
        else dupKey.set(k, R);
      }
    }

    // 분포 점검 — 모든 항목이 저위험이면 과소평가 의심
    const scored = items.map((i) => (i.risk != null ? i.risk : calc(i.freq, i.sev))).filter((x) => x != null);
    if (scored.length >= 5) {
      const highCnt = scored.filter((x) => x > acceptable).length;
      if (highCnt === 0) add("warn", null, "DIST_ALL_LOW", `전체 ${scored.length}개 항목이 모두 허용 기준(${acceptable}) 이하입니다. 위험성이 과소평가되지 않았는지 확인하세요.`);
      const uniq = new Set(scored);
      if (uniq.size === 1) add("warn", null, "DIST_SAME", `모든 항목의 위험성이 ${scored[0]}(으)로 동일합니다. 항목별로 실제 평가가 이루어졌는지 확인하세요.`);
    }
    const sevVals = items.map((i) => i.sev).filter((x) => x != null);
    const lethal = /(추락|협착|끼임|감전|폭발|화재|질식|붕괴|매몰|중독|산소결핍|깔림)/;
    for (const i of items) {
      if (i.sev != null && sScale >= 3 && i.sev === 1 && lethal.test(norm(i.hazard)))
        add("warn", i.rowNum, "SEV_UNDER", `중대재해 가능성이 있는 위험요인(추락·끼임·감전·화재 등)의 중대성이 최저값(1)입니다. 과소평가 여부를 확인하세요.`, "sev");
    }
    void sevVals;

    return finalize();

    function finalize() {
      findings.sort((a, b) => (a.row ?? 0) - (b.row ?? 0));
      const counts = { error: 0, warn: 0, info: 0 };
      for (const f of findings) counts[f.level]++;
      const rowsWithErr = new Set(findings.filter((f) => f.row && f.level === "error").map((f) => f.row));
      const rowsWithWarn = new Set(findings.filter((f) => f.row && f.level === "warn").map((f) => f.row));
      const n = items.length || 1;
      // 점수: 100점에서 감점 (행 단위 오류 비율 + 구조 문제)
      const structErr = findings.filter((f) => !f.row && f.level === "error").length;
      const structWarn = findings.filter((f) => !f.row && f.level === "warn").length;
      let score = 100 - (rowsWithErr.size / n) * 55 - (rowsWithWarn.size / n) * 25 - structErr * 12 - structWarn * 4;
      score = Math.max(0, Math.min(100, Math.round(score)));
      const verdict = score >= 85 && counts.error === 0 ? "적정" : score >= 60 ? "보완 필요" : "재평가 필요";
      return {
        ok: true,
        headerRow: hr + 1,
        dataStartRow: dataStart + 1,
        columns: Object.fromEntries(Object.entries(cols).map(([k, c]) => [k, { col: c, label: labels[c] ? labels[c].full : "" }])),
        settings: { method, fScale, sScale, maxRisk, acceptable, defaultAcceptable },
        meta,
        items: items.map((i) => ({ ...i, due: fmtDate(i.due), done: fmtDate(i.done) })),
        findings,
        counts,
        score,
        verdict,
      };
    }
  }

  const api = { analyze, findHeaderRow, _internal: { toDate, gradeOf, norm } };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.RiskAnalyzer = api;
})(typeof window !== "undefined" ? window : globalThis);
