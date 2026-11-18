// 4M 위험성평가표 양식(2단 헤더, 상단 결재/평가정보)을 흉내 낸 테스트 파일 생성
const ExcelJS = require("exceljs");
(async () => {
  const wb = new ExcelJS.Workbook();
  const info = wb.addWorksheet("위험정보");
  info.addRow(["코드", "위험요인 분류"]); info.addRow(["A01", "끼임"]); info.addRow(["A02", "화상"]);
  const ws = wb.addWorksheet("PM21 A반");
  const B = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
  ws.getCell("A1").value = "평가대상 공정명"; ws.getCell("B1").value = "Dryer Part";
  ws.getCell("A2").value = "세부공정"; ws.getCell("B2").value = "Sheet 건조공정 (PM21-A)";
  ws.getCell("A3").value = "평가일시"; ws.getCell("B3").value = "정기■\n수시□"; ws.getCell("C3").value = "2025.04.25";
  ws.getCell("E1").value = "위험성평가표\n(4M-Risk Assessment)"; ws.mergeCells("E1:H3");
  ws.getCell("I1").value = "평가자\n(리더 및 팀원)"; ws.getCell("K1").value = "반장 : 하OO 반장"; ws.getCell("K2").value = "반원 : 황OO, 이OO, 김OO";
  ws.getCell("I3").value = "평균위험도"; ws.getCell("K3").value = "현재"; ws.getCell("L3").value = "개선후";
  ws.getCell("K4").value = 3.47; ws.getCell("L4").value = 3.46;
  const r1 = 5;
  ws.getRow(r1).values = ["작업내용", "평가\n구분", "위험요인 및 재해형태", "재해형태", "현재안전조치", "개선전 위험도", "", "", "개선대책", "코드NO", "개선후 위험도", "", ""];
  ws.getRow(r1 + 1).values = ["", "", "", "", "", "빈도", "강도", "위험도", "", "", "빈도", "강도", "위험도"];
  ["A", "B", "C", "D", "E", "I", "J"].forEach((c) => ws.mergeCells(`${c}${r1}:${c}${r1 + 1}`));
  ws.mergeCells(`F${r1}:H${r1}`); ws.mergeCells(`K${r1}:M${r1}`);
  const D = [
    ["지필연결작업", "설비적", "상부에서 떨어지는 응축수로 인해 연결중 화상을 입는다.", "화상", "위험예지", 1, 2, 2, "", "", 1, 2, 2],
    ["", "", "C-rope와 실린더 사이에 끼여 다친다.", "끼임", "안전대 복원", 2, 2, 4, "", "", 2, 2, 4],
    ["", "", "족장또는 워크레일이 불균일 하여 이동시 발이걸려 넘어져 다친다.", "넘어짐", "침착한 행동", 3, 1, 3, "", "", 3, 1, 3],
    ["", "환경적", "고온 Steam에 화상을 입는다.", "화상", "보온커버 설치 / 토시 착용", 2, 2, 4, "", "", 2, 2, 4],
    ["", "관리적", "지필 연결방법 지도 부족으로 화상을 입는다.", "화상", "반복교육", 1, 3, 3, "", "", 1, 3, 3],
    ["Dryer Part 구동 작업", "설비적", "내부에 진입 내부 확인시 구동되어 다친다", "기타", "내부 진입시 작업표찰 부착\n전원 차단/진입허가서 작성", 1, 4, 4, "", "", 1, 4, 4],
    ["", "", "구동후 지분 제거시 지분이 비산되어 눈을 다친다", "비래", "일반 보안경 착용", 2, 2, 4, "", "", 2, 2, 4],
    ["Doctor blade 교체", "설비적", "#34번 DOCTOR BLADE 교체시 배관에 발이걸려 넘어진다.", "넘어짐", "현장확인 및 위험예지", 4, 1, 4, "", "", 4, 1, 4],
    ["", "인적", "회전상태에서 교체하다 손이 협착되어 다친다.", "끼임", "M/C 중지후 작업", 2, 2, 4, "", "", 2, 2, 4],
    ["", "", "면장갑을 끼고 칼날을 잡고 작업하다 손을 베어 다친다.", "베임", "소도구 이용", 2, 3, 6, "칼날 전용 홀더 사용, 베임방지 장갑 지급", "", 1, 3, 4],
  ];
  D.forEach((d, i) => { const row = ws.getRow(r1 + 2 + i); row.values = d; row.height = 18; });
  ws.mergeCells(`A${r1 + 2}:A${r1 + 6}`); ws.mergeCells(`B${r1 + 2}:B${r1 + 4}`);
  for (let r = r1; r < r1 + 2 + D.length; r++) for (let c = 1; c <= 13; c++) {
    const cell = ws.getCell(r, c); cell.border = B; cell.font = { size: 9 };
    if (r < r1 + 2) { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } }; cell.font = { bold: true, size: 9 }; }
  }
  [12, 6, 48, 9, 24, 5, 5, 6, 24, 7, 5, 5, 6].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  // 실제 양식처럼: 개선후 위험도 옆 메모 열(연도), 위험도 빨간 글씨 조건부 서식, 로고 그림, 인쇄 영역
  ws.getCell(`N${r1 + 2}`).value = "23년"; ws.getCell(`N${r1 + 4}`).value = "21년"; ws.getCell(`N${r1 + 11}`).value = 22;
  ws.getColumn(14).width = 4;
  ws.addConditionalFormatting({ ref: `H${r1 + 2}:H${r1 + 11}`, rules: [{ type: "cellIs", operator: "greaterThan", formulae: [2], priority: 1, style: { font: { color: { argb: "FFFF0000" } } } }] });
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAEAAAAAQCAIAAAD4YuoOAAAAGklEQVR4nGNgGAWjYBSMglEwCkbBKBgFgwYAAOAAAZ0ZNkkAAAAASUVORK5CYII=", "base64");
  const img = wb.addImage({ buffer: png, extension: "png" });
  ws.addImage(img, { tl: { col: 2.2, row: 0.2 }, ext: { width: 90, height: 24 } });
  ws.pageSetup.printArea = `A1:M${r1 + 11}`;
  await wb.xlsx.writeFile(process.argv[2]);
})();
