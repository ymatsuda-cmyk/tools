/**
 * parser.js
 * Excelシートからシナリオデータを読み取るモジュール
 */

// ─── 列定義 ───────────────────────────────────────────
const SHEET_DEFS = {
  "異常（通常）": {
    dataStart: 14,   // 0-indexed row
    colMode2: 0, colAutoNo: 1, colBrand: 3,
    colOp1Func: 4, colOp1Stat: 9,
    colOp2Func: 15, colOp2Stat: 20,
    colOp3Func: 26, colOp3Stat: 31,
    colCh: 85,
    colPhase: 96, colBlock: 87, colMinor: 88,
    colStart: 90, colEnd: 91,
  },
  "異常（電源断）": {
    dataStart: 14,
    colMode2: 0, colAutoNo: 1, colBrand: 3,
    colOp1Func: 4, colOp1Stat: 9,
    colOp2Func: 15, colOp2Stat: 20,
    colOp3Func: 26, colOp3Stat: 31,
    colCh: -1,
    colPhase: 70, colBlock: 61, colMinor: 62,
    colStart: 64, colEnd: 65,
  },
  "異常（通信断）": {
    dataStart: 14,
    colMode2: 0, colAutoNo: 1, colBrand: 3,
    colOp1Func: 4, colOp1Stat: 9,
    colOp2Func: 15, colOp2Stat: 20,
    colOp3Func: 26, colOp3Stat: 31,
    colCh: -1,
    colPhase: 89, colBlock: 80, colMinor: 81,
    colStart: 83, colEnd: 84,
  },
};

// ─── ユーティリティ ───────────────────────────────────
function cleanVal(v) {
  if (v === null || v === undefined || v === "") return "";
  return String(v).replace(/\n/g, "").trim();
}

function cleanBrand(b) {
  return b.replace(/^\d+:/, "").trim();
}

function dateStr(v) {
  if (!v) return "";
  const s = String(v).trim();
  // "2026-03-30 00:00:00" → "2026-03-30"
  return s.length >= 10 ? s.substring(0, 10) : s;
}

function extractBugIds(text) {
  // "バ9,バ16（済）,バ17" → [{id:"バ9",resolved:false},{id:"バ16",resolved:true},...]
  const matches = text.match(/バ\d+(?:（済）)?/g) || [];
  return matches.map(m => ({
    id: m.replace("（済）", ""),
    resolved: m.includes("（済）"),
  }));
}

function getLane(start, end, blockText) {
  const hasUnresolved = (text) => {
    if (!text) return false;
    const bugs = text.match(/バ\d+(?:（済）)?/g) || [];
    return bugs.some(b => !b.includes("（済）"));
  };
  const unresolved = hasUnresolved(blockText);
  if (!start) return unresolved ? "バグ保留" : "未着手";
  if (!end)   return unresolved ? "バグ保留" : "対応中";
  return unresolved ? "完了（条件付き）" : "完了";
}

// ─── メイン読み取り関数 ───────────────────────────────
/**
 * Excelワークブックからシナリオデータとバグデータを読み取る
 * @param {Excel.Workbook} workbook - Office JS workbook context
 * @returns {Promise<{creationData: Array, bugData: Array}>}
 */
async function readWorkbook(context) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  const sheetNames = sheets.items.map(s => s.name);
  const creationData = [];
  const bugMap = {};  // bugId -> {id, scenarios[]}
  const seen = new Set();

  // ─── 電マネ系3シート ─────────────────────────────────
  for (const [sheetName, def] of Object.entries(SHEET_DEFS)) {
    if (!sheetNames.includes(sheetName)) continue;

    const ws = sheets.getItem(sheetName);
    const usedRange = ws.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const rows = usedRange.values;
    const colCount = rows[0] ? rows[0].length : 0;

    for (let i = def.dataStart; i < rows.length; i++) {
      const row = rows[i];

      // A列が●のみ対象
      if (cleanVal(row[def.colMode2]) !== "●") continue;

      const autoNoRaw = row[def.colAutoNo];
      if (!autoNoRaw && autoNoRaw !== 0) continue;
      const autoNo = parseInt(autoNoRaw);
      if (isNaN(autoNo)) continue;

      // シートラベル（異常（通常）はCH列で分岐）
      let sheetLabel = sheetName;
      if (sheetName === "異常（通常）" && def.colCh >= 0 && colCount > def.colCh) {
        const chVal = cleanVal(row[def.colCh]);
        sheetLabel = chVal.includes("正常") ? "正常" : "異常（通常）";
      }

      const brand   = cleanBrand(cleanVal(row[def.colBrand]));
      const op1Func = cleanVal(row[def.colOp1Func]);
      const op1Stat = cleanVal(row[def.colOp1Stat]);
      const op2Func = cleanVal(row[def.colOp2Func]);
      const op2Stat = cleanVal(row[def.colOp2Stat]);
      const op3Func = colCount > def.colOp3Func ? cleanVal(row[def.colOp3Func]) : "";
      const op3Stat = colCount > def.colOp3Stat ? cleanVal(row[def.colOp3Stat]) : "";

      let phase = colCount > def.colPhase ? cleanVal(row[def.colPhase]) : "";
      if (!phase) phase = "PH1";

      const blockText = colCount > def.colBlock ? cleanVal(row[def.colBlock]) : "";
      const minorText = colCount > def.colMinor ? cleanVal(row[def.colMinor]) : "";
      const start = colCount > def.colStart ? dateStr(row[def.colStart]) : "";
      const end   = colCount > def.colEnd   ? dateStr(row[def.colEnd])   : "";
      const lane  = getLane(start, end, blockText);

      const key = `${sheetLabel}|${autoNo}|${brand}|${op1Func}`;
      if (seen.has(key)) continue;
      seen.add(key);

      creationData.push({
        sheet: sheetLabel, no: autoNo, brand,
        op1Func, op1Stat, op2Func, op2Stat, op3Func, op3Stat,
        phase, lane, blockText, minorText,
        excelSheet: sheetName,  // Excelへの書き戻しに使用
        rowIdx: i,              // 0-indexed row in usedRange
        colBlock: def.colBlock,
        colMinor: def.colMinor,
      });

      // バグ影響データ収集
      for (const [colType, text] of [["block", blockText], ["minor", minorText]]) {
        const bugs = extractBugIds(text);
        for (const { id, resolved } of bugs) {
          if (!bugMap[id]) bugMap[id] = { id, scenarios: [] };
          bugMap[id].scenarios.push({
            sheet: sheetLabel, no: autoNo, brand,
            op1Func, op1Stat, op2Func, op2Stat, op3Func, op3Stat,
            isBlock: colType === "block",
            resolved,
            excelSheet: sheetName,
            rowIdx: i,
            colBlock: def.colBlock,
            colMinor: def.colMinor,
          });
        }
      }
    }
  }

  // ─── 正常（クレ・銀聯）────────────────────────────────
  const kuSheetName = "正常（クレ・銀聯）";
  if (sheetNames.includes(kuSheetName)) {
    const ws = sheets.getItem(kuSheetName);
    const usedRange = ws.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const rows = usedRange.values;
    let rowNum = 1;
    for (let i = 5; i < rows.length; i++) {
      const row = rows[i];
      const scenarioId = cleanVal(row[2]);
      if (!scenarioId) continue;
      const brand = cleanVal(row[10]);
      if (!brand) continue;

      const gyoumu   = cleanVal(row[3]);
      const func     = cleanVal(row[4]);
      const haraiKu  = cleanVal(row[5]);
      const signPin  = cleanVal(row[6]);
      let phase = cleanVal(row[20]) || "PH1";
      const start = dateStr(row[14]);
      const end   = dateStr(row[15]);
      const lane  = getLane(start, end, "");

      const op1Func = func + (gyoumu ? `（${gyoumu}）` : "");
      const op1Stat = [haraiKu, signPin].filter(Boolean).join("　");

      const key = `正常（クレ・銀聯）|${scenarioId}|${brand}|${op1Func}`;
      if (seen.has(key)) continue;
      seen.add(key);

      creationData.push({
        sheet: "正常（クレ・銀聯）", no: rowNum, brand,
        op1Func, op1Stat, op2Func: "", op2Stat: "", op3Func: "", op3Stat: "",
        phase, lane, blockText: "", minorText: "",
        excelSheet: kuSheetName, rowIdx: i, colBlock: -1, colMinor: -1,
      });
      rowNum++;
    }
  }

  // bugMapをbugDataに変換
  const bugData = Object.values(bugMap)
    .sort((a, b) => {
      const na = parseInt(a.id.replace("バ", ""));
      const nb = parseInt(b.id.replace("バ", ""));
      return na - nb;
    })
    .map(b => ({
      id: b.id,
      resolved: b.scenarios.every(s => s.resolved),
      scenarios: b.scenarios,
    }));

  return { creationData, bugData };
}

/**
 * バグIDの解消確認をExcelに書き戻す
 * @param {Excel.RequestContext} context
 * @param {string} bugId - 例: "バ9"
 * @param {boolean} resolved - true=済み, false=済みを解除
 * @param {Array} scenarios - 対象シナリオ一覧
 */
async function writeBugResolution(context, bugId, resolved, scenarios) {
  const sheets = context.workbook.worksheets;

  // シートごとに処理をまとめる
  const bySheet = {};
  for (const s of scenarios) {
    if (!bySheet[s.excelSheet]) bySheet[s.excelSheet] = [];
    bySheet[s.excelSheet].push(s);
  }

  for (const [sheetName, scenList] of Object.entries(bySheet)) {
    const ws = sheets.getItem(sheetName);
    const usedRange = ws.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const values = usedRange.values;

    for (const scen of scenList) {
      const ri = scen.rowIdx;
      if (ri >= values.length) continue;

      // block列とminor列それぞれ更新
      for (const [colType, colIdx] of [["block", scen.colBlock], ["minor", scen.colMinor]]) {
        if (colIdx < 0 || colIdx >= values[ri].length) continue;
        const cellVal = String(values[ri][colIdx] || "").trim();
        if (!cellVal) continue;

        // そのバグIDが含まれていなければスキップ
        if (!cellVal.includes(bugId)) continue;

        let newVal;
        if (resolved) {
          // バ9 → バ9（済）  ※すでに（済）がついている場合はそのまま
          newVal = cellVal.replace(
            new RegExp(bugId + "(?!（済）)", "g"),
            bugId + "（済）"
          );
        } else {
          // バ9（済） → バ9
          newVal = cellVal.replace(new RegExp(bugId + "（済）", "g"), bugId);
        }

        if (newVal !== cellVal) {
          const cell = ws.getCell(ri, colIdx);
          cell.values = [[newVal]];
        }
      }
    }
  }

  await context.sync();
}
