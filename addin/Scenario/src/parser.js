/**
 * parser.js
 * Excelシートからシナリオデータを読み取るモジュール
 */

// ─── 列定義 ───────────────────────────────────────────
const SHEET_DEFS = {
  "異常（通常）": {
    dataStart: 13,   // 0-indexed row (14行目)
    colMode2: 0, colAutoNo: 1, colBrand: 3,
    colOp1Func: 4, colOp1Stat: 9,
    colOp2Func: 15, colOp2Stat: 20,
    colOp3Func: 26, colOp3Stat: 31,
    colCh: 85,
    colPhase: 96, colBlock: 87, colMinor: 88,
    colStart: 90, colEnd: 91, colToday: 95, // CR列（本日列）修正：94→95
  },
  "異常（電源断）": {
    dataStart: 13,
    colMode2: 0, colAutoNo: 1, colBrand: 3,
    colOp1Func: 4, colOp1Stat: 9,
    colOp2Func: 15, colOp2Stat: 20,
    colOp3Func: 26, colOp3Stat: 31,
    colCh: -1,
    colPhase: 70, colBlock: 61, colMinor: 62,
    colStart: 64, colEnd: 65, colToday: 69, // BR列（本日列）
  },
  "異常（通信断）": {
    dataStart: 13,
    colMode2: 0, colAutoNo: 1, colBrand: 3,
    colOp1Func: 4, colOp1Stat: 9,
    colOp2Func: 15, colOp2Stat: 20,
    colOp3Func: 26, colOp3Stat: 31,
    colCh: -1,
    colPhase: 89, colBlock: 80, colMinor: 81,
    colStart: 83, colEnd: 84, colToday: 88, // CK列（本日列）
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
  // "バ9,課16（済）,バ17" → [{id:"バ9",resolved:false},{id:"課16",resolved:true},...]
  const matches = text.match(/[バ課]\d+(?:（済）)?/g) || [];
  return matches.map(m => ({
    id: m.replace("（済）", ""),
    resolved: m.includes("（済）"),
  }));
}

function getLane(start, end, blockText, minorText) {
  const hasUnresolved = (text) => {
    if (!text) return false;
    const bugs = text.match(/[バ課]\d+(?:（済）)?/g) || [];
    return bugs.some(b => !b.includes("（済）"));
  };
  const unresolvedBlock = hasUnresolved(blockText);
  const unresolvedMinor = hasUnresolved(minorText);
  const unresolved = unresolvedBlock || unresolvedMinor;
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
    usedRange.load(["values", "rowIndex"]);
    await context.sync();

    const rows = usedRange.values;
    const usedRangeStartRow = usedRange.rowIndex; // usedRangeの開始行番号
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
      const todayValue = colCount > def.colToday ? cleanVal(row[def.colToday]) : "";
      const isStar = todayValue === "〇"; // 本日列が〇の場合は★
      const lane  = getLane(start, end, blockText, minorText);

      const key = `${sheetLabel}|${autoNo}|${brand}|${op1Func}`;
      if (seen.has(key)) continue;
      seen.add(key);

      creationData.push({
        sheet: sheetLabel, no: autoNo, brand,
        op1Func, op1Stat, op2Func, op2Stat, op3Func, op3Stat,
        phase, lane, blockText, minorText,
        isStar, // 本日列の★/☆状態
        excelSheet: sheetName,  // Excelへの書き戻しに使用
        rowIdx: usedRangeStartRow + i,  // 実際のExcel行番号（1-indexed）
        colBlock: def.colBlock,
        colMinor: def.colMinor,
        colToday: def.colToday, // 本日列のインデックス
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
            rowIdx: usedRangeStartRow + i,  // 実際のExcel行番号（1-indexed）
            colBlock: def.colBlock,
            colMinor: def.colMinor,
          });
        }
      }
    }
  }

  // ─── 正常シート ─────────────────────────────────────
  const normalSheetName = "正常";
  if (sheetNames.includes(normalSheetName)) {
    const ws = sheets.getItem(normalSheetName);
    const usedRange = ws.getUsedRange();
    usedRange.load(["values", "rowIndex"]);
    await context.sync();

    const rows = usedRange.values;
    const usedRangeStartRow = usedRange.rowIndex; // usedRangeの開始行番号
    const colCount = rows[0] ? rows[0].length : 0;
    
    // 正常シートも異常系シートと同様の構造と仮定
    const normalDef = {
      dataStart: 13, colMode2: 0, colAutoNo: 1, colBrand: 3,
      colOp1Func: 4, colOp1Stat: 9, colOp2Func: 15, colOp2Stat: 20,
      colOp3Func: 26, colOp3Stat: 31, colPhase: 96,
      colStart: 90, colEnd: 91, colToday: 95 // CR列（本日列）修正：94→95
    };

    for (let i = normalDef.dataStart; i < rows.length; i++) {
      const row = rows[i];
      if (cleanVal(row[normalDef.colMode2]) !== "●") continue;
      const autoNoRaw = row[normalDef.colAutoNo];
      if (!autoNoRaw && autoNoRaw !== 0) continue;
      const autoNo = parseInt(autoNoRaw);
      if (isNaN(autoNo)) continue;

      const brand = cleanBrand(cleanVal(row[normalDef.colBrand]));
      const op1Func = cleanVal(row[normalDef.colOp1Func]);
      const op1Stat = cleanVal(row[normalDef.colOp1Stat]);
      const op2Func = cleanVal(row[normalDef.colOp2Func]);
      const op2Stat = cleanVal(row[normalDef.colOp2Stat]);
      const op3Func = colCount > normalDef.colOp3Func ? cleanVal(row[normalDef.colOp3Func]) : "";
      const op3Stat = colCount > normalDef.colOp3Stat ? cleanVal(row[normalDef.colOp3Stat]) : "";
      let phase = colCount > normalDef.colPhase ? cleanVal(row[normalDef.colPhase]) : "";
      if (!phase) phase = "PH1";
      const start = colCount > normalDef.colStart ? dateStr(row[normalDef.colStart]) : "";
      const end = colCount > normalDef.colEnd ? dateStr(row[normalDef.colEnd]) : "";
      const todayValue = colCount > normalDef.colToday ? cleanVal(row[normalDef.colToday]) : "";
      const isStar = todayValue === "〇";
      const lane = getLane(start, end, "", "");

      const key = `正常|${autoNo}|${brand}|${op1Func}`;
      if (seen.has(key)) continue;
      seen.add(key);

      creationData.push({
        sheet: "正常", no: autoNo, brand,
        op1Func, op1Stat, op2Func, op2Stat, op3Func, op3Stat,
        phase, lane, blockText: "", minorText: "",
        isStar,
        excelSheet: normalSheetName, rowIdx: usedRangeStartRow + i, colBlock: -1, colMinor: -1,
        colToday: normalDef.colToday, // CR列（本日列）
      });
    }
  }

  // ─── 正常（クレ・銀聯）────────────────────────────────
  const kuSheetName = "正常（クレ・銀聯）";
  if (sheetNames.includes(kuSheetName)) {
    const ws = sheets.getItem(kuSheetName);
    const usedRange = ws.getUsedRange();
    usedRange.load(["values", "rowIndex"]);
    await context.sync();

    const rows = usedRange.values;
    const usedRangeStartRow = usedRange.rowIndex; // usedRangeの開始行番号
    let rowNum = 1;
    for (let i = 4; i < rows.length; i++) {
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
      const todayValue = cleanVal(row[19]); // T列（本日列）
      const isStar = todayValue === "〇"; // 本日列が〇の場合は★
      const lane  = getLane(start, end, "", "");

      const op1Func = func + (gyoumu ? `（${gyoumu}）` : "");
      const op1Stat = [haraiKu, signPin].filter(Boolean).join("　");

      const key = `正常（クレ・銀聯）|${scenarioId}|${brand}|${op1Func}`;
      if (seen.has(key)) continue;
      seen.add(key);

      const finalRowIdx = usedRangeStartRow + i;
      
      // カード#1、#2、#3のデバッグ出力
      if (rowNum <= 3) {
        console.log(`正常（クレ・銀聯）カード#${rowNum}のデバッグ情報:`);
        console.log(`  usedRangeStartRow: ${usedRangeStartRow}`);
        console.log(`  配列インデックスi: ${i} (Excel行番号: ${i + 1})`);
        console.log(`  計算されたrowIdx: ${finalRowIdx}`);
        console.log(`  scenarioId: ${scenarioId}, brand: ${brand}`);
        console.log(`  シート: ${kuSheetName}`);
      }

      creationData.push({
        sheet: "正常（クレ・銀聯）", no: rowNum, brand,
        op1Func, op1Stat, op2Func: "", op2Stat: "", op3Func: "", op3Stat: "",
        phase, lane, blockText: "", minorText: "",
        isStar, // T列から★/☆状態を取得
        excelSheet: kuSheetName, rowIdx: finalRowIdx, colBlock: -1, colMinor: -1,
        colToday: 19, // T列（本日列）
      });
      rowNum++;
    }
  }

  // bugMapをbugDataに変換
  const bugData = Object.values(bugMap)
    .sort((a, b) => {
      const getPrefix = (id) => id.charAt(0); // 「バ」または「課」
      const getNumber = (id) => parseInt(id.slice(1));
      
      const prefixA = getPrefix(a.id);
      const prefixB = getPrefix(b.id);
      
      // 同じプリフィックスなら番号で比較、異なるならプリフィックスで比較
      if (prefixA === prefixB) {
        return getNumber(a.id) - getNumber(b.id);
      }
      return prefixA.localeCompare(prefixB);
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

        let newVal = cellVal;
        if (resolved) {
          // バ9 → バ9（済）  ※すでに（済）がついている場合はそのまま
          // カンマ・読点・スペース区切りで分割して個別に処理
          const parts = cellVal.split(/[,、\s]+/).map(part => {
            const trimmed = part.trim();
            if (trimmed === bugId && !trimmed.includes('（済）')) {
              return bugId + '（済）';
            }
            return trimmed;
          }).filter(part => part); // 空文字を除去
          
          // 元の区切り文字を保持（全角読点があれば全角、なければ半角カンマ）
          const separator = cellVal.includes('、') ? '、' : ',';
          newVal = parts.join(separator);
        } else {
          // バ9（済） → バ9  
          // カンマ・読点・スペース区切りで分割して個別に処理
          const parts = cellVal.split(/[,、\s]+/).map(part => {
            const trimmed = part.trim();
            if (trimmed === bugId + '（済）') {
              return bugId;
            }
            return trimmed;
          }).filter(part => part); // 空文字を除去
          
          // 元の区切り文字を保持（全角読点があれば全角、なければ半角カンマ）
          const separator = cellVal.includes('、') ? '、' : ',';
          newVal = parts.join(separator);
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
