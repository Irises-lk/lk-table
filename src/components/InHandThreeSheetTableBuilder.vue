<template>
  <section class="in-hand-three-sheet-page">
    <h3>在手合同三表填充输出</h3>

    <div v-if="!hideUploader" class="upload-row">
      <label for="threeSheetFile">上传 Excel 文件：</label>
      <input
        id="threeSheetFile"
        type="file"
        accept=".xlsx,.xls"
        @change="onFileChange"
      />
      <!-- <button type="button" @click="exportResultExcel" :disabled="!canExport">
        导出结果表格
      </button> -->
    </div>

    <div v-if="statusText" class="result-panel">{{ statusText }}</div>
    <div v-else-if="errorText" class="error-panel">{{ errorText }}</div>

    <div v-if="canExport" class="table-wrap">
      <table class="result-table">
        <thead></thead>
        <tbody>
          <tr>
            <td :colspan="7" style="text-align: center">新签数据</td>
          </tr>
          <tr>
            <td v-for="head in TABLE_HEADERS" :key="`new-${head}`">
              {{ head }}
            </td>
          </tr>
          <tr
            v-for="row in resultData.newSignRows"
            :key="`new-${row.serial}-${row.projectName}`"
          >
            <td>{{ row.serial }}</td>
            <td>{{ row.customerName }}</td>
            <td>{{ row.projectName }}</td>
            <td>{{ row.owner }}</td>
            <td>{{ row.coOwner }}</td>
            <td>{{ row.product }}</td>
            <td class="amount">{{ row.amount }}</td>
          </tr>
          <tr>
            <td :colspan="7" style="text-align: center">营收数据</td>
          </tr>
          <tr>
            <td v-for="head in TABLE_HEADERS" :key="`rev-${head}`">
              {{ head }}
            </td>
          </tr>
          <tr
            v-for="row in resultData.revenueRows"
            :key="`rev-${row.serial}-${row.projectName}`"
          >
            <td>{{ row.serial }}</td>
            <td>{{ row.customerName }}</td>
            <td>{{ row.projectName }}</td>
            <td>{{ row.owner }}</td>
            <td>{{ row.coOwner }}</td>
            <td>{{ row.product }}</td>
            <td class="amount">{{ row.amount }}</td>
          </tr>

          <tr>
            <td :colspan="7" style="text-align: center">回款数据</td>
          </tr>
          <tr>
            <td v-for="head in TABLE_HEADERS" :key="`rcp-${head}`">
              {{ head }}
            </td>
          </tr>
          <tr
            v-for="row in resultData.receiptRows"
            :key="`rcp-${row.serial}-${row.projectName}`"
          >
            <td>{{ row.serial }}</td>
            <td>{{ row.customerName }}</td>
            <td>{{ row.projectName }}</td>
            <td>{{ row.owner }}</td>
            <td>{{ row.coOwner }}</td>
            <td>{{ row.product }}</td>
            <td class="amount">{{ row.amount }}</td>
          </tr>
        </tbody>
      </table>

      <!-- <div class="actions">
        <button type="button" @click="exportResultExcel">导出结果表格</button>
      </div> -->
    </div>
  </section>
</template>

<script setup>
// 中文注释：本组件按固定模板输出“新签/营收/回款”三分区数据，并支持导出 Excel。
import { computed, ref, watch } from "vue";
import { useStore } from "vuex";
import * as XLSX from "xlsx";

const TABLE_HEADERS = [
  "序号",
  "客户名称",
  "项目名称",
  "对应责任方",
  "对应协办方",
  "产品名称和数量",
  "对应金额",
];

const SHEET_RULES = {
  newSign: {
    candidates: ["新签"],
    dateColumn: ["合同签订日期"],
    customerColumn: ["客户名称"],
    projectColumn: ["项目名称"],
    amountColumn: ["合同金额"],
    ownerColumn: ["承揽主办方"],
    coOwnerColumn: ["承揽协办方"],
    productColumn: ["产品名称和数量"],
  },
  revenue: {
    candidates: ["营收", "收入"],
    dateColumn: ["确认收入月日"],
    customerColumn: ["客户名称"],
    projectColumn: ["项目名称"],
    amountColumn: ["销售收入金额(不含税)", "销售收入金额（不含税）"],
    ownerColumn: ["执行主办方"],
    coOwnerColumn: ["执行协办方"],
    productColumn: ["产品名称和数量"],
  },
  receipt: {
    candidates: ["回款"],
    dateColumn: ["回款月日"],
    customerColumn: ["客户名称"],
    projectColumn: ["项目名称"],
    amountColumn: ["回款金额"],
    ownerColumn: ["执行主办方"],
    coOwnerColumn: ["执行协办方"],
    productColumn: ["产品名称和数量"],
  },
};

const props = defineProps({
  externalFile: { type: Object, default: null },
  generateKey: { type: Number, default: 0 },
  hideUploader: { type: Boolean, default: false },
});

const store = useStore();
const statusText = ref(
  store.state.inHandThreeSheetTableResult?.statusText || "",
);
const errorText = ref(store.state.inHandThreeSheetTableResult?.errorText || "");
const resultData = ref(
  store.state.inHandThreeSheetTableResult?.data || {
    targetMonth: "",
    newSignRows: [],
    revenueRows: [],
    receiptRows: [],
  },
);

const canExport = computed(() => {
  return (
    resultData.value.newSignRows.length > 0 ||
    resultData.value.revenueRows.length > 0 ||
    resultData.value.receiptRows.length > 0
  );
});

watch(
  [statusText, errorText, resultData],
  () => {
    store.commit("setInHandThreeSheetTableResult", {
      statusText: statusText.value,
      errorText: errorText.value,
      data: resultData.value,
    });
  },
  { deep: true },
);

watch(
  () => props.generateKey,
  () => {
    if (!props.externalFile) return;
    onFileChange({ target: { files: [props.externalFile] } });
  },
);

function onFileChange(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  statusText.value = "";
  errorText.value = "";
  resultData.value = {
    targetMonth: "",
    newSignRows: [],
    revenueRows: [],
    receiptRows: [],
  };

  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const data = new Uint8Array(ev.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetMap = recognizeSheets(workbook);

      if (!sheetMap.newSign || !sheetMap.revenue || !sheetMap.receipt) {
        const missing = [];
        if (!sheetMap.newSign) missing.push("新签");
        if (!sheetMap.revenue) missing.push("营收");
        if (!sheetMap.receipt) missing.push("回款");
        errorText.value = `未识别到以下子表：${missing.join("、")}`;
        return;
      }

      const prev = getPreviousMonth();
      const newSignRows = extractSectionRows(
        workbook,
        sheetMap.newSign,
        SHEET_RULES.newSign,
        prev.month,
      );
      const revenueRows = extractSectionRows(
        workbook,
        sheetMap.revenue,
        SHEET_RULES.revenue,
        prev.month,
      );
      const receiptRows = extractSectionRows(
        workbook,
        sheetMap.receipt,
        SHEET_RULES.receipt,
        prev.month,
      );

      resultData.value = {
        targetMonth: String(prev.month),
        newSignRows,
        revenueRows,
        receiptRows,
      };

      statusText.value = `处理完成：已按${prev.month}月筛选并生成表格（新签${newSignRows.length}行，营收${revenueRows.length}行，回款${receiptRows.length}行）。`;
    } catch (error) {
      errorText.value = `解析失败：${error && error.message ? error.message : String(error)}`;
    }
  };

  reader.readAsArrayBuffer(file);
}

function recognizeSheets(workbook) {
  const names = workbook.SheetNames || [];
  const map = {
    newSign: "",
    revenue: "",
    receipt: "",
    ledger: "",
  };

  for (const name of names) {
    const normalized = normalizeHeader(name);
    if (
      !map.newSign &&
      SHEET_RULES.newSign.candidates.some((k) =>
        normalized.includes(normalizeHeader(k)),
      )
    ) {
      map.newSign = name;
    }
    if (
      !map.revenue &&
      SHEET_RULES.revenue.candidates.some((k) =>
        normalized.includes(normalizeHeader(k)),
      )
    ) {
      map.revenue = name;
    }
    if (
      !map.receipt &&
      SHEET_RULES.receipt.candidates.some((k) =>
        normalized.includes(normalizeHeader(k)),
      )
    ) {
      map.receipt = name;
    }
    if (!map.ledger && normalized.includes(normalizeHeader("业绩台账"))) {
      map.ledger = name;
    }
  }

  return map;
}

function extractSectionRows(workbook, sheetName, rule, targetMonth) {
  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
  if (!rows.length) return [];

  const headerInfo = findHeaderInfo(rows, rule);
  if (!headerInfo) {
    throw new Error(`子表“${sheetName}”缺少必需列，请检查列标题完整性。`);
  }

  const result = [];
  for (let r = headerInfo.headerRowIndex + 1; r < rows.length; r++) {
    const monthValue = parseMonthFromCell(
      worksheet,
      rows,
      r,
      headerInfo.colIndex.dateCol,
    );
    if (monthValue !== targetMonth) continue;

    const customerName = normalizeCellText(
      getCellDisplayValue(worksheet, rows, r, headerInfo.colIndex.customerCol),
    );
    const projectName = normalizeCellText(
      getCellDisplayValue(worksheet, rows, r, headerInfo.colIndex.projectCol),
    );
    const owner = normalizeCellText(
      getCellDisplayValue(worksheet, rows, r, headerInfo.colIndex.ownerCol),
    );
    const coOwner = normalizeCellText(
      getCellDisplayValue(worksheet, rows, r, headerInfo.colIndex.coOwnerCol),
    );
    const product = normalizeCellText(
      getCellDisplayValue(worksheet, rows, r, headerInfo.colIndex.productCol),
    );

    const amountRaw = getCellDisplayValue(
      worksheet,
      rows,
      r,
      headerInfo.colIndex.amountCol,
    );
    const amountNumber = parseAmount(amountRaw);
    const amountWan = roundHalfUp(amountNumber / 10000);

    result.push({
      serial: result.length + 1,
      customerName,
      projectName,
      owner,
      coOwner,
      product,
      amount: String(amountWan),
    });
  }

  return result;
}

function findHeaderInfo(rows, rule) {
  for (let i = 0; i < rows.length; i++) {
    const headers = (rows[i] || []).map((cell) => normalizeHeader(cell));

    const dateCol = findColumnIndex(headers, rule.dateColumn);
    const customerCol = findColumnIndex(headers, rule.customerColumn);
    const projectCol = findColumnIndex(headers, rule.projectColumn);
    const amountCol = findColumnIndex(headers, rule.amountColumn);
    const ownerCol = findColumnIndex(headers, rule.ownerColumn);
    const coOwnerCol = findColumnIndex(headers, rule.coOwnerColumn);
    const productCol = findColumnIndex(headers, rule.productColumn);

    if (
      [
        dateCol,
        customerCol,
        projectCol,
        amountCol,
        ownerCol,
        coOwnerCol,
        productCol,
      ].every((idx) => idx >= 0)
    ) {
      return {
        headerRowIndex: i,
        colIndex: {
          dateCol,
          customerCol,
          projectCol,
          amountCol,
          ownerCol,
          coOwnerCol,
          productCol,
        },
      };
    }
  }

  return null;
}

function findColumnIndex(headers, candidates) {
  for (const candidate of candidates) {
    const idx = headers.indexOf(normalizeHeader(candidate));
    if (idx >= 0) return idx;
  }
  return -1;
}

function getCellDisplayValue(worksheet, rows, rowIndex, colIndex) {
  if (colIndex < 0) return "";

  // 中文注释：优先使用 format_cell，保证百分号等显示格式原样保留。
  const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = worksheet ? worksheet[cellAddress] : null;
  if (cell) {
    const shown = XLSX.utils.format_cell(cell);
    if (shown != null && normalizeCellText(shown) !== "") return shown;
    if (cell.v != null) return cell.v;
  }

  const row = rows[rowIndex] || [];
  return row[colIndex];
}

function parseMonthFromCell(worksheet, rows, rowIndex, colIndex) {
  const value = getCellDisplayValue(worksheet, rows, rowIndex, colIndex);

  const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = worksheet ? worksheet[cellAddress] : null;
  if (
    cell &&
    typeof cell.v === "number" &&
    (cell.t === "n" || cell.t === "d")
  ) {
    const parsed = XLSX.SSF.parse_date_code(cell.v);
    if (parsed && parsed.m >= 1 && parsed.m <= 12) {
      return parsed.m;
    }
  }

  const text = normalizeCellText(value);
  if (!text) return -1;

  const monthHit = text.match(/(\d{1,2})\s*月/);
  if (monthHit) {
    const month = Number(monthHit[1]);
    return month >= 1 && month <= 12 ? month : -1;
  }

  const fullDate = text.match(/(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
  if (fullDate) {
    const month = Number(fullDate[2]);
    return month >= 1 && month <= 12 ? month : -1;
  }

  const monthOnly = text.match(/^(\d{1,2})$/);
  if (monthOnly) {
    const month = Number(monthOnly[1]);
    return month >= 1 && month <= 12 ? month : -1;
  }

  return -1;
}

function parseAmount(value) {
  if (typeof value === "number") return value;
  const text = normalizeCellText(value);
  if (!text) return 0;

  const cleaned = text.replace(/,/g, "").replace(/[^0-9.\-]/g, "");
  const n = Number.parseFloat(cleaned);
  return Number.isNaN(n) ? 0 : n;
}

function exportResultExcel() {
  if (!canExport.value) return;

  const aoa = [];
  aoa.push(["新签数据"]);
  aoa.push([...TABLE_HEADERS]);
  for (const row of resultData.value.newSignRows) {
    aoa.push([
      row.serial,
      row.customerName,
      row.projectName,
      row.owner,
      row.coOwner,
      row.product,
      row.amount,
    ]);
  }

  aoa.push([]);
  aoa.push(["营收数据"]);
  aoa.push([...TABLE_HEADERS]);
  for (const row of resultData.value.revenueRows) {
    aoa.push([
      row.serial,
      row.customerName,
      row.projectName,
      row.owner,
      row.coOwner,
      row.product,
      row.amount,
    ]);
  }

  aoa.push([]);
  aoa.push(["回款数据"]);
  aoa.push([...TABLE_HEADERS]);
  for (const row of resultData.value.receiptRows) {
    aoa.push([
      row.serial,
      row.customerName,
      row.projectName,
      row.owner,
      row.coOwner,
      row.product,
      row.amount,
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [
    { wch: 8 },
    { wch: 22 },
    { wch: 28 },
    { wch: 24 },
    { wch: 24 },
    { wch: 28 },
    { wch: 12 },
  ];

  // 中文注释：尝试设置文本列左对齐、金额列右对齐；若读写库不支持样式则自动忽略。
  applyAlignment(ws, aoa);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "三表填充结果");

  const name = `三表填充结果_${resultData.value.targetMonth || "上月"}月.xlsx`;
  XLSX.writeFile(wb, name);
}

function applyAlignment(ws, aoa) {
  for (let r = 0; r < aoa.length; r++) {
    for (let c = 0; c < (aoa[r] || []).length; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (!cell) continue;

      if (c === 6) {
        cell.s = { alignment: { horizontal: "right" } };
      } else {
        cell.s = { alignment: { horizontal: "left" } };
      }
    }
  }
}

function getPreviousMonth() {
  const now = new Date();
  const m = now.getMonth() + 1;
  if (m === 1) return { month: 12 };
  return { month: m - 1 };
}

function normalizeCellText(value) {
  return String(value == null ? "" : value)
    .replace(/\r?\n/g, "")
    .trim();
}

function normalizeHeader(value) {
  return normalizeCellText(value)
    .replace(/\s+/g, "")
    .replace(/（/g, "(")
    .replace(/）/g, ")");
}

function roundHalfUp(value) {
  const n = Number(value);
  if (Number.isNaN(n)) return 0;
  return n >= 0 ? Math.round(n) : -Math.round(Math.abs(n));
}
</script>

<style scoped>
.in-hand-three-sheet-page {
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.table-wrap {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.result-table {
  width: 100%;
  border-collapse: collapse;
  background: #fff;
}

.result-table th,
.result-table td {
  border: 1px solid #dbe1ea;
  padding: 6px 8px;
  text-align: left;
}

.result-table th {
  background: #f1f5f9;
}

.result-table td.amount {
  text-align: right;
}
</style>
