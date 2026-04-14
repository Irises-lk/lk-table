<template>
  <div class="app-shell">
    <h2>华东区域经营月报一键生成</h2>

    <section class="upload-panel">
      <div class="upload-row">
        <label for="allFiles">一次性上传全部源文件（支持多选）：</label>
        <input
          id="allFiles"
          type="file"
          accept=".xlsx,.xls"
          multiple
          @change="onAllFilesChange"
        />
        <button class="generate-btn" @click="generateAll">
          一键生成全部页面结果
        </button>
                <button type="button" class="generate-btn" @click="exportWordReport">
          生成 Word 月报
        </button>
      </div>

      <!-- <div class="upload-row word-row">
        <label>Word 模板来源：</label>
        <span class="template-path">src/template/华东区域经营月报模板.docx</span>
        <button type="button" class="generate-btn" @click="exportWordReport">
          生成 Word 月报
        </button>
      </div> -->
      <div v-if="wordStatus" class="word-status">{{ wordStatus }}</div>
      <details class="word-tags">
        <summary>Word 可用占位符清单</summary>
        <p class="rules-tip">
          中文注释：在 Word 中直接输入花括号占位符，例如
          {ledgerReportText}，导出时会自动替换为页面结果。
        </p>
        <ul class="tag-list">
          <li v-for="tag in wordTagList" :key="tag">{{ tag }}</li>
        </ul>
      </details>

      <details class="word-tags">
        <summary>三表表格模板写法（真实表格循环）</summary>
        <p class="rules-tip">
          中文注释：在 Word
          表格中，先建好表头，再在“数据行”7个单元格中按以下占位符放置，生成时会自动按行扩展。
        </p>
        <pre class="preview-value">
新签数据行：
{#inHandThreeSheetNewSignRows}{serial}{customerName}{projectName}{owner}{coOwner}{product}{amount}{/inHandThreeSheetNewSignRows}

营收数据行：
{#inHandThreeSheetRevenueRows}{serial}{customerName}{projectName}{owner}{coOwner}{product}{amount}{/inHandThreeSheetRevenueRows}

回款数据行：
{#inHandThreeSheetReceiptRows}{serial}{customerName}{projectName}{owner}{coOwner}{product}{amount}{/inHandThreeSheetReceiptRows}</pre
        >
      </details>

      <details class="word-tags">
        <summary>业绩台账表格模板写法（固定8行手动占位符）</summary>
        <p class="rules-tip">
          中文注释：按固定 8 行写占位符，不使用循环；字段对应 G 到 AB 共 22 列。
        </p>
        <pre class="preview-value">
示例：第1行占位符
{ledgerRawR1G}{ledgerRawR1H}{ledgerRawR1I}...{ledgerRawR1AB}

示例：第8行占位符
{ledgerRawR8G}{ledgerRawR8H}{ledgerRawR8I}...{ledgerRawR8AB}</pre
        >
      </details>

      <details class="word-preview">
        <summary>占位符内容预览</summary>
        <p class="rules-tip">
          中文注释：这里展示每个占位符当前会写入 Word 的实际内容。
        </p>
        <div class="preview-grid">
          <div
            v-for="item in wordPreviewList"
            :key="item.key"
            class="preview-card"
          >
            <div class="preview-key">{{ item.key }}</div>
            <pre class="preview-value">{{ item.value || "（空）" }}</pre>
          </div>
        </div>
      </details>

      <details class="rules-config">
        <summary>匹配规则配置（可编辑）</summary>
        <p class="rules-tip">
          中文注释：每行表示一组匹配规则；同一行使用 +
          连接多个关键字表示“同时包含”。
        </p>
        <div class="rules-grid">
          <div v-for="item in ruleMeta" :key="item.key" class="rule-item">
            <label :for="`rule-${item.key}`">{{ item.title }}</label>
            <textarea
              :id="`rule-${item.key}`"
              v-model="editableRules[item.key]"
              rows="2"
              :placeholder="item.placeholder"
            />
          </div>
        
        </div>
          <div class="rule-item">
            <label for="rule-ledger-rows">业绩台账指定行号（固定读取）</label>
            <input
              id="rule-ledger-rows"
              v-model="editableRules.ledgerRawTargetRows"
              type="text"
              placeholder="例如：6,10,11,14,17,18,19,20"
            />
          </div>
        <div class="rules-actions">
          <button type="button" @click="resetRules">恢复默认规则</button>
        </div>
      </details>

      <ul class="match-list">
        <li v-for="item in fileMatchView" :key="item.title">
          <span class="item-title">{{ item.title }}：</span>
          <span :class="item.matched ? 'ok' : 'warn'">{{ item.fileName }}</span>
        </li>
      </ul>
    </section>

    <section class="page-block">
      <LedgerReportGenerator
        :external-file="matchedFiles.ledger"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>

    <section class="page-block">
      <LedgerRawTableViewer
        :external-file="matchedFiles.ledger"
        :generate-key="generateKey"
        :hide-uploader="true"
        :target-rows-text="editableRules.ledgerRawTargetRows"
      />
    </section>

    <section class="page-block">
      <OverdueReceiptMonthlyReport
        :external-file="matchedFiles.ledger"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>

    <section class="page-block">
      <EastRegionOverdueStock
        :external-file="matchedFiles.stockOverdue"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>

    <section class="page-block">
      <EastRegionNewOverdue
        :external-file="matchedFiles.newOverdue"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>

    <section class="page-block">
      <MainInHandContractAmount
        :external-file="matchedFiles.inHand"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>

    <section class="page-block">
      <InHandThreeSheetTableBuilder
        :external-file="matchedFiles.ledger"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>

    <section class="page-block">
      <ProjectTrackingSummary
        :external-file="matchedFiles.tracking"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>

    <section class="page-block">
      <StageDynamicsSummary
        :external-file="matchedFiles.stageDynamics"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>

    <section class="page-block">
      <BidImplementationSummary
        :ledger-external-file="matchedFiles.ledger"
        :in-hand-external-file="matchedFiles.inHand"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>

    <section class="page-block">
      <OutboundProjectFollowup
        :external-file="matchedFiles.outbound"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>

    <section class="page-block">
      <CompetitorAnalysisStrict
        :external-file="matchedFiles.competitor"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>
    <section class="page-block">
      <NextMonthForecastSummary
        :external-file="matchedFiles.forecast"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>

    <section class="page-block">
      <OverdueCompressionSummary :auto-build-key="generateKey" />
    </section>
  </div>
</template>

<script setup>
import { computed, ref, watch } from "vue";
import { useStore } from "vuex";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import NextMonthForecastSummary from "./components/NextMonthForecastSummary.vue";
import LedgerReportGenerator from "./components/LedgerReportGenerator.vue";
import LedgerRawTableViewer from "./components/LedgerRawTableViewer.vue";
import OverdueReceiptMonthlyReport from "./components/OverdueReceiptMonthlyReport.vue";
import EastRegionOverdueStock from "./components/EastRegionOverdueStock.vue";
import EastRegionNewOverdue from "./components/EastRegionNewOverdue.vue";
import MainInHandContractAmount from "./components/MainInHandContractAmount.vue";
import InHandThreeSheetTableBuilder from "./components/InHandThreeSheetTableBuilder.vue";
import ProjectTrackingSummary from "./components/ProjectTrackingSummary.vue";
import StageDynamicsSummary from "./components/StageDynamicsSummary.vue";
import BidImplementationSummary from "./components/BidImplementationSummary.vue";
import OutboundProjectFollowup from "./components/OutboundProjectFollowup.vue";
import CompetitorAnalysisStrict from "./components/CompetitorAnalysisStrict.vue";
import OverdueCompressionSummary from "./components/OverdueCompressionSummary.vue";

const selectedFiles = ref([]);
const generateKey = ref(0);
const RULE_STORAGE_KEY = "lkgzl_rule_editor_v1";
const LOCAL_WORD_TEMPLATE_URL = new URL("./template/华东区域经营月报模板.docx", import.meta.url).href;
const LEDGER_FIXED_ROW_COUNT = 8;
const LEDGER_RAW_COL_LETTERS = [
  "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB"
];
const store = useStore();
const wordStatus = ref("");

const wordTagList = [
  "lastMonth",
  "nextMonthForecastText",
  "competitorStrictText",
  "ledgerReportText",
  "overdueReceiptText",
  "stockOverdueText",
  "newOverdueText",
  "inHandContractText",
  "projectTrackingText",
  "stageDynamicsText",
  "bidImplementationText",
  "outboundFollowupText",
  "overdueCompressionText",
  // 表格
  "inHandThreeSheetTableText",
  "inHandThreeSheetNewSignTableText",
  "inHandThreeSheetRevenueTableText",
  "inHandThreeSheetReceiptTableText",
  "inHandThreeSheetNewSignRows",
  "inHandThreeSheetRevenueRows",
  "inHandThreeSheetReceiptRows",
  "ledgerRawTableText",
  ...buildLedgerRawFixedTagList(),
  // 'competitorAnalysisText',
  // 'excelParserText'
];

const wordPreviewList = computed(() => {
  const data = buildWordData();
  return wordTagList.map((key) => {
    return {
      key,
      value: formatWordPreviewValue(data[key]),
    };
  });
});

// 中文注释：规则元数据用于渲染可编辑配置区。
const ruleMeta = [
  {
    key: "forecast",
    title: "次月指标预计汇总",
    placeholder: "月份预计指标表\n预计指标表",
  },
  {
    key: "competitor",
    title: "竞争对手承揽统计（含严格版）",
    placeholder: "竞争对手经营承揽情况",
  },
  {
    key: "ledger",
    title: "业绩台账（含完成逾期/中标实施）",
    placeholder: "业绩台账",
  },
  { key: "stockOverdue", title: "存量逾期", placeholder: "存量逾期" },
  { key: "newOverdue", title: "新增逾期", placeholder: "新增逾期" },
  {
    key: "inHand",
    title: "在手主办合同总额（含中标实施）",
    placeholder: "在手合同台账\n在手合同",
  },
  {
    key: "tracking",
    title: "信息跟踪阶段统计",
    placeholder: "项目信息更新核对",
  },
  {
    key: "stageDynamics",
    title: "三阶段动态统计",
    placeholder: "重点项目进度信息表",
  },
  {
    key: "outbound",
    title: "中资外带项目信息跟进",
    placeholder: "属地企业带出去项目信息跟进反馈表\n带出去项目信息跟进反馈表",
  },
];

const defaultEditableRules = {
  forecast: "月份预计指标表\n预计指标表",
  competitor: "竞争对手经营承揽情况",
  ledger: "业绩台账",
  ledgerRawTargetRows: "6,10,11,14,17,18,19,20",
  stockOverdue: "存量逾期",
  newOverdue: "新增逾期",
  inHand: "在手合同台账\n在手合同",
  tracking: "项目信息更新核对",
  stageDynamics: "重点项目进度信息表",
  outbound: "属地企业带出去项目信息跟进反馈表\n带出去项目信息跟进反馈表",
};

const editableRules = ref(loadEditableRules());

const parsedRules = computed(() => {
  const output = {};
  for (const item of ruleMeta) {
    output[item.key] = parseRuleText(editableRules.value[item.key]);
  }
  return output;
});

watch(
  editableRules,
  () => {
    // 中文注释：规则编辑后写入本地存储，避免刷新丢失。
    localStorage.setItem(RULE_STORAGE_KEY, JSON.stringify(editableRules.value));
  },
  { deep: true },
);

const matchedFiles = computed(() => {
  return {
    forecast: pickByRules(selectedFiles.value, parsedRules.value.forecast),
    competitor: pickByRules(selectedFiles.value, parsedRules.value.competitor),
    ledger: pickByRules(selectedFiles.value, parsedRules.value.ledger),
    stockOverdue: pickByRules(
      selectedFiles.value,
      parsedRules.value.stockOverdue,
    ),
    newOverdue: pickByRules(selectedFiles.value, parsedRules.value.newOverdue),
    inHand: pickByRules(selectedFiles.value, parsedRules.value.inHand),
    tracking: pickByRules(selectedFiles.value, parsedRules.value.tracking),
    stageDynamics: pickByRules(
      selectedFiles.value,
      parsedRules.value.stageDynamics,
    ),
    outbound: pickByRules(selectedFiles.value, parsedRules.value.outbound),
  };
});

const fileMatchView = computed(() => {
  return [
    {
      title: "次月指标预计汇总",
      fileName: fileNameOrMissing(matchedFiles.value.forecast),
      matched: !!matchedFiles.value.forecast,
    },
    {
      title: "竞争对手承揽统计（含严格版）",
      fileName: fileNameOrMissing(matchedFiles.value.competitor),
      matched: !!matchedFiles.value.competitor,
    },
    {
      title: "业绩台账/完成逾期/中标实施阶段（业绩台账）",
      fileName: fileNameOrMissing(matchedFiles.value.ledger),
      matched: !!matchedFiles.value.ledger,
    },
    {
      title: "存量逾期",
      fileName: fileNameOrMissing(matchedFiles.value.stockOverdue),
      matched: !!matchedFiles.value.stockOverdue,
    },
    {
      title: "新增逾期",
      fileName: fileNameOrMissing(matchedFiles.value.newOverdue),
      matched: !!matchedFiles.value.newOverdue,
    },
    {
      title: "在手主办合同总额/中标实施阶段（在手合同）",
      fileName: fileNameOrMissing(matchedFiles.value.inHand),
      matched: !!matchedFiles.value.inHand,
    },
    {
      title: "信息跟踪阶段统计",
      fileName: fileNameOrMissing(matchedFiles.value.tracking),
      matched: !!matchedFiles.value.tracking,
    },
    {
      title: "三阶段动态统计",
      fileName: fileNameOrMissing(matchedFiles.value.stageDynamics),
      matched: !!matchedFiles.value.stageDynamics,
    },
    {
      title: "中资外带项目信息跟进",
      fileName: fileNameOrMissing(matchedFiles.value.outbound),
      matched: !!matchedFiles.value.outbound,
    },
  ];
});

function onAllFilesChange(event) {
  const files = Array.from((event.target && event.target.files) || []);
  selectedFiles.value = files;
}

function generateAll() {
  // 中文注释：通过递增 key 触发各子组件的 watch，从而统一开始解析。
  generateKey.value += 1;
}

function fileNameOrMissing(file) {
  return file ? file.name : "未匹配到文件";
}

function pickByRules(files, rules) {
  if (!Array.isArray(files) || !files.length) return null;

  for (const file of files) {
    const normalized = normalizeText(file.name);
    for (const keywords of rules) {
      const hit = keywords.every((keyword) =>
        normalized.includes(normalizeText(keyword)),
      );
      if (hit) return file;
    }
  }
  return null;
}

function normalizeText(text) {
  return String(text == null ? "" : text)
    .replace(/\s+/g, "")
    .replace(/[()（）\[\]【】_.-]/g, "")
    .toLowerCase();
}

function parseRuleText(text) {
  const lines = String(text == null ? "" : text)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line);

  // 中文注释：每行支持“关键词A+关键词B”语法，表示文件名需同时包含多个关键字。
  return lines.map((line) => {
    return line
      .split("+")
      .map((part) => part.trim())
      .filter((part) => part);
  });
}

function loadEditableRules() {
  try {
    const raw = localStorage.getItem(RULE_STORAGE_KEY);
    if (!raw) return { ...defaultEditableRules };
    const parsed = JSON.parse(raw);
    return {
      ...defaultEditableRules,
      ...(parsed && typeof parsed === "object" ? parsed : {}),
    };
  } catch (error) {
    return { ...defaultEditableRules };
  }
}

function resetRules() {
  editableRules.value = { ...defaultEditableRules };
}

async function exportWordReport() {
  try {
    const response = await fetch(LOCAL_WORD_TEMPLATE_URL);
    if (!response.ok) {
      throw new Error(`本地模板读取失败：${response.status}`);
    }

    const arrayBuffer = await response.arrayBuffer();
    const zip = new PizZip(arrayBuffer);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    doc.render(buildWordData());

    const outputBlob = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    const name = `华东区经营月报（${formatDateCN()}).docx`;
    downloadBlob(outputBlob, name);
    wordStatus.value = `Word 已生成：${name}`;
  } catch (error) {
    wordStatus.value = `Word 生成失败：${error && error.message ? error.message : String(error)}`;
  }
}

function buildWordData() {
  const state = store.state;
  const threeSheetData = state.inHandThreeSheetTableResult?.data || {
    targetMonth: "",
    newSignRows: [],
    revenueRows: [],
    receiptRows: [],
  };
  const ledgerRawRows = Array.isArray(state.ledgerRawTableResult?.tableRows)
    ? state.ledgerRawTableResult.tableRows
    : [];

  const newSignText = buildThreeSheetSectionText(
    "新签数据",
    threeSheetData.newSignRows,
  );
  const revenueText = buildThreeSheetSectionText(
    "营收数据",
    threeSheetData.revenueRows,
  );
  const receiptText = buildThreeSheetSectionText(
    "回款数据",
    threeSheetData.receiptRows,
  );

  return {
    lastMonth: formatDateCN(),
    nextMonthForecastText: state.nextMonthForecastResult?.resultText || "",
    competitorStrictText:
      state.competitorAnalysisStrictResult?.resultText || "",
    ledgerReportText: state.ledgerReportResult?.resultText || "",
    overdueReceiptText: state.overdueReceiptMonthlyResult?.resultText || "",
    stockOverdueText: state.eastRegionOverdueStockResult?.resultText || "",
    newOverdueText: state.eastRegionNewOverdueResult?.resultText || "",
    inHandContractText: state.mainInHandContractAmountResult?.resultText || "",
    projectTrackingText: state.projectTrackingSummaryResult?.resultText || "",
    stageDynamicsText: state.stageDynamicsSummaryResult?.resultText || "",
    bidImplementationText:
      state.bidImplementationSummaryResult?.resultText || "",
    outboundFollowupText: state.outboundProjectFollowupResult?.resultText || "",
    overdueCompressionText:
      state.overdueCompressionSummaryResult?.resultText || "",
    inHandThreeSheetTableText: [newSignText, revenueText, receiptText]
      .filter((item) => item)
      .join("\n\n"),
    inHandThreeSheetNewSignTableText: newSignText,
    inHandThreeSheetRevenueTableText: revenueText,
    inHandThreeSheetReceiptTableText: receiptText,
    inHandThreeSheetNewSignRows: buildWordRows(threeSheetData.newSignRows),
    inHandThreeSheetRevenueRows: buildWordRows(threeSheetData.revenueRows),
    inHandThreeSheetReceiptRows: buildWordRows(threeSheetData.receiptRows),
    ledgerRawTableText: buildLedgerRawTableText(ledgerRawRows),
    ...buildLedgerRawFixedData(ledgerRawRows),
    // competitorAnalysisText: state.competitorAnalysisResult?.resultText || '',
    // excelParserText: state.excelParserResult?.message || ''
  };
}

function buildWordRows(rows) {
  const safeRows = Array.isArray(rows) ? rows : [];
  return safeRows.map((row) => ({
    serial: String(row.serial == null ? "" : row.serial),
    customerName: String(row.customerName == null ? "" : row.customerName),
    projectName: String(row.projectName == null ? "" : row.projectName),
    owner: String(row.owner == null ? "" : row.owner),
    coOwner: String(row.coOwner == null ? "" : row.coOwner),
    product: String(row.product == null ? "" : row.product),
    amount: String(row.amount == null ? "" : row.amount),
  }));
}

function buildLedgerRawFixedTagList() {
  const tags = [];
  for (let row = 1; row <= LEDGER_FIXED_ROW_COUNT; row += 1) {
    for (const col of LEDGER_RAW_COL_LETTERS) {
      tags.push(`ledgerRawR${row}${col}`);
    }
  }
  return tags;
}

function buildLedgerRawFixedData(rows) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const result = {};

  // 中文注释：固定输出8行占位符，不足行或空值统一补空字符串。
  for (let row = 1; row <= LEDGER_FIXED_ROW_COUNT; row += 1) {
    const current = Array.isArray(safeRows[row - 1]) ? safeRows[row - 1] : [];
    for (let colIndex = 0; colIndex < LEDGER_RAW_COL_LETTERS.length; colIndex += 1) {
      const col = LEDGER_RAW_COL_LETTERS[colIndex];
      result[`ledgerRawR${row}${col}`] = String(current[colIndex] == null ? "" : current[colIndex]);
    }
  }

  return result;
}

function buildLedgerRawTableText(rows) {
  const safeRows = Array.isArray(rows) ? rows : [];
  if (!safeRows.length) return "";

  const header = [];
  for (let col = 6; col <= 27; col += 1) {
    header.push(columnIndexToExcelLetter(col));
  }

  const lines = [header.join("\t")];
  for (const row of safeRows) {
    const normalized = Array.isArray(row) ? row : [];
    const line = [];
    for (let col = 0; col < 22; col += 1) {
      line.push(String(normalized[col] == null ? "" : normalized[col]));
    }
    lines.push(line.join("\t"));
  }

  return lines.join("\n");
}

function columnIndexToExcelLetter(index) {
  let num = index + 1;
  let letter = "";
  while (num > 0) {
    const mod = (num - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    num = Math.floor((num - mod) / 26);
  }
  return letter;
}

function formatWordPreviewValue(value) {
  if (value == null) return "";
  if (Array.isArray(value)) {
    if (!value.length) return "[]";
    return JSON.stringify(value, null, 2);
  }
  if (typeof value === "object") {
    return JSON.stringify(value, null, 2);
  }
  return String(value);
}

function buildThreeSheetSectionText(title, rows) {
  const header = [
    "序号",
    "客户名称",
    "项目名称",
    "对应责任方",
    "对应协办方",
    "产品名称和数量",
    "对应金额",
  ];

  const safeRows = Array.isArray(rows) ? rows : [];
  const lines = [];
  lines.push(title);
  lines.push(header.join("\t"));

  if (!safeRows.length) {
    lines.push("（无数据）");
    return lines.join("\n");
  }

  for (const row of safeRows) {
    lines.push(
      [
        row.serial,
        row.customerName,
        row.projectName,
        row.owner,
        row.coOwner,
        row.product,
        row.amount,
      ]
        .map((value) => String(value == null ? "" : value))
        .join("\t"),
    );
  }

  return lines.join("\n");
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}

function formatDateCN(date) {
  // const y = date.getFullYear()
  // const m = String(date.getMonth() + 1).padStart(2, '0')
  // const d = String(date.getDate()).padStart(2, '0')
  // return `${y}年${m}月${d}日`
  const lastMonth = new Date(new Date().setDate(0)).getMonth() + 1;
  return `${lastMonth}月`;
}

function formatDateCompact(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}${m}${d}`;
}
</script>

<style scoped>
.app-shell {
  --brand-700: #14532d;
  --brand-600: #166534;
  --brand-500: #1f9d65;
  --teal-600: #0f766e;
  --teal-500: #0ea5a0;
  --slate-900: #0f172a;
  --slate-700: #334155;
  --slate-500: #64748b;
  --line: #dbe3ee;
  --soft-bg: #f8fafc;
  --card-bg: #ffffff;
  max-width: 1500px;
  margin: 22px auto;
  padding: 0 14px 24px;
  font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
  background:
    radial-gradient(circle at 0% 0%, rgba(34, 197, 94, 0.08) 0%, transparent 45%),
    radial-gradient(circle at 100% 10%, rgba(14, 165, 233, 0.07) 0%, transparent 40%),
    linear-gradient(180deg, #f5faf7 0%, #f8fbff 100%);
}

.app-shell h2 {
  margin: 0 0 14px;
  padding: 12px 16px;
  border-radius: 12px;
  background: linear-gradient(135deg, var(--brand-600) 0%, var(--brand-500) 100%);
  box-shadow: 0 10px 22px rgba(22, 101, 52, 0.2);
  color: #fff;
  letter-spacing: 0.6px;
  font-size: 22px;
}

.upload-panel {
  padding: 14px;
  border: 1px solid var(--line);
  border-radius: 14px;
  background: var(--card-bg);
  box-shadow: 0 8px 26px rgba(15, 23, 42, 0.06);
  margin-bottom: 18px;
}

.upload-row {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}

.word-row {
  margin-top: 10px;
}

.template-path {
  padding: 7px 10px;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  background: var(--soft-bg);
  color: var(--slate-700);
  font-size: 13px;
}

.word-status {
  margin-top: 10px;
  padding: 8px 10px;
  border: 1px solid #cce8da;
  border-radius: 8px;
  background: #f1fbf4;
  color: #166534;
  font-size: 13px;
}

.word-tags {
  margin-top: 10px;
  border: 1px solid var(--line);
  border-radius: 10px;
  background: #f8fbff;
  padding: 10px;
}

.word-preview {
  margin-top: 10px;
  border: 1px solid var(--line);
  border-radius: 10px;
  background: #ffffff;
  padding: 10px;
}

.preview-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
  gap: 10px;
}

.preview-card {
  border: 1px solid #e2e8f0;
  border-radius: 10px;
  background: #fbfdff;
  padding: 10px;
}

.preview-key {
  font-weight: 700;
  color: var(--slate-900);
  margin-bottom: 6px;
}

.preview-value {
  margin: 0;
  white-space: pre-wrap;
  word-break: break-word;
  line-height: 1.65;
  color: var(--slate-700);
  font-size: 13px;
}

.tag-list {
  margin: 10px 0 0;
  padding-left: 18px;
  column-count: 3;
  column-gap: 20px;
}

.rules-config {
  margin-top: 10px;
  border: 1px solid var(--line);
  border-radius: 10px;
  background: #fff;
  padding: 10px;
}

.rules-config summary,
.word-tags summary,
.word-preview summary {
  cursor: pointer;
  font-weight: 600;
  color: var(--slate-900);
}

.rules-tip {
  margin: 8px 0;
  color: var(--slate-500);
  font-size: 13px;
  line-height: 1.6;
}

.rules-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
  gap: 10px;
}

.rule-item {
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.rule-item textarea {
  width: 90%;
  resize: vertical;
  min-height: 56px;
  padding: 8px;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  background: #f8fafc;
  color: #0f172a;
}

.rule-item input {
  width: 300px;
  padding: 8px;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  background: #f8fafc;
  color: #0f172a;
}

.rules-actions {
  margin-top: 8px;
}

.generate-btn {
  margin-top: 6px;
  min-width: 210px;
  padding: 9px 16px;
  border: 0;
  border-radius: 10px;
  color: #fff;
  background: linear-gradient(135deg, var(--teal-600) 0%, var(--teal-500) 100%);
  box-shadow: 0 8px 18px rgba(14, 116, 110, 0.28);
  cursor: pointer;
}

.generate-btn:hover {
  transform: translateY(-1px);
  filter: brightness(1.06);
}

.match-list {
  margin-top: 12px;
  padding-left: 18px;
  line-height: 1.75;
}

.item-title {
  color: var(--slate-900);
  font-weight: 600;
}

.ok {
  color: #146b2e;
}

.warn {
  color: #a41515;
}

.page-block {
  margin-bottom: 14px;
  padding: 12px;
  border: 1px solid var(--line);
  border-radius: 14px;
  background: #fff;
  box-shadow: 0 6px 20px rgba(15, 23, 42, 0.05);
}

/* 中文注释：使用深度选择器统一子组件页面视觉，减少逐个组件维护成本。 */
.page-block :deep(h3) {
  margin: 0 0 10px;
  padding-left: 10px;
  border-left: 4px solid var(--teal-500);
  font-size: 18px;
  color: var(--slate-900);
}

.page-block :deep(.desc) {
  margin: 0 0 10px;
  color: #4b5563;
}

.page-block :deep(.upload-row),
.page-block :deep(.actions) {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
  margin-bottom: 10px;
}

.page-block :deep(input[type="file"]) {
  padding: 6px;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  background: #f8fafc;
}

.page-block :deep(button) {
  padding: 6px 12px;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  background: #f8fafc;
  cursor: pointer;
}

.page-block :deep(button:hover) {
  background: #eef2ff;
}

.page-block :deep(.panel),
.page-block :deep(.result-panel),
.page-block :deep(.result-text) {
  padding: 12px;
  border: 1px solid #cce6dc;
  border-radius: 10px;
  background: #f3fbf7;
  color: #14532d;
  line-height: 1.8;
  font-weight: 700;
}

.page-block :deep(.error-panel),
.page-block :deep(.error-text) {
  padding: 12px;
  border: 1px solid #fecaca;
  border-radius: 10px;
  background: #fef2f2;
  color: #991b1b;
  line-height: 1.8;
  font-weight: 700;
}

.page-block :deep(.content),
.page-block :deep(.text-block),
.page-block :deep(.result-only) {
  margin-top: 6px;
  white-space: pre-wrap;
}

@media (max-width: 992px) {
  .tag-list {
    column-count: 2;
  }
}

@media (max-width: 720px) {
  .app-shell {
    margin: 14px auto;
    padding: 0 10px 18px;
  }

  .app-shell h2 {
    font-size: 18px;
    padding: 10px 12px;
  }

  .upload-panel {
    padding: 10px;
  }

  .generate-btn {
    min-width: 160px;
  }

  .preview-grid,
  .rules-grid {
    grid-template-columns: 1fr;
  }

  .tag-list {
    column-count: 1;
  }
}
</style>
