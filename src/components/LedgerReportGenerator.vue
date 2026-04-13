<template>
  <section class="ledger-page">
    <h3>业绩台账文本生成</h3>
    <p class="desc">上传 Excel 后，仅解析“业绩台账”工作表，并定位“华东区域指挥部总计”行生成业务文本。</p>

    <div class="upload-row">
      <label for="excelFile">上传 Excel 文件：</label>
      <input id="excelFile" type="file" accept=".xlsx,.xls" @change="onFileChange" />
    </div>

    <div v-if="statusText" class="panel">
      <strong>识别状态：</strong>
      <div class="content">{{ statusText }}</div>
      <ul>
        <li v-for="(v, k) in sheetRecognition" :key="k">{{ k }}：{{ v }}</li>
      </ul>
    </div>

    <!-- <div v-if="rowData" class="panel">
      <strong>提取数据：</strong>
      <table class="data-table">
        <tbody>
          <tr v-for="(v, k) in rowData" :key="k">
            <td>{{ k }}</td>
            <td>{{ v }}</td>
          </tr>
        </tbody>
      </table>
    </div> -->

    <div v-if="resultText" class="panel">
      <strong>结果：</strong>
      <div class="content text-block">{{ resultText }}</div>
      <button @click="exportTxt">导出为 TXT</button>
    </div>
  </section>
</template>

<script setup>
// 中文注释：仅使用 SheetJS 读取用户上传的本地 Excel，不做网络请求。
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const REQUIRED_SHEETS = ['新签', '营收', '回款', '业绩台账']
const TARGET_ROW_TITLE = '华东区域指挥部总计'
const REQUIRED_COLUMNS = [
  '本月新签',
  '累计新签',
  '新签目标完成率',
  '新签同比增长率',
  '本月营收',
  '累计营收',
  '营收目标',
  '营收目标完成率',
  '营收同比增长率',
  '本月回款',
  '累计回款',
  '回款目标完成率',
  '回款同比增长率'
]

const store = useStore()
const statusText = ref(store.state.ledgerReportResult.statusText || '')
const sheetRecognition = ref(store.state.ledgerReportResult.sheetRecognition || {})
const rowData = ref(store.state.ledgerReportResult.rowData || null)
const resultText = ref(store.state.ledgerReportResult.resultText || '')

// 中文注释：统一把页面结果写入 Vuex，持久化后刷新或切路由仍可展示。
watch(
  [statusText, sheetRecognition, rowData, resultText],
  () => {
    store.commit('setLedgerReportResult', {
      statusText: statusText.value,
      sheetRecognition: sheetRecognition.value,
      rowData: rowData.value,
      resultText: resultText.value
    })
  },
  { deep: true }
)

function onFileChange(e) {
  const file = e.target.files && e.target.files[0]
  if (!file) return

  statusText.value = ''
  sheetRecognition.value = {}
  rowData.value = null
  resultText.value = ''

  const reader = new FileReader()
  reader.onload = (ev) => {
    try {
      const data = new Uint8Array(ev.target.result)
      const workbook = XLSX.read(data, { type: 'array' })

      // 中文注释：先识别四个子表是否存在（按“包含关键词”识别）。
      const sheetMap = recognizeSheets(workbook)
      sheetRecognition.value = buildSheetRecognitionText(sheetMap)

      if (!sheetMap['业绩台账']) {
        statusText.value = '未找到“业绩台账”工作表，无法继续解析。'
        return
      }

      // 中文注释：严格只读取“业绩台账”工作表，不读取其他表内容。
      const ws = workbook.Sheets[sheetMap['业绩台账']]
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
      if (!rows.length) {
        statusText.value = '“业绩台账”工作表为空。'
        return
      }

      const targetRowIndex = findTargetRowIndex(rows, TARGET_ROW_TITLE)
      if (targetRowIndex === -1) {
        statusText.value = '未找到行标题“华东区域指挥部总计”。'
        return
      }

      const headerRowIndex = findHeaderRowIndex(rows, targetRowIndex, REQUIRED_COLUMNS)
      if (headerRowIndex === -1) {
        statusText.value = '未能识别字段标题行，请检查“业绩台账”列标题。'
        return
      }

      const extracted = extractByHeader(ws, rows, headerRowIndex, targetRowIndex, REQUIRED_COLUMNS)
      const missingColumns = REQUIRED_COLUMNS.filter((name) => extracted[name] == null)
      if (missingColumns.length) {
        statusText.value = `缺少以下列标题：${missingColumns.join('、')}`
        return
      }

      // 中文注释：第三段文案需要“回款目标”，若表里有则读取，没有则显示“未提供”。
      const optionalTarget = extractByHeader(ws, rows, headerRowIndex, targetRowIndex, ['回款目标'])['回款目标']

      const normalized = normalizeData(extracted, optionalTarget)
      rowData.value = normalized
      resultText.value = composeBusinessText(normalized)
      statusText.value = '解析成功：已精准定位“业绩台账”中“华东区域指挥部总计”并完成文案填充。'
    } catch (error) {
      statusText.value = `解析失败：${error && error.message ? error.message : String(error)}`
    }
  }

  reader.readAsArrayBuffer(file)
}

function recognizeSheets(workbook) {
  const map = {}
  const names = workbook.SheetNames || []
  for (const key of REQUIRED_SHEETS) {
    const hit = names.find((n) => String(n).trim().includes(key))
    map[key] = hit || ''
  }
  return map
}

function buildSheetRecognitionText(sheetMap) {
  const result = {}
  for (const key of REQUIRED_SHEETS) {
    result[key] = sheetMap[key] ? `已识别（${sheetMap[key]}）` : '未识别'
  }
  return result
}

function normalizeCellText(v) {
  return String(v == null ? '' : v).replace(/\r?\n/g, '').trim()
}

function findTargetRowIndex(rows, rowTitle) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || []
    if (row.some((c) => normalizeCellText(c) === rowTitle)) {
      return i
    }
  }
  return -1
}

function findHeaderRowIndex(rows, targetRowIndex, requiredColumns) {
  let bestIndex = -1
  let bestCount = 0
  for (let i = 0; i < targetRowIndex; i++) {
    const headerTexts = (rows[i] || []).map((c) => normalizeCellText(c))
    const hitCount = requiredColumns.filter((col) => headerTexts.includes(col)).length
    if (hitCount > bestCount) {
      bestCount = hitCount
      bestIndex = i
    }
  }
  return bestCount > 0 ? bestIndex : -1
}

function extractByHeader(ws, rows, headerRowIndex, dataRowIndex, columns) {
  const header = (rows[headerRowIndex] || []).map((c) => normalizeCellText(c))
  const result = {}
  for (const col of columns) {
    const idx = header.indexOf(col)
    result[col] = idx >= 0 ? getCellDisplayValue(ws, rows, dataRowIndex, idx) : null
  }
  return result
}

function getCellDisplayValue(ws, rows, r, c) {
  // 中文注释：优先返回单元格格式化显示值，保证百分号等展示格式不丢失。
  const addr = XLSX.utils.encode_cell({ r, c })
  const cell = ws ? ws[addr] : null
  if (cell) {
    const shown = XLSX.utils.format_cell(cell)
    if (shown != null && String(shown).trim() !== '') return shown
    if (cell.v != null) return cell.v
  }
  const row = rows[r] || []
  return row[c]
}

function parseNumeric(value) {
  if (typeof value === 'number') return value
  const text = normalizeCellText(value)
  if (!text) return NaN
  const cleaned = text.replace(/,/g, '').replace(/%/g, '').replace(/[^0-9.\-]/g, '')
  const n = parseFloat(cleaned)
  return Number.isNaN(n) ? NaN : n
}

function formatAmount(value) {
  const n = parseNumeric(value)
  return Number.isNaN(n) ? '0' : String(Math.round(n))
}

function formatRate(value) {
  const text = normalizeCellText(value)
  const hadPercent = /%/.test(text)
  let n = parseNumeric(value)
  if (Number.isNaN(n)) return '0.00%'
  if (!hadPercent && Math.abs(n) <= 1) {
    // 中文注释：当原值为 0-1 小数时，按比例转百分值。
    n = n * 100
  }
  return `${n.toFixed(2)}%`
}

function normalizeData(extracted, optionalTarget) {
  return {
    本月新签: formatAmount(extracted['本月新签']),
    累计新签: formatAmount(extracted['累计新签']),
    新签目标完成率: formatRate(extracted['新签目标完成率']),
    新签同比增长率: formatRate(extracted['新签同比增长率']),
    本月营收: formatAmount(extracted['本月营收']),
    累计营收: formatAmount(extracted['累计营收']),
    营收目标: formatAmount(extracted['营收目标']),
    营收目标完成率: formatRate(extracted['营收目标完成率']),
    营收同比增长率: formatRate(extracted['营收同比增长率']),
    本月回款: formatAmount(extracted['本月回款']),
    累计回款: formatAmount(extracted['累计回款']),
    回款目标: optionalTarget == null ? '未提供' : formatAmount(optionalTarget),
    回款目标完成率: formatRate(extracted['回款目标完成率']),
    回款同比增长率: formatRate(extracted['回款同比增长率'])
  }
}

function composeBusinessText(data) {
  // 中文注释：百分比字段在格式化阶段已带“%”，模板中直接使用，避免重复拼接。
  return [
    `（一）新签合同额：本月累计完成${data.本月新签}万元，年度累计完成${data.累计新签}万元，同比增长${data.新签同比增长率}，完成年度目标70000万元的${data.新签目标完成率}。`,
    `（二）营业收入：本月累计完成${data.本月营收}万元，年度累计完成${data.累计营收}万元，同比增长${data.营收同比增长率}，完成年度目标71000万元的${data.营收目标完成率}。`,
    `（三）货款回笼：本月累计完成${data.本月回款}万元，年度累计完成${data.累计回款}万元，同比增长${data.回款同比增长率}，完成年度目标${data.回款目标}万元的${data.回款目标完成率}。`
  ].join('\n')
}

function exportTxt() {
  const content = resultText.value || ''
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = '业绩台账文本结果.txt'
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}
</script>

<style scoped>
.ledger-page {
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.desc {
  margin: 0;
  color: #666;
}

.upload-row {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}

.panel {
  padding: 12px;
  border: 1px solid #ddd;
  background: #fafafa;
}

.content {
  margin-top: 8px;
}

.text-block {
  white-space: pre-wrap;
  line-height: 1.7;
}

.data-table {
  width: 100%;
  border-collapse: collapse;
  margin-top: 8px;
}

.data-table td {
  border: 1px solid #ddd;
  padding: 6px 8px;
}

.data-table td:first-child {
  width: 220px;
  background: #f2f2f2;
}
</style>
