<template>
  <section class="overdue-page">
    <h3>逾期回款月报</h3>
    <p class="desc">
      上传 Excel 后，系统会识别“新签/营收/回款/业绩台账”四个子表，但仅处理“回款”子表。
    </p>

    <div class="upload-row">
      <label for="overdueFile">上传 Excel 文件：</label>
      <input id="overdueFile" type="file" accept=".xlsx,.xls" @change="onFileChange" />
    </div>

    <div v-if="statusText" class="panel">
      <strong>识别状态：</strong>
      <div class="content">{{ statusText }}</div>
      <ul>
        <li v-for="(v, k) in sheetRecognition" :key="k">{{ k }}：{{ v }}</li>
      </ul>
    </div>

    <div v-if="resultText" class="panel">
      <strong>结果：</strong>
      <div class="content result-only">{{ resultText }}</div>
    </div>
  </section>
</template>

<script setup>
// 中文注释：页面只在浏览器端读取本地 Excel 文件，不进行网络上传。
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const REQUIRED_SHEETS = ['新签', '营收', '回款', '业绩台账']
const RECEIPT_DATE_COL = '回款月日'
const RECEIPT_OVERDUE_COL = '是否是逾期回款'
const RECEIPT_AMOUNT_COL = '归属华东金额'

const store = useStore()
const statusText = ref(store.state.overdueReceiptMonthlyResult.statusText || '')
const sheetRecognition = ref(store.state.overdueReceiptMonthlyResult.sheetRecognition || {})
const resultText = ref(store.state.overdueReceiptMonthlyResult.resultText || '')

// 中文注释：将计算状态与结果同步到 Vuex，保证刷新后仍能回显。
watch(
  [statusText, sheetRecognition, resultText],
  () => {
    store.commit('setOverdueReceiptMonthlyResult', {
      statusText: statusText.value,
      sheetRecognition: sheetRecognition.value,
      resultText: resultText.value
    })
  },
  { deep: true }
)

function onFileChange(event) {
  const file = event.target.files && event.target.files[0]
  if (!file) return

  statusText.value = ''
  sheetRecognition.value = {}
  resultText.value = ''

  const reader = new FileReader()
  reader.onload = (ev) => {
    try {
      const data = new Uint8Array(ev.target.result)
      const workbook = XLSX.read(data, { type: 'array' })

      // 中文注释：先识别四个子表是否存在，再继续处理“回款”。
      const sheetMap = recognizeSheets(workbook)
      sheetRecognition.value = buildSheetRecognitionText(sheetMap)

      const receiptSheetName = sheetMap['回款']
      if (!receiptSheetName) {
        statusText.value = '未识别到“回款”子表，无法继续处理。'
        return
      }

      // 中文注释：只读取“回款”工作表内容，其他表不参与计算。
      const worksheet = workbook.Sheets[receiptSheetName]
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
      if (!rows.length) {
        statusText.value = '“回款”子表为空。'
        return
      }

      const headerRowIndex = findHeaderRowIndex(rows, [RECEIPT_DATE_COL, RECEIPT_OVERDUE_COL, RECEIPT_AMOUNT_COL])
      if (headerRowIndex === -1) {
        statusText.value = '未识别到“回款月日/是否是逾期回款/归属华东金额”列标题。'
        return
      }

      const header = (rows[headerRowIndex] || []).map((cell) => normalizeCellText(cell))
      const dateColIndex = header.indexOf(RECEIPT_DATE_COL)
      const overdueColIndex = header.indexOf(RECEIPT_OVERDUE_COL)
      const amountColIndex = header.indexOf(RECEIPT_AMOUNT_COL)

      const targetMonth = getPreviousMonthNumber()
      let amountSum = 0

      for (let r = headerRowIndex + 1; r < rows.length; r++) {
        const dateDisplay = getCellDisplayValue(worksheet, rows, r, dateColIndex)
        const overdueDisplay = getCellDisplayValue(worksheet, rows, r, overdueColIndex)
        const amountDisplay = getCellDisplayValue(worksheet, rows, r, amountColIndex)

        // 中文注释：按“回款月日”筛选上个月数据，仅比较月份（x月）。
        const monthNumber = parseMonthFromValue(dateDisplay, worksheet, r, dateColIndex)
        if (monthNumber !== targetMonth) continue

        // 中文注释：逾期状态严格匹配“是”。
        if (normalizeCellText(overdueDisplay) !== '是') continue

        const amountNumber = parseAmountNumber(amountDisplay)
        if (!Number.isNaN(amountNumber)) {
          amountSum += amountNumber
        }
      }

      const overdueAmountWan = Math.floor(amountSum / 10000)
      resultText.value = `本月完成逾期回款${overdueAmountWan}万元`
      statusText.value = `处理完成：已从“回款”子表筛选${targetMonth}月逾期回款并完成计算。`
    } catch (error) {
      statusText.value = `解析失败：${error && error.message ? error.message : String(error)}`
    }
  }

  reader.readAsArrayBuffer(file)
}

function recognizeSheets(workbook) {
  const sheetNames = workbook.SheetNames || []
  const map = {}
  for (const key of REQUIRED_SHEETS) {
    const matched = sheetNames.find((name) => normalizeCellText(name).includes(key))
    map[key] = matched || ''
  }
  return map
}

function buildSheetRecognitionText(sheetMap) {
  const output = {}
  for (const key of REQUIRED_SHEETS) {
    output[key] = sheetMap[key] ? `已识别（${sheetMap[key]}）` : '未识别'
  }
  return output
}

function findHeaderRowIndex(rows, requiredColumns) {
  let bestIndex = -1
  let bestHitCount = 0

  for (let i = 0; i < rows.length; i++) {
    const rowTexts = (rows[i] || []).map((cell) => normalizeCellText(cell))
    const hitCount = requiredColumns.filter((name) => rowTexts.includes(name)).length

    if (hitCount > bestHitCount) {
      bestHitCount = hitCount
      bestIndex = i
    }
  }

  return bestHitCount === requiredColumns.length ? bestIndex : -1
}

function normalizeCellText(value) {
  return String(value == null ? '' : value).replace(/\r?\n/g, '').trim()
}

function getCellDisplayValue(worksheet, rows, rowIndex, colIndex) {
  if (colIndex < 0) return ''

  // 中文注释：优先使用 format_cell 读取格式化展示值，确保百分号等显示内容可被保留。
  const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
  const cell = worksheet ? worksheet[cellAddress] : null
  if (cell) {
    const shown = XLSX.utils.format_cell(cell)
    if (shown != null && normalizeCellText(shown) !== '') return shown
    if (cell.v != null) return cell.v
  }

  const row = rows[rowIndex] || []
  return row[colIndex]
}

function getPreviousMonthNumber() {
  const now = new Date()
  const currentMonth = now.getMonth() + 1
  return currentMonth === 1 ? 12 : currentMonth - 1
}

function parseMonthFromValue(displayValue, worksheet, rowIndex, colIndex) {
  const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
  const cell = worksheet ? worksheet[cellAddress] : null

  // 中文注释：若是 Excel 日期序列号，先按日期类型解析，提高兼容性。
  if (cell && typeof cell.v === 'number' && (cell.t === 'n' || cell.t === 'd')) {
    const parsed = XLSX.SSF.parse_date_code(cell.v)
    if (parsed && parsed.m >= 1 && parsed.m <= 12) {
      return parsed.m
    }
  }

  if (displayValue instanceof Date) {
    return displayValue.getMonth() + 1
  }

  const text = normalizeCellText(displayValue)
  if (!text) return -1

  const monthHit = text.match(/(\d{1,2})\s*月/)
  if (monthHit) {
    const month = Number(monthHit[1])
    return month >= 1 && month <= 12 ? month : -1
  }

  // 中文注释：兼容 yyyy-mm-dd / yyyy/mm/dd / mm-dd / mm/dd 等常见字符串格式。
  const dateMatch = text.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/) || text.match(/^(\d{1,2})[-\/](\d{1,2})$/)
  if (dateMatch) {
    const month = Number(dateMatch.length === 4 ? dateMatch[2] : dateMatch[1])
    return month >= 1 && month <= 12 ? month : -1
  }

  const monthOnly = text.match(/^(\d{1,2})$/)
  if (monthOnly) {
    const month = Number(monthOnly[1])
    return month >= 1 && month <= 12 ? month : -1
  }

  return -1
}

function parseAmountNumber(displayValue) {
  if (typeof displayValue === 'number') return displayValue
  const text = normalizeCellText(displayValue)
  if (!text) return NaN

  const cleaned = text.replace(/,/g, '').replace(/[^0-9.\-]/g, '')
  const numericValue = Number.parseFloat(cleaned)
  return Number.isNaN(numericValue) ? NaN : numericValue
}
</script>

<style scoped>
.overdue-page {
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

.result-only {
  font-weight: 700;
  color: #0a6a3a;
}
</style>
