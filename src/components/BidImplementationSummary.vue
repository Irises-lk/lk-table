<template>
  <section class="bid-page">
    <h3>中标实施阶段汇总</h3>

    <div class="upload-row">
      <label for="ledgerFile">上传业绩台账文件：</label>
      <input id="ledgerFile" type="file" accept=".xlsx,.xls" @change="onLedgerFileChange" />
    </div>

    <div class="upload-row">
      <label for="inHandFile">上传在手合同台账文件：</label>
      <input id="inHandFile" type="file" accept=".xlsx,.xls" @change="onInHandFileChange" />
    </div>

    <div v-if="resultText" class="result-panel">{{ resultText }}</div>
    <div v-else-if="errorText" class="error-panel">{{ errorText }}</div>
  </section>
</template>

<script setup>
// 中文注释：本页面按需求统计去重后的合同编号数量，并仅输出指定句子。
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const LEDGER_SHEET_KEYS = ['新签', '收入', '回款']
const IN_HAND_SHEET_KEYS = ['在手合同台账（主办）']

const store = useStore()
const resultText = ref(store.state.bidImplementationSummaryResult?.resultText || '')
const errorText = ref(store.state.bidImplementationSummaryResult?.errorText || '')
const ledgerFileBuffer = ref(null)
const inHandFileBuffer = ref(null)

watch([resultText, errorText], () => {
  store.commit('setBidImplementationSummaryResult', {
    resultText: resultText.value,
    errorText: errorText.value
  })
})

function onLedgerFileChange(event) {
  const file = event.target.files && event.target.files[0]
  if (!file) return

  readExcelFile(file, (arrayBuffer) => {
    ledgerFileBuffer.value = arrayBuffer
    tryBuildResult()
  })
}

function onInHandFileChange(event) {
  const file = event.target.files && event.target.files[0]
  if (!file) return

  readExcelFile(file, (arrayBuffer) => {
    inHandFileBuffer.value = arrayBuffer
    tryBuildResult()
  })
}

function readExcelFile(file, onLoaded) {
  const reader = new FileReader()
  reader.onload = (ev) => {
    try {
      const data = new Uint8Array(ev.target.result)
      onLoaded(data)
    } catch (error) {
      errorText.value = `读取失败：${error && error.message ? error.message : String(error)}`
    }
  }
  reader.readAsArrayBuffer(file)
}

function tryBuildResult() {
  resultText.value = ''
  errorText.value = ''

  if (!ledgerFileBuffer.value || !inHandFileBuffer.value) {
    return
  }

  try {
    const ledgerWorkbook = XLSX.read(ledgerFileBuffer.value, { type: 'array' })
    const inHandWorkbook = XLSX.read(inHandFileBuffer.value, { type: 'array' })

    const ledgerSheetMap = recognizeSheets(ledgerWorkbook, LEDGER_SHEET_KEYS)
    const inHandSheetMap = recognizeSheets(inHandWorkbook, IN_HAND_SHEET_KEYS)

    const missingLedger = LEDGER_SHEET_KEYS.filter((key) => !ledgerSheetMap[key])
    const missingInHand = IN_HAND_SHEET_KEYS.filter((key) => !inHandSheetMap[key])
    if (missingLedger.length || missingInHand.length) {
      const allMissing = [...missingLedger, ...missingInHand]
      errorText.value = `未识别到以下子表：${allMissing.join('、')}`
      return
    }

    const prev = getPreviousMonthYear()

    const aa = countUniqueByMonth(
      ledgerWorkbook.Sheets[ledgerSheetMap['新签']],
      '合同签订日期',
      '合同编号',
      prev.year,
      prev.month
    )
    const bb = countUniqueByMonth(
      ledgerWorkbook.Sheets[ledgerSheetMap['收入']],
      '确认收入月日',
      '合同编号',
      prev.year,
      prev.month
    )
    const cc = countUniqueByMonth(
      ledgerWorkbook.Sheets[ledgerSheetMap['回款']],
      '回款月日',
      '合同编号',
      prev.year,
      prev.month
    )
    const dd = bb + cc

    const ee = countUniqueAll(
      inHandWorkbook.Sheets[inHandSheetMap['在手合同台账（主办）']],
      '合同编号'
    )

    resultText.value = `中标实施阶段项目共计${ee}个，本月新增项目${bb}个，共计更新项目${dd}个。`
  } catch (error) {
    errorText.value = `解析失败：${error && error.message ? error.message : String(error)}`
  }
}

function recognizeSheets(workbook, sheetKeys) {
  const map = {}
  const names = workbook.SheetNames || []
  for (const key of sheetKeys) {
    const hit = names.find((name) => normalizeCellText(name).includes(key))
    map[key] = hit || ''
  }
  return map
}

function countUniqueByMonth(worksheet, dateColumnName, contractColumnName, year, month) {
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
  if (!rows.length) return 0

  const headerInfo = findHeaderRow(rows, [dateColumnName, contractColumnName])
  if (!headerInfo) {
    throw new Error(`未识别到列标题“${dateColumnName}/${contractColumnName}”。`)
  }

  const dateColIndex = headerInfo.colIndexMap[dateColumnName]
  const contractColIndex = headerInfo.colIndexMap[contractColumnName]
  const uniqueContracts = new Set()

  for (let r = headerInfo.headerRowIndex + 1; r < rows.length; r++) {
    const dateValue = parseDateCell(worksheet, rows, r, dateColIndex)
    if (!isTargetMonth(dateValue, year, month)) continue

    const contractNo = normalizeCellText(getCellDisplayValue(worksheet, rows, r, contractColIndex))
    if (!contractNo) continue
    uniqueContracts.add(contractNo)
  }

  return uniqueContracts.size
}

function countUniqueAll(worksheet, contractColumnName) {
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
  if (!rows.length) return 0

  const headerInfo = findHeaderRow(rows, [contractColumnName])
  if (!headerInfo) {
    throw new Error(`未识别到列标题“${contractColumnName}”。`)
  }

  const contractColIndex = headerInfo.colIndexMap[contractColumnName]
  const uniqueContracts = new Set()

  for (let r = headerInfo.headerRowIndex + 1; r < rows.length; r++) {
    const contractNo = normalizeCellText(getCellDisplayValue(worksheet, rows, r, contractColIndex))
    if (!contractNo) continue
    uniqueContracts.add(contractNo)
  }

  return uniqueContracts.size
}

function findHeaderRow(rows, requiredColumns) {
  for (let i = 0; i < rows.length; i++) {
    const header = (rows[i] || []).map((cell) => normalizeCellText(cell))
    const colIndexMap = {}
    let allHit = true

    for (const name of requiredColumns) {
      const idx = header.indexOf(name)
      if (idx === -1) {
        allHit = false
        break
      }
      colIndexMap[name] = idx
    }

    if (allHit) {
      return { headerRowIndex: i, colIndexMap }
    }
  }
  return null
}

function getCellDisplayValue(worksheet, rows, rowIndex, colIndex) {
  if (colIndex < 0) return ''

  // 中文注释：优先返回格式化显示值，确保百分号等格式可原样读取。
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

function parseDateCell(worksheet, rows, rowIndex, colIndex) {
  if (colIndex < 0) return null

  const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
  const cell = worksheet ? worksheet[cellAddress] : null

  // 中文注释：优先按 Excel 日期序列解析，提高兼容性。
  if (cell && typeof cell.v === 'number' && (cell.t === 'n' || cell.t === 'd')) {
    const parsed = XLSX.SSF.parse_date_code(cell.v)
    if (parsed && parsed.y && parsed.m && parsed.d) {
      return new Date(parsed.y, parsed.m - 1, parsed.d)
    }
  }

  const display = getCellDisplayValue(worksheet, rows, rowIndex, colIndex)
  if (display instanceof Date) return display

  const text = normalizeCellText(display)
  if (!text) return null

  const direct = Date.parse(text)
  if (!Number.isNaN(direct)) return new Date(direct)

  const dateMatch = text.match(/(\d{4})[./-](\d{1,2})[./-](\d{1,2})/)
  if (dateMatch) {
    return new Date(Number(dateMatch[1]), Number(dateMatch[2]) - 1, Number(dateMatch[3]))
  }

  return null
}

function getPreviousMonthYear() {
  const now = new Date()
  const currentYear = now.getFullYear()
  const currentMonth = now.getMonth() + 1
  if (currentMonth === 1) {
    return { year: currentYear - 1, month: 12 }
  }
  return { year: currentYear, month: currentMonth - 1 }
}

function isTargetMonth(date, year, month) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return false
  return date.getFullYear() === year && date.getMonth() + 1 === month
}

function normalizeCellText(value) {
  return String(value == null ? '' : value).replace(/\r?\n/g, '').trim()
}
</script>

<style scoped>
.bid-page {
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.upload-row {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}

.result-panel {
  padding: 12px;
  border: 1px solid #c8e6c9;
  background: #f1fbf3;
  color: #146b2e;
  font-weight: 700;
}

.error-panel {
  padding: 12px;
  border: 1px solid #ffcdd2;
  background: #fff5f5;
  color: #a41515;
}
</style>
