<template>
  <section class="strict-page">
    <h3>竞争对手承揽统计（严格版）</h3>

    <div v-if="!hideUploader" class="upload-row">
      <label for="strictFile">上传 Excel 文件：</label>
      <input id="strictFile" type="file" accept=".xlsx,.xls" @change="onFileChange" />
    </div>

    <div v-if="resultText" class="result-panel">{{ resultText }}</div>
    <div v-else-if="errorText" class="error-panel">{{ errorText }}</div>
  </section>
</template>

<script setup>
// 中文注释：本页面按严格规则统计“掘进机/特种装备”两张表，并只输出最终文案。
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const SHEET_KEYS = ['掘进机', '特种装备']

const store = useStore()
const resultText = ref(store.state.competitorAnalysisStrictResult?.resultText || '')
const errorText = ref(store.state.competitorAnalysisStrictResult?.errorText || '')
const props = defineProps({
  externalFile: { type: Object, default: null },
  generateKey: { type: Number, default: 0 },
  hideUploader: { type: Boolean, default: false }
})

watch([resultText, errorText], () => {
  store.commit('setCompetitorAnalysisStrictResult', {
    resultText: resultText.value,
    errorText: errorText.value
  })
})

watch(
  () => props.generateKey,
  () => {
    if (!props.externalFile) return
    onFileChange({ target: { files: [props.externalFile] } })
  }
)

function onFileChange(event) {
  const file = event.target.files && event.target.files[0]
  if (!file) return

  resultText.value = ''
  errorText.value = ''

  const reader = new FileReader()
  reader.onload = (ev) => {
    try {
      const data = new Uint8Array(ev.target.result)
      const workbook = XLSX.read(data, { type: 'array' })
      const sheetMap = recognizeSheets(workbook)

      const missing = SHEET_KEYS.filter((key) => !sheetMap[key])
      if (missing.length) {
        errorText.value = `未识别到以下工作表：${missing.join('、')}`
        return
      }

      const prev = getPreviousMonthYear()
      const machineStats = analyzeSheet(workbook.Sheets[sheetMap['掘进机']], prev.year, prev.month)
      const specialStats = analyzeSheet(workbook.Sheets[sheetMap['特种装备']], prev.year, prev.month)

      const aa = machineStats.totalAmount
      const bb = machineStats.tjzgAmount
      const cc = roundHalfUp(aa - bb)
      const gg = machineStats.prevMonthCount
      const jj = machineStats.totalUnits
      const kk = machineStats.tjzgUnits
      const ll = roundHalfUp(jj - kk)

      const dd = specialStats.totalAmount
      const ee = specialStats.tjzgAmount
      const ff = roundHalfUp(dd - ee)
      const hh = specialStats.prevMonthCount
      const mm = specialStats.totalUnits
      const nn = specialStats.tjzgUnits
      const oo = roundHalfUp(mm - nn)

      const xx = roundHalfUp(gg + hh)
      const ii = prev.month

      // 中文注释：按要求仅输出单行最终文案，不添加额外说明。
      resultText.value = `本月新增${xx}项竞争对手承揽统计与分析内容。2026年1-${ii}月，掘进机产品，我司承揽${kk}台，金额${bb}万元；竞争对手承揽${ll}台，金额${cc}万元。特种装备产品，我司承揽${nn}台，金额${ee}万元；竞争对手承揽${oo}台，金额${ff}万元。具体详情见附件：《竞争对手承揽统计与分析表》。`
    } catch (error) {
      errorText.value = `解析失败：${error && error.message ? error.message : String(error)}`
    }
  }

  reader.readAsArrayBuffer(file)
}

function recognizeSheets(workbook) {
  const names = workbook.SheetNames || []
  const map = {}
  for (const key of SHEET_KEYS) {
    const hit = names.find((name) => normalizeCellText(name).includes(key))
    map[key] = hit || ''
  }
  return map
}

function analyzeSheet(worksheet, targetYear, targetMonth) {
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
  if (!rows.length) {
    return { totalAmount: 0, prevMonthCount: 0, totalUnits: 0, tjzgAmount: 0, tjzgUnits: 0 }
  }

  const headerInfo = findHeaderRow(rows)
  if (!headerInfo) {
    throw new Error('未识别到必需列标题。')
  }

  const amountCol = headerInfo.colIndexMap['承揽合同额（万元）']
  const dateCol = headerInfo.colIndexMap['承揽确定时间（中标、竞谈或合同签订）']
  const unitsCol = headerInfo.colIndexMap['台 / 套数']
  const orgCol = headerInfo.colIndexMap['承揽单位']

  let totalAmount = 0
  let prevMonthCount = 0
  let totalUnits = 0
  let tjzgAmount = 0
  let tjzgUnits = 0

  for (let r = headerInfo.headerRowIndex + 1; r < rows.length; r++) {
    const amountText = getCellDisplayValue(worksheet, rows, r, amountCol)
    const unitsText = getCellDisplayValue(worksheet, rows, r, unitsCol)
    const orgText = normalizeCellText(getCellDisplayValue(worksheet, rows, r, orgCol))

    const amount = parseNumber(amountText)
    const units = parseNumber(unitsText)

    if (!Number.isNaN(amount)) totalAmount += amount
    if (!Number.isNaN(units)) totalUnits += units

    const dateValue = parseDateCell(worksheet, rows, r, dateCol)
    if (isTargetMonth(dateValue, targetYear, targetMonth)) {
      prevMonthCount++
    }

    if (orgText === '铁建重工') {
      if (!Number.isNaN(amount)) tjzgAmount += amount
      if (!Number.isNaN(units)) tjzgUnits += units
    }
  }

  return {
    totalAmount: roundHalfUp(totalAmount),
    prevMonthCount: roundHalfUp(prevMonthCount),
    totalUnits: roundHalfUp(totalUnits),
    tjzgAmount: roundHalfUp(tjzgAmount),
    tjzgUnits: roundHalfUp(tjzgUnits)
  }
}

function findHeaderRow(rows) {
  for (let i = 0; i < rows.length; i++) {
    const rawHeader = rows[i] || []
    const normalizedHeader = rawHeader.map((cell) => normalizeHeaderText(cell))
    const colIndexMap = {}

    const amountIdx = indexOfHeader(normalizedHeader, ['承揽合同额（万元）', '承揽合同额(万元)'])
    const dateIdx = indexOfHeader(normalizedHeader, ['承揽确定时间（中标、竞谈或合同签订）', '承揽确定时间(中标、竞谈或合同签订)'])
    const unitsIdx = indexOfHeader(normalizedHeader, ['台/套数', '台/套'])
    const orgIdx = indexOfHeader(normalizedHeader, ['承揽单位'])

    if (amountIdx === -1 || dateIdx === -1 || unitsIdx === -1 || orgIdx === -1) continue

    colIndexMap['承揽合同额（万元）'] = amountIdx
    colIndexMap['承揽确定时间（中标、竞谈或合同签订）'] = dateIdx
    colIndexMap['台 / 套数'] = unitsIdx
    colIndexMap['承揽单位'] = orgIdx

    return { headerRowIndex: i, colIndexMap }
  }
  return null
}

function indexOfHeader(header, candidates) {
  for (const candidate of candidates) {
    const idx = header.indexOf(normalizeHeaderText(candidate))
    if (idx !== -1) return idx
  }
  return -1
}

function getCellDisplayValue(worksheet, rows, rowIndex, colIndex) {
  if (colIndex < 0) return ''

  // 中文注释：优先读取格式化显示值，确保百分号等展示格式可原样保留。
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

function parseNumber(value) {
  if (typeof value === 'number') return value
  const text = normalizeCellText(value)
  if (!text) return NaN

  const cleaned = text.replace(/,/g, '').replace(/[^0-9.\-]/g, '')
  const n = Number.parseFloat(cleaned)
  return Number.isNaN(n) ? NaN : n
}

function getPreviousMonthYear() {
  const now = new Date()
  const y = now.getFullYear()
  const m = now.getMonth() + 1
  if (m === 1) return { year: y - 1, month: 12 }
  return { year: y, month: m - 1 }
}

function isTargetMonth(date, year, month) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return false
  return date.getFullYear() === year && date.getMonth() + 1 === month
}

function roundHalfUp(value) {
  const number = Number(value)
  if (Number.isNaN(number)) return 0
  return number >= 0 ? Math.round(number) : -Math.round(Math.abs(number))
}

function normalizeCellText(value) {
  return String(value == null ? '' : value).replace(/\r?\n/g, '').trim()
}

function normalizeHeaderText(value) {
  return normalizeCellText(value).replace(/\s+/g, '')
}
</script>

<style scoped>
.strict-page {
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
