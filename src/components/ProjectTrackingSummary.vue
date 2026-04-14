<template>
  <section class="tracking-page">
    <h3>信息跟踪阶段统计</h3>

    <div v-if="!hideUploader" class="upload-row">
      <label for="trackingFile">上传 Excel 文件：</label>
      <input id="trackingFile" type="file" accept=".xlsx,.xls" @change="onFileChange" />
    </div>

    <div v-if="resultText" class="result-panel">{{ resultText }}</div>
    <div v-else-if="errorText" class="error-panel">{{ errorText }}</div>
  </section>
</template>

<script setup>
// 中文注释：本页面按要求统计“在跟进项目/本月新增/本月更新”，并将结果写入 Vuex。
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const REQUIRED_COLUMNS = ['取消跟进原因', '项目名称', '最后修改时间', '创建时间']

const store = useStore()
const resultText = ref(store.state.projectTrackingSummaryResult?.resultText || '')
const errorText = ref(store.state.projectTrackingSummaryResult?.errorText || '')
const props = defineProps({
  externalFile: { type: Object, default: null },
  generateKey: { type: Number, default: 0 },
  hideUploader: { type: Boolean, default: false }
})

watch([resultText, errorText], () => {
  store.commit('setProjectTrackingSummaryResult', {
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

      // 中文注释：默认读取第一个工作表作为项目表。
      const firstSheetName = (workbook.SheetNames || [])[0]
      if (!firstSheetName) {
        errorText.value = '未识别到任何工作表。'
        return
      }

      const worksheet = workbook.Sheets[firstSheetName]
      cancelWorksheetFilters(worksheet)

      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
      if (!rows.length) {
        errorText.value = '工作表为空。'
        return
      }

      const headerInfo = findHeaderRow(rows, REQUIRED_COLUMNS)
      if (!headerInfo) {
        errorText.value = '未识别到必需列：取消跟进原因、项目名称、最后修改时间、创建时间。'
        return
      }

      const { headerRowIndex, colIndexMap } = headerInfo
      const validRows = []
      for (let r = headerRowIndex + 1; r < rows.length; r++) {
        const reasonText = normalizeCellText(getCellDisplayValue(worksheet, rows, r, colIndexMap['取消跟进原因']))
        if (reasonText !== '') continue
        validRows.push(r)
      }

      const prev = getPreviousMonthYear()

      // 中文注释：aa 按有效行中的“项目名称”计数，空项目名不计入。
      let aa = 0
      let bb = 0
      let cc = 0

      for (const rowIndex of validRows) {
        const projectName = normalizeCellText(getCellDisplayValue(worksheet, rows, rowIndex, colIndexMap['项目名称']))
        if (projectName) aa++

        const modifiedDate = parseDateCell(worksheet, rows, rowIndex, colIndexMap['最后修改时间'])
        if (isTargetMonth(modifiedDate, prev.year, prev.month)) cc++

        const createdDate = parseDateCell(worksheet, rows, rowIndex, colIndexMap['创建时间'])
        if (isTargetMonth(createdDate, prev.year, prev.month)) bb++
      }

      resultText.value = `信息跟踪阶段项目共计${aa}个，本月新增项目${bb}个，共计更新项目${cc}个。`
    } catch (error) {
      errorText.value = `解析失败：${error && error.message ? error.message : String(error)}`
    }
  }

  reader.readAsArrayBuffer(file)
}

function cancelWorksheetFilters(worksheet) {
  if (!worksheet) return

  // 中文注释：取消自动筛选并解除隐藏行，避免遗漏被筛选隐藏的数据。
  delete worksheet['!autofilter']
  if (Array.isArray(worksheet['!rows'])) {
    worksheet['!rows'] = worksheet['!rows'].map((row) => {
      if (!row || typeof row !== 'object') return row
      return { ...row, hidden: false }
    })
  }
}

function findHeaderRow(rows, requiredColumns) {
  for (let i = 0; i < rows.length; i++) {
    const header = (rows[i] || []).map((cell) => normalizeCellText(cell))
    const colIndexMap = {}
    let allHit = true

    for (const columnName of requiredColumns) {
      const idx = header.indexOf(columnName)
      if (idx === -1) {
        allHit = false
        break
      }
      colIndexMap[columnName] = idx
    }

    if (allHit) {
      return { headerRowIndex: i, colIndexMap }
    }
  }
  return null
}

function getCellDisplayValue(worksheet, rows, rowIndex, colIndex) {
  if (colIndex < 0) return ''

  // 中文注释：优先读取格式化显示值，确保百分号等展示格式保持原样。
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

  // 中文注释：优先按 Excel 原始日期序列值解析，避免字符串格式差异造成误判。
  if (cell && typeof cell.v === 'number' && (cell.t === 'n' || cell.t === 'd')) {
    const parsed = XLSX.SSF.parse_date_code(cell.v)
    if (parsed && parsed.y && parsed.m && parsed.d) {
      return new Date(parsed.y, parsed.m - 1, parsed.d)
    }
  }

  const displayValue = getCellDisplayValue(worksheet, rows, rowIndex, colIndex)
  if (displayValue instanceof Date) return displayValue

  const text = normalizeCellText(displayValue)
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
.tracking-page {
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
