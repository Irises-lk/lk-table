<template>
  <section class="stock-page">
    <h3>华东区域指挥部存量逾期</h3>
    <div v-if="!hideUploader" class="upload-row">
      <label for="stockFile">上传 Excel 文件：</label>
      <input id="stockFile" type="file" accept=".xlsx,.xls" @change="onFileChange" />
    </div>

    <div v-if="resultText" class="result-text">{{ resultText }}</div>
    <div v-else-if="errorText" class="error-text">{{ errorText }}</div>
  </section>
</template>

<script setup>
// 中文注释：仅在前端本地读取用户上传的 Excel，不做任何网络传输。
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const REQUIRED_SHEETS = ['汇总表', '华东区域指挥部主办逾期', '华东区域指挥部协办逾期']
const TARGET_ROW_TITLE = '华东区域指挥部合计'
const TARGET_COL_TITLE = '合计金额'
const TARGET_INITIAL_COL_TITLE = '期初逾期'

const store = useStore()
const resultText = ref(store.state.eastRegionOverdueStockResult.resultText || '')
const errorText = ref(store.state.eastRegionOverdueStockResult.errorText || '')
const rawAmount = ref(store.state.eastRegionOverdueStockResult.rawAmount)
const initialOverdue = ref(store.state.eastRegionOverdueStockResult.initialOverdue)
const props = defineProps({
  externalFile: { type: Object, default: null },
  generateKey: { type: Number, default: 0 },
  hideUploader: { type: Boolean, default: false }
})

// 中文注释：把页面展示结果持久化到 Vuex，避免路由切换或刷新后丢失。
watch([resultText, errorText, rawAmount, initialOverdue], () => {
  store.commit('setEastRegionOverdueStockResult', {
    resultText: resultText.value,
    errorText: errorText.value,
    rawAmount: rawAmount.value,
    initialOverdue: initialOverdue.value
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
  rawAmount.value = null
  initialOverdue.value = null

  const reader = new FileReader()
  reader.onload = (ev) => {
    try {
      const data = new Uint8Array(ev.target.result)
      const workbook = XLSX.read(data, { type: 'array' })

      // 中文注释：先识别三个子表是否存在，再且仅处理“汇总表”。
      const sheetMap = recognizeSheets(workbook)
      if (!sheetMap['汇总表']) {
        errorText.value = '未识别到“汇总表”工作表。'
        return
      }

      const worksheet = workbook.Sheets[sheetMap['汇总表']]
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
      if (!rows.length) {
        errorText.value = '“汇总表”为空。'
        return
      }

      const targetRowIndex = findRowIndexByTitle(rows, TARGET_ROW_TITLE)
      if (targetRowIndex === -1) {
        errorText.value = '未找到“华东区域指挥部合计”行。'
        return
      }

      const amountColIndex = findColumnIndex(rows, targetRowIndex, TARGET_COL_TITLE)
      if (amountColIndex === -1) {
        errorText.value = '未找到“合计金额”列。'
        return
      }

      const initialColIndex = findColumnIndex(rows, targetRowIndex, TARGET_INITIAL_COL_TITLE)
      if (initialColIndex === -1) {
        errorText.value = '未找到“期初逾期”列。'
        return
      }

      // 中文注释：优先读取单元格原始值，避免被Excel显示格式（如隐藏小数位）截断。
      const rawValue = getCellRawValue(worksheet, rows, targetRowIndex, amountColIndex)
      const decimalValue = toActualNumberText(rawValue)
      rawAmount.value = parseActualNumber(rawValue)

      // 中文注释：同步读取“期初逾期”原始值，供逾期压降汇总动态使用。
      const initialRawValue = getCellRawValue(worksheet, rows, targetRowIndex, initialColIndex)
      initialOverdue.value = parseActualNumber(initialRawValue)

      resultText.value = `华东区域指挥部当前存量逾期为 ${decimalValue} 万元`
    } catch (error) {
      errorText.value = `解析失败：${error && error.message ? error.message : String(error)}`
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

function findRowIndexByTitle(rows, title) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || []
    if (row.some((cell) => normalizeCellText(cell) === title)) {
      return i
    }
  }
  return -1
}

function findColumnIndex(rows, rowIndex, title) {
  // 中文注释：优先在目标行之上的行里查找标题行，选距离目标行最近的一行。
  for (let i = rowIndex - 1; i >= 0; i--) {
    const header = (rows[i] || []).map((cell) => normalizeCellText(cell))
    const colIndex = header.indexOf(title)
    if (colIndex !== -1) return colIndex
  }

  // 中文注释：兜底策略，若上方未找到则全表扫描标题列。
  for (let i = 0; i < rows.length; i++) {
    const header = (rows[i] || []).map((cell) => normalizeCellText(cell))
    const colIndex = header.indexOf(title)
    if (colIndex !== -1) return colIndex
  }

  return -1
}

function getCellRawValue(worksheet, rows, rowIndex, colIndex) {
  if (colIndex < 0) return ''

  const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
  const cell = worksheet ? worksheet[cellAddress] : null
  if (cell) {
    // 中文注释：优先返回原始值 cell.v，确保保留完整精度。
    if (cell.v != null && normalizeCellText(cell.v) !== '') return cell.v

    const shown = XLSX.utils.format_cell(cell)
    if (shown != null && normalizeCellText(shown) !== '') return shown
  }

  const row = rows[rowIndex] || []
  return row[colIndex]
}

function normalizeCellText(value) {
  return String(value == null ? '' : value).replace(/\r?\n/g, '').trim()
}

function toActualNumberText(value) {
  if (typeof value === 'number') return String(value)

  const text = normalizeCellText(value)
  if (!text) return '0'

  const cleaned = text.replace(/,/g, '').replace(/[^0-9.\-]/g, '')
  const numericValue = Number.parseFloat(cleaned)
  if (Number.isNaN(numericValue)) return text

  return String(numericValue)
}

function parseActualNumber(value) {
  if (typeof value === 'number') return value

  const text = normalizeCellText(value)
  if (!text) return NaN

  const cleaned = text.replace(/,/g, '').replace(/[^0-9.\-]/g, '')
  const numericValue = Number.parseFloat(cleaned)
  return Number.isNaN(numericValue) ? NaN : numericValue
}
</script>

<style scoped>
.stock-page {
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

.result-text {
  padding: 12px;
  border: 1px solid #c8e6c9;
  background: #f1fbf3;
  color: #146b2e;
  font-weight: 700;
}

.error-text {
  padding: 12px;
  border: 1px solid #ffcdd2;
  background: #fff5f5;
  color: #a41515;
}
</style>
