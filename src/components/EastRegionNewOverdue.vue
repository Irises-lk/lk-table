<template>
  <section class="new-overdue-page">
    <h3>华东区域指挥部新增逾期</h3>
    <div v-if="!hideUploader" class="upload-row">
      <label for="newOverdueFile">上传 Excel 文件：</label>
      <input id="newOverdueFile" type="file" accept=".xlsx,.xls" @change="onFileChange" />
    </div>

    <div v-if="resultText" class="result-text">{{ resultText }}</div>
    <div v-else-if="errorText" class="error-text">{{ errorText }}</div>
  </section>
</template>

<script setup>
// 中文注释：仅在浏览器端读取用户上传文件，避免改动后端与其他页面逻辑。
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const REQUIRED_SHEETS = ['汇总表', '华东区域指挥部主办逾期', '华东区域指挥部协办逾期']
const TARGET_ROW_TITLE = '华东区域指挥部合计'
const TARGET_COL_TITLE = '合计金额'

const store = useStore()
const resultText = ref(store.state.eastRegionNewOverdueResult.resultText || '')
const errorText = ref(store.state.eastRegionNewOverdueResult.errorText || '')
const props = defineProps({
  externalFile: { type: Object, default: null },
  generateKey: { type: Number, default: 0 },
  hideUploader: { type: Boolean, default: false }
})

// 中文注释：结果变化后写入 Vuex，确保页面刷新后仍展示最近一次计算值。
watch([resultText, errorText], () => {
  store.commit('setEastRegionNewOverdueResult', {
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

      // 中文注释：先识别三个子表，然后仅处理“汇总表”。
      const sheetMap = recognizeSheets(workbook)
      const summarySheetName = sheetMap['汇总表']
      if (!summarySheetName) {
        errorText.value = '未识别到“汇总表”工作表。'
        return
      }

      const worksheet = workbook.Sheets[summarySheetName]
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

      // 中文注释：优先读取 Excel 格式化显示值，保证百分号等显示格式可原样保留。
      const rawDisplayValue = getCellDisplayValue(worksheet, rows, targetRowIndex, amountColIndex)
      const decimalValue = toDecimal(rawDisplayValue)

      // 中文注释：按需求仅输出指定句子，不追加其他业务文本。
      resultText.value = `华东区域指挥部当前新增逾期为 ${decimalValue} 万元`
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
  // 中文注释：优先在目标行上方寻找列标题，通常更接近真实表头。
  for (let i = rowIndex - 1; i >= 0; i--) {
    const header = (rows[i] || []).map((cell) => normalizeCellText(cell))
    const colIndex = header.indexOf(title)
    if (colIndex !== -1) return colIndex
  }

  // 中文注释：兜底全表扫描，兼容非标准排版。
  for (let i = 0; i < rows.length; i++) {
    const header = (rows[i] || []).map((cell) => normalizeCellText(cell))
    const colIndex = header.indexOf(title)
    if (colIndex !== -1) return colIndex
  }

  return -1
}

function getCellDisplayValue(worksheet, rows, rowIndex, colIndex) {
  if (colIndex < 0) return ''

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

function normalizeCellText(value) {
  return String(value == null ? '' : value).replace(/\r?\n/g, '').trim()
}

function toDecimal(value) {
  if (typeof value === 'number') return value.toFixed(2)

  const text = normalizeCellText(value)
  if (!text) return '0.00'

  const cleaned = text.replace(/,/g, '').replace(/[^0-9.\-]/g, '')
  const numericValue = Number.parseFloat(cleaned)
  if (Number.isNaN(numericValue)) return '0.00'

  return numericValue.toFixed(2)
}
</script>

<style scoped>
.new-overdue-page {
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
