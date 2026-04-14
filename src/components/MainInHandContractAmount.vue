<template>
  <section class="main-in-hand-page">
    <h3>在手主办未拆分项目可执行合同总额</h3>

    <div v-if="!hideUploader" class="upload-row">
      <label for="mainInHandFile">上传 Excel 文件：</label>
      <input id="mainInHandFile" type="file" accept=".xlsx,.xls" @change="onFileChange" />
    </div>

    <div v-if="resultText" class="result-panel">{{ resultText }}</div>
    <div v-else-if="errorText" class="error-panel">{{ errorText }}</div>
  </section>
</template>

<script setup>
// 中文注释：本页只处理“在手合同台账（主办）”子表，并将结果写入 Vuex 以持久化。
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const REQUIRED_SHEETS = ['在手合同台账（主办）', '在手合同台账（协办）']
const TARGET_SHEET_KEY = '在手合同台账（主办）'
const TARGET_ROW_TITLE = '华东区域指挥部合计'
const TARGET_COL_TITLE = '尚未确认收入金额（含税）'
const TARGET_HEADER_ROW_INDEX = 2

const store = useStore()
const resultText = ref(store.state.mainInHandContractAmountResult?.resultText || '')
const errorText = ref(store.state.mainInHandContractAmountResult?.errorText || '')
const props = defineProps({
  externalFile: { type: Object, default: null },
  generateKey: { type: Number, default: 0 },
  hideUploader: { type: Boolean, default: false }
})

watch([resultText, errorText], () => {
  store.commit('setMainInHandContractAmountResult', {
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
      const sheetName = sheetMap[TARGET_SHEET_KEY]
      if (!sheetName) {
        errorText.value = '未识别到“在手合同台账（主办）”工作表。'
        return
      }

      const worksheet = workbook.Sheets[sheetName]
      cancelWorksheetFilters(worksheet)

      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
      if (!rows.length) {
        errorText.value = '“在手合同台账（主办）”工作表为空。'
        return
      }

      const headerRow = (rows[TARGET_HEADER_ROW_INDEX] || []).map((cell) => normalizeCellText(cell))
      const targetColIndex = headerRow.indexOf(TARGET_COL_TITLE)
      if (targetColIndex === -1) {
        errorText.value = '第3行未找到“尚未确认收入金额（含税）”列标题。'
        return
      }

      const targetRowIndex = findRowIndexByTitle(rows, TARGET_ROW_TITLE)
      if (targetRowIndex === -1) {
        errorText.value = '未找到“华东区域指挥部合计”行。'
        return
      }

      // 中文注释：交叉单元格优先按 Excel 格式化显示值读取，兼容百分号等格式。
      const rawDisplayValue = getCellDisplayValue(worksheet, rows, targetRowIndex, targetColIndex)
      const amountValue = parseAmount(rawDisplayValue)
      const amountWan = roundHalfUp(amountValue / 10000)

      resultText.value = `华在手主办未拆分项目可执行合同总额${amountWan}万元`
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
    const hit = sheetNames.find((name) => normalizeCellText(name).includes(key))
    map[key] = hit || ''
  }
  return map
}

function cancelWorksheetFilters(worksheet) {
  if (!worksheet) return

  // 中文注释：在内存中移除自动筛选与隐藏标记，确保后续读取拿到完整数据。
  delete worksheet['!autofilter']
  if (Array.isArray(worksheet['!rows'])) {
    worksheet['!rows'] = worksheet['!rows'].map((row) => {
      if (!row || typeof row !== 'object') return row
      return { ...row, hidden: false }
    })
  }
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

function parseAmount(value) {
  if (typeof value === 'number') return value
  const text = normalizeCellText(value)
  if (!text) return 0

  const cleaned = text.replace(/,/g, '').replace(/[^0-9.\-]/g, '')
  const n = Number.parseFloat(cleaned)
  return Number.isNaN(n) ? 0 : n
}

function roundHalfUp(value) {
  const number = Number(value)
  if (Number.isNaN(number)) return 0
  return number >= 0 ? Math.round(number) : -Math.round(Math.abs(number))
}
</script>

<style scoped>
.main-in-hand-page {
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

.error-panel {
  padding: 12px;
  border: 1px solid #ffcdd2;
  background: #fff5f5;
  color: #a41515;
}
</style>
