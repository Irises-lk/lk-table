<template>
  <section class="ledger-raw-viewer-page">
    <h3>业绩台账数据展示（G 到 AB，固定8行）</h3>

    <div v-if="!hideUploader" class="upload-row">
      <label for="ledgerRawFile">上传 Excel 文件：</label>
      <input id="ledgerRawFile" type="file" accept=".xlsx,.xls" @change="onFileChange" />
    </div>

    <div v-if="statusText" class="result-panel">{{ statusText }}</div>
    <div v-else-if="errorText" class="error-panel">{{ errorText }}</div>

    <div v-if="tableRows.length" class="table-wrap">
      <table class="result-table">
        <thead>
          <tr>
            <th v-for="(col, index) in columnHeaders" :key="`head-${index}`">{{ col }}</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(row, rowIndex) in tableRows" :key="`row-${rowIndex}`">
            <td v-for="(cell, colIndex) in row" :key="`cell-${rowIndex}-${colIndex}`">
              {{ cell }}
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  </section>
</template>

<script setup>
// 中文注释：本组件仅用于展示“业绩台账”子表中固定区间（G~AB）的筛选数据。
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const props = defineProps({
  externalFile: { type: Object, default: null },
  generateKey: { type: Number, default: 0 },
  hideUploader: { type: Boolean, default: false },
  targetRowsText: { type: String, default: '6,10,11,14,17,18,19,20' }
})

const store = useStore()
const statusText = ref(store.state.ledgerRawTableResult?.statusText || '')
const errorText = ref(store.state.ledgerRawTableResult?.errorText || '')
const tableRows = ref(store.state.ledgerRawTableResult?.tableRows || [])
const columnHeaders = ref(buildColumnHeaders())

const START_COL_INDEX = 6
const END_COL_INDEX = 27

watch(
  () => props.generateKey,
  () => {
    if (!props.externalFile) return
    onFileChange({ target: { files: [props.externalFile] } })
  }
)

watch(
  [statusText, errorText, tableRows],
  () => {
    // 中文注释：将页面读取结果持久化到 Vuex，便于 Word 导出占位符填充。
    store.commit('setLedgerRawTableResult', {
      statusText: statusText.value,
      errorText: errorText.value,
      tableRows: tableRows.value
    })
  },
  { deep: true }
)

function onFileChange(event) {
  const file = event.target.files && event.target.files[0]
  if (!file) return

  statusText.value = ''
  errorText.value = ''
  tableRows.value = []

  const reader = new FileReader()
  reader.onload = (ev) => {
    try {
      const data = new Uint8Array(ev.target.result)
      const workbook = XLSX.read(data, { type: 'array' })

      const ledgerSheetName = findLedgerSheetName(workbook)
      if (!ledgerSheetName) {
        errorText.value = '未识别到“业绩台账”子表。'
        return
      }

      const worksheet = workbook.Sheets[ledgerSheetName]
      // 中文注释：raw=false 优先保留单元格显示格式（如百分号）。
      const allRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false })
      const targetRows = parseTargetRows(props.targetRowsText)
      if (!targetRows.length) {
        errorText.value = '指定行号无效，请输入如：6,10,11,14,17,18,19,20'
        return
      }

      const selectedRows = selectRowsByFixedRows(allRows, targetRows)
      tableRows.value = selectedRows
      statusText.value = `读取成功：已展示“${ledgerSheetName}”中指定行（${targetRows.join('、')}）的 G~AB 数据，共${selectedRows.length}行。`
    } catch (error) {
      errorText.value = `解析失败：${error && error.message ? error.message : String(error)}`
    }
  }

  reader.readAsArrayBuffer(file)
}

function findLedgerSheetName(workbook) {
  const names = workbook.SheetNames || []
  const key = normalizeCellText('业绩台账')
  return names.find((name) => normalizeCellText(name).includes(key)) || ''
}

function normalizeCellText(value) {
  return String(value == null ? '' : value).replace(/\r?\n/g, '').trim()
}

function parseTargetRows(text) {
  const unique = new Set()
  String(text == null ? '' : text)
    .split(/[，,\s]+/)
    .map((item) => item.trim())
    .filter((item) => item)
    .forEach((item) => {
      const num = parseInt(item, 10)
      if (Number.isInteger(num) && num > 0) {
        unique.add(num)
      }
    })

  return Array.from(unique)
}

function selectRowsByFixedRows(allRows, targetRows) {
  const result = []

  // 中文注释：仅按指定行号读取，行号按 Excel 习惯从 1 开始。
  for (const rowNumber of targetRows) {
    const rowIndex = rowNumber - 1
    const row = allRows[rowIndex] || []
    const selectedRow = []

    for (let colIndex = START_COL_INDEX; colIndex <= END_COL_INDEX; colIndex += 1) {
      selectedRow.push(row[colIndex] == null ? '' : row[colIndex])
    }

    result.push(selectedRow)
  }

  return result
}

function buildColumnHeaders() {
  const headers = []
  for (let colIndex = START_COL_INDEX; colIndex <= END_COL_INDEX; colIndex += 1) {
    headers.push(columnIndexToLetter(colIndex))
  }
  return headers
}

function columnIndexToLetter(index) {
  let num = index + 1
  let letter = ''
  while (num > 0) {
    const mod = (num - 1) % 26
    letter = String.fromCharCode(65 + mod) + letter
    num = Math.floor((num - mod) / 26)
  }
  return letter
}
</script>

<style scoped>
.ledger-raw-viewer-page {
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.table-wrap {
  overflow-x: auto;
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
  white-space: nowrap;
}

.result-table th {
  background: #f1f5f9;
}
</style>
