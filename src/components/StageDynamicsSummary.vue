<template>
  <section class="stage-page">
    <h3>三阶段项目动态统计</h3>

    <div v-if="!hideUploader" class="upload-row">
      <label for="stageFile">上传 Excel 文件：</label>
      <input id="stageFile" type="file" accept=".xlsx,.xls" @change="onFileChange" />
    </div>

    <div v-if="resultText" class="result-panel">{{ resultText }}</div>
    <div v-else-if="errorText" class="error-panel">{{ errorText }}</div>
  </section>
</template>

<script setup>
// 中文注释：本页面按三个阶段子表统计项目动态并输出固定文案。
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const STAGE_KEYS = ['信息跟踪阶段', '正在投标阶段', '中标实施阶段']
const REQUIRED_COLUMNS = ['项目名称', '项目动态']

const store = useStore()
const resultText = ref(store.state.stageDynamicsSummaryResult?.resultText || '')
const errorText = ref(store.state.stageDynamicsSummaryResult?.errorText || '')
const props = defineProps({
  externalFile: { type: Object, default: null },
  generateKey: { type: Number, default: 0 },
  hideUploader: { type: Boolean, default: false }
})

watch([resultText, errorText], () => {
  store.commit('setStageDynamicsSummaryResult', {
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
      const sheetMap = recognizeStageSheets(workbook)

      const missing = STAGE_KEYS.filter((key) => !sheetMap[key])
      if (missing.length) {
        errorText.value = `未识别到以下子表：${missing.join('、')}`
        return
      }

      const infoStats = analyzeStage(workbook.Sheets[sheetMap['信息跟踪阶段']], '信息跟踪阶段')
      const biddingStats = analyzeStage(workbook.Sheets[sheetMap['正在投标阶段']], '正在投标阶段')
      const implementationStats = analyzeStage(workbook.Sheets[sheetMap['中标实施阶段']], '中标实施阶段')

      const aa = infoStats.totalProjects
      const bb = infoStats.newCount
      const cc = infoStats.updatedCount
      const dd = aa - bb - cc

      const ee = biddingStats.totalProjects
      const ff = biddingStats.newCount
      const gg = biddingStats.updatedCount
      const hh = ee - ff - gg

      const ii = implementationStats.totalProjects
      const jj = implementationStats.newCount
      const kk = implementationStats.updatedCount
      const ll = ii - jj - kk

      resultText.value = [
        `另重点跟踪阶段项目${aa}个，新增项目${bb}个，更新进展项目${cc}个，未更新进展项目${dd}个`,
        `正在投标阶段项目共计${ee}个，本月新增项目${ff}个，共计更新项目${gg}个。另重点投标阶段项目${ee}个，新增项目${ff}个，共计更新进展项目${gg}个，未更新进展项目${hh}个`,
        `另重点中标实施阶段项目${ii}个，新增项目${jj}个，更新进展项目${kk}个，未更新进展项目${ll}个`
      ].join('\n')
    } catch (error) {
      errorText.value = `解析失败：${error && error.message ? error.message : String(error)}`
    }
  }

  reader.readAsArrayBuffer(file)
}

function recognizeStageSheets(workbook) {
  const map = {}
  const names = workbook.SheetNames || []
  for (const key of STAGE_KEYS) {
    const hit = names.find((name) => normalizeCellText(name).includes(key))
    map[key] = hit || ''
  }
  return map
}

function analyzeStage(worksheet, stageName) {
  cancelWorksheetFilters(worksheet)
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
  if (!rows.length) {
    throw new Error(`子表“${stageName}”为空。`)
  }

  const headerInfo = findHeaderRow(rows, REQUIRED_COLUMNS)
  if (!headerInfo) {
    throw new Error(`子表“${stageName}”未识别到列标题“项目名称/项目动态”。`)
  }

  const { headerRowIndex, colIndexMap } = headerInfo
  let totalProjects = 0
  let newCount = 0
  let updatedCount = 0

  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const projectName = normalizeCellText(getCellDisplayValue(worksheet, rows, r, colIndexMap['项目名称']))
    if (!projectName) continue

    totalProjects++

    const dynamicText = normalizeCellText(getCellDisplayValue(worksheet, rows, r, colIndexMap['项目动态']))
    if (dynamicText === '新增') newCount++
    if (dynamicText === '更新') updatedCount++
  }

  return { totalProjects, newCount, updatedCount }
}

function cancelWorksheetFilters(worksheet) {
  if (!worksheet) return

  // 中文注释：移除筛选和隐藏标记，避免遗漏被筛选隐藏的数据。
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

  // 中文注释：优先读取格式化展示值，保证百分号等格式可原样读取。
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
</script>

<style scoped>
.stage-page {
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
  white-space: pre-wrap;
  line-height: 1.8;
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
