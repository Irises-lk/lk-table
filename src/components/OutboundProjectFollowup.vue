<template>
  <section class="outbound-page">
    <h3>中资外带项目信息跟进</h3>

    <div v-if="!hideUploader" class="upload-row">
      <label for="outboundFile">上传 Excel 文件：</label>
      <input id="outboundFile" type="file" accept=".xlsx,.xls" @change="onFileChange" />
    </div>

    <div v-if="resultText" class="result-panel">{{ resultText }}</div>
    <div v-else-if="errorText" class="error-panel">{{ errorText }}</div>
  </section>
</template>

<script setup>
// 中文注释：本页面统计项目动态，并把结果持久化到 Vuex。
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const REQUIRED_COLUMNS = ['项目名称', '项目动态']

const store = useStore()
const resultText = ref(store.state.outboundProjectFollowupResult?.resultText || '')
const errorText = ref(store.state.outboundProjectFollowupResult?.errorText || '')
const props = defineProps({
  externalFile: { type: Object, default: null },
  generateKey: { type: Number, default: 0 },
  hideUploader: { type: Boolean, default: false }
})

watch([resultText, errorText], () => {
  store.commit('setOutboundProjectFollowupResult', {
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
      const firstSheetName = (workbook.SheetNames || [])[0]
      if (!firstSheetName) {
        errorText.value = '未识别到工作表。'
        return
      }

      const worksheet = workbook.Sheets[firstSheetName]
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
      if (!rows.length) {
        errorText.value = '工作表为空。'
        return
      }

      const headerInfo = findHeaderRow(rows, REQUIRED_COLUMNS)
      if (!headerInfo) {
        errorText.value = '未识别到列标题“项目名称/项目动态”。'
        return
      }

      const nameColIndex = headerInfo.colIndexMap['项目名称']
      const dynamicColIndex = headerInfo.colIndexMap['项目动态']

      let aa = 0
      let bb = 0
      let cc = 0

      for (let r = headerInfo.headerRowIndex + 1; r < rows.length; r++) {
        const projectName = normalizeCellText(getCellDisplayValue(worksheet, rows, r, nameColIndex))
        if (!projectName) continue

        aa++

        const dynamicText = normalizeCellText(getCellDisplayValue(worksheet, rows, r, dynamicColIndex))
        if (dynamicText === '新增') bb++
        if (dynamicText === '更新') cc++
      }

      const dd = aa - bb - cc
      resultText.value = `（二）中资外带项目信息跟进情况\n中资外带项目共计${aa}个，新增项目${bb}个，更新进展项目${cc}个，未更新进展项目${dd}个。具体详情见附件二：《属地企业带出去项目信息跟进反馈表》。`
    } catch (error) {
      errorText.value = `解析失败：${error && error.message ? error.message : String(error)}`
    }
  }

  reader.readAsArrayBuffer(file)
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

  // 中文注释：优先读取格式化显示值，保证百分号等格式可原样保留。
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
.outbound-page {
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
