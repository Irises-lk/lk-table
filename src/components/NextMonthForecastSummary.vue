<template>
  <section class="forecast-page">
    <h3>次月指标预计汇总</h3>

    <div v-if="!hideUploader" class="upload-row">
      <label for="forecastFile">上传 Excel 文件：</label>
      <input id="forecastFile" type="file" accept=".xlsx,.xls" @change="onFileChange" />
    </div>

    <div v-if="resultText" class="result-panel">{{ resultText }}</div>
    <div v-else-if="errorText" class="error-panel">{{ errorText }}</div>
  </section>
</template>

<script setup>
// 中文注释：本页面按“新签/营收/回款”三张表统计并输出固定文案。
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const SHEET_KEYS = ['新签', '营收', '回款']

const store = useStore()
const resultText = ref(store.state.nextMonthForecastResult?.resultText || '')
const errorText = ref(store.state.nextMonthForecastResult?.errorText || '')
const props = defineProps({
  externalFile: { type: Object, default: null },
  generateKey: { type: Number, default: 0 },
  hideUploader: { type: Boolean, default: false }
})

watch([resultText, errorText], () => {
  store.commit('setNextMonthForecastResult', {
    resultText: resultText.value,
    errorText: errorText.value
  })
})

watch(
  () => props.generateKey,
  () => {
    if (!props.externalFile) return
    // 中文注释：统一用虚拟 change 事件复用原有解析逻辑。
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
        errorText.value = `未识别到以下子表：${missing.join('、')}`
        return
      }

      const newSign = analyzeSheet(
        workbook.Sheets[sheetMap['新签']],
        ['序号'],
        ['预计新签金额（元）*']
      )

      const revenue = analyzeSheet(
        workbook.Sheets[sheetMap['营收']],
        ['序号'],
        ['预计确认收入金额*（不含税）（元）']
      )

      const receipt = analyzeSheet(
        workbook.Sheets[sheetMap['回款']],
        ['序号'],
        ['预计回款金额*（元）']
      )

      const aa = newSign.maxSerial
      const bb = roundHalfUp(newSign.amountSum / 10000)
      const cc = revenue.maxSerial
      const dd = roundHalfUp(revenue.amountSum / 10000)
      const ee = receipt.maxSerial
      const ff = roundHalfUp(receipt.amountSum / 10000)

      // 中文注释：按要求仅输出最终句子，不附加其他内容。
      resultText.value = `预计实现新签项目${aa}个，合同额${bb}万元；预计实现营业收入项目${cc}个，金额${dd}万元；预计货款回笼项目${ee}个，金额${ff}万元。具体详情见附件三：《区域指挥部指标预计情况表》。`
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

function analyzeSheet(worksheet, serialCandidates, amountCandidates) {
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' })
  if (!rows.length) return { maxSerial: 0, amountSum: 0 }
  console.log('rows, serialCandidates, amountCandidates',rows, serialCandidates, amountCandidates);
  

  const headerInfo = findHeaderRow(rows, serialCandidates, amountCandidates)
  if (!headerInfo) {
    throw new Error('未识别到必需列标题（序号/金额）。')
  }

  let maxSerial = 0
  let amountSum = 0

  for (let r = headerInfo.headerRowIndex + 1; r < rows.length; r++) {
    const serialValue = parseNumber(getCellDisplayValue(worksheet, rows, r, headerInfo.serialColIndex))
    if (!Number.isNaN(serialValue)) {
      maxSerial = Math.max(maxSerial, roundHalfUp(serialValue))
    }

    const amountValue = parseNumber(getCellDisplayValue(worksheet, rows, r, headerInfo.amountColIndex))
    if (!Number.isNaN(amountValue)) {
      amountSum += amountValue
    }
  }

  return { maxSerial: roundHalfUp(maxSerial), amountSum }
}

function findHeaderRow(rows, serialCandidates, amountCandidates) {
  for (let i = 0; i < rows.length; i++) {
    const header = (rows[i] || []).map((cell) => normalizeHeaderText(cell))
    const serialColIndex = findHeaderIndex(header, serialCandidates)
    const amountColIndex = findHeaderIndex(header, amountCandidates)

    if (serialColIndex !== -1 && amountColIndex !== -1) {
      return { headerRowIndex: i, serialColIndex, amountColIndex }
    }
  }
  return null
}

function findHeaderIndex(header, candidates) {
  for (const candidate of candidates) {
    const idx = header.indexOf(normalizeHeaderText(candidate))
    if (idx !== -1) return idx
  }
  return -1
}

function getCellDisplayValue(worksheet, rows, rowIndex, colIndex) {
  if (colIndex < 0) return ''

  // 中文注释：优先读取格式化显示值，确保百分号等内容原样可读。
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

function parseNumber(value) {
  if (typeof value === 'number') return value
  const text = normalizeCellText(value)
  if (!text) return NaN

  const cleaned = text.replace(/,/g, '').replace(/[^0-9.\-]/g, '')
  const n = Number.parseFloat(cleaned)
  return Number.isNaN(n) ? NaN : n
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
.forecast-page {
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

</style>
