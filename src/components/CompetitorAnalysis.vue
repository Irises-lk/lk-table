<template>
  <div>
    <h3>竞争对手承揽统计与分析</h3>
    <template v-if="!hideUploader">
      <label>上传 Excel 文件（含子表：掘进机、特装装备）：</label>
      <input type="file" accept=".xlsx,.xls" @change="onFileChange" />
    </template>

    <div v-if="resultText" style="margin-top:12px;padding:12px;border:1px solid #ddd;background:#fafafa;">
      <strong>输出文案：</strong>
      <div style="margin-top:8px;white-space:pre-wrap;">{{ resultText }}</div>
      <div style="margin-top:8px;">
        <button @click="exportTxt">导出为 TXT</button>
      </div>
      <div style="margin-top:12px;color:#666">解析明细：</div>
      <ul>
        <li v-for="(v,k) in details" :key="k">{{ k }}：{{ v }}</li>
      </ul>
    </div>
  </div>
</template>

<script setup>
// 中文注释：本文件实现上传 Excel、模糊匹配列名、按要求统计并输出文案
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const store = useStore()
const resultText = ref(store.state.competitorAnalysisResult.resultText || '')
const details = ref(store.state.competitorAnalysisResult.details || {})
const props = defineProps({
  externalFile: { type: Object, default: null },
  generateKey: { type: Number, default: 0 },
  hideUploader: { type: Boolean, default: false }
})

// 中文注释：计算结果变化时同步到 Vuex，确保跨路由与刷新后可回显。
watch(
  [resultText, details],
  () => {
    store.commit('setCompetitorAnalysisResult', {
      resultText: resultText.value,
      details: details.value
    })
  },
  { deep: true }
)

watch(
  () => props.generateKey,
  () => {
    if (!props.externalFile) return
    // 中文注释：复用既有 onFileChange，避免改动核心统计逻辑。
    onFileChange({ target: { files: [props.externalFile] } })
  }
)

// 触发点为当前日期，统计上一个月
function getPrevMonthYear() {
  const now = new Date()
  const y = now.getFullYear()
  const m = now.getMonth() + 1 // 1-12
  if (m === 1) return { year: y - 1, month: 12 }
  return { year: y, month: m - 1 }
}

// 解析上传文件
function onFileChange(e) {
  const file = e.target.files && e.target.files[0]
  if (!file) return
  const reader = new FileReader()
  reader.onload = (ev) => {
    try {
      const data = new Uint8Array(ev.target.result)
      const workbook = XLSX.read(data, { type: 'array' })

      // 处理两个子表：掘进机 与 特装装备（模糊匹配表名）
      const sheetJ = findSheet(workbook, '掘进机')
      const sheetT = findSheet(workbook, '特装装备')

      const prev = getPrevMonthYear()

      const resJ = sheetJ ? analyzeSheet(sheetJ, prev) : emptyResult()
      const resT = sheetT ? analyzeSheet(sheetT, prev) : emptyResult()

      // 计算最终字母变量并取整数（不保留小数）
      const aa = int(resJ.totalAmountWan)
      const gg = int(resJ.countPrevMonth)
      const jj = int(resJ.totalQuantity)
      const bb = int(resJ.tjzj_amountWan)
      const kk = int(resJ.tjzj_quantity)
      const cc = int(Math.max(0, aa - bb))
      const ll = int(Math.max(0, jj - kk))

      const dd = int(resT.totalAmountWan)
      const hh = int(resT.countPrevMonth)
      const mm = int(resT.totalQuantity)
      const ee = int(resT.tjzj_amountWan)
      const nn = int(resT.tjzj_quantity)
      const ff = int(Math.max(0, dd - ee))
      const oo = int(Math.max(0, mm - nn))

      const xx = int(gg + hh)
      const ii = String(prev.month) // 上一个月的月份数字

      const year = new Date().getFullYear()

      // 形成指定输出段落（中文，数字为整数）
      resultText.value = `本月新增${xx}项竞争对手承揽统计与分析内容。${year}年1-${ii}月，掘进机产品，我司承揽${kk}台，金额${bb}万元；竞争对手承揽${ll}台，金额${cc}万元。特种装备产品，我司承揽${nn}台，金额${ee}万元；竞争对手承揽${oo}台，金额${ff}万元。具体详情见附件：《竞争对手承揽统计与分析表》。（其中：xx即为gg加上hh，ii月为触发该节点的上一个月，例如3月份触发该节点，ii月即为2月,其他字母则为前两步计算的数字，且所有数字均不保留小数）`

      // 填充详情便于核验
      details.value = {
        '掘进机-原始解析(行数/金额(万元)/台数/上月新增项目数)': `${resJ.rowCount}/${round(resJ.totalAmountWan)}/${round(resJ.totalQuantity)}/${resJ.countPrevMonth}`,
        '掘进机-我司(铁建重工)金额(万元)/台数': `${round(resJ.tjzj_amountWan)}/${round(resJ.tjzj_quantity)}`,
        '特装装备-原始解析(行数/金额(万元)/台数/上月新增项目数)': `${resT.rowCount}/${round(resT.totalAmountWan)}/${round(resT.totalQuantity)}/${resT.countPrevMonth}`,
        '特装装备-我司(铁建重工)金额(万元)/台数': `${round(resT.tjzj_amountWan)}/${round(resT.tjzj_quantity)}`
      }
    } catch (err) {
      resultText.value = '解析失败：' + (err && err.message ? err.message : err)
    }
  }
  reader.readAsArrayBuffer(file)
}

function emptyResult() {
  return { totalAmountWan: 0, totalQuantity: 0, rowCount: 0, countPrevMonth: 0, tjzj_amountWan: 0, tjzj_quantity: 0 }
}

// 将值四舍五入为整数（展示用）
function round(v) { return Math.round(Number(v) || 0) }
function int(v) { return Math.round(Number(v) || 0) }

// 在 workbook 中按包含匹配查找表格名
function findSheet(workbook, keyword) {
  const names = workbook.SheetNames || []
  for (const n of names) {
    if (String(n).includes(keyword)) return workbook.Sheets[n]
  }
  // 若未找到精确包含，则尝试模糊：包含关键字的任何近似
  for (const n of names) {
    if (String(n).replace(/\s+/g, '').indexOf(keyword.replace(/\s+/g, '')) >= 0) return workbook.Sheets[n]
  }
  return null
}

// 分析单个工作表，返回所需统计量
function analyzeSheet(ws, prevMonthYear) {
  // 使用 defval 防止短行导致 undefined
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
  if (!rows || rows.length === 0) return emptyResult()
  // 去除表头中的换行并去除首尾空白，确保换行不会影响精确匹配
  const headerRow = rows[0].map(h => (h == null ? '' : String(h).replace(/\r?\n/g, '').trim()))

  // 使用模糊匹配查找列索引（返回数字索引或 -1）
  const idxAmount = findHeaderIndex(headerRow, 'amount')
  const idxTime = findHeaderIndex(headerRow, 'time')
  const idxQty = findHeaderIndex(headerRow, 'qty')
  const idxUnit = findHeaderIndex(headerRow, 'unit')

  let totalAmountWan = 0
  let totalQuantity = 0
  let rowCount = 0
  let countPrevMonth = 0

  let tjzj_amountWan = 0 // 铁建重工金额（万元）
  let tjzj_quantity = 0

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r]
    if (!row || row.length === 0) continue
    rowCount++

    // 安全访问单元格，避免短行或索引越界
    const rawAmount = (idxAmount >= 0 && row.length > idxAmount) ? row[idxAmount] : ''
    const rawQty = (idxQty >= 0 && row.length > idxQty) ? row[idxQty] : ''
    const rawTime = (idxTime >= 0 && row.length > idxTime) ? row[idxTime] : ''
    const rawUnit = (idxUnit >= 0 && row.length > idxUnit) ? row[idxUnit] : ''

    const amountWan = parseAmountWan(rawAmount, idxAmount >= 0 ? headerRow[idxAmount] : '')
    const qty = parseNumber(rawQty)
    const unitName = rawUnit == null ? '' : String(rawUnit).trim()

    totalAmountWan += amountWan
    totalQuantity += (isNaN(qty) ? 0 : qty)

    // 判断是否为上个月
    if (rawTime != null && rawTime !== '') {
      const dt = parseDateCell(rawTime)
      if (dt && dt.getFullYear() === prevMonthYear.year && (dt.getMonth() + 1) === prevMonthYear.month) {
        countPrevMonth++
      }
    }

    // 判断是否为铁建重工（模糊匹配）
    if (unitName && /铁建重工/.test(unitName)) {
      tjzj_amountWan += amountWan
      tjzj_quantity += (isNaN(qty) ? 0 : qty)
    }
  }

  return { totalAmountWan, totalQuantity, rowCount, countPrevMonth, tjzj_amountWan, tjzj_quantity }
}

// 解析金额并返回单位为万元的数字（尝试根据表头关键词判断单位）
function parseAmountWan(raw, headerText) {
  const v = parseNumber(raw)
  if (isNaN(v)) return 0
  const h = headerText ? String(headerText) : ''
  const low = h.toLowerCase()
  if (/万元/.test(h) || /万元/.test(low)) return v
  if (/元/.test(h) && !/万元/.test(h)) return v / 10000
  // 否则按启发式：若值很大（>1e6），视作元；否则视作万元
  if (Math.abs(v) > 1000000) return v / 10000
  return v
}

// 简单数值清洗
function parseNumber(raw) {
  if (raw == null) return NaN
  if (typeof raw === 'number') return raw
  const s = String(raw).replace(/[,\s]/g, '').replace(/[^0-9.\-]/g, '')
  if (s === '') return NaN
  const n = parseFloat(s)
  return isNaN(n) ? NaN : n
}

// 解析单元格为 Date（支持字符串、Date 对象、Excel 序列号）
function parseDateCell(v) {
  if (v == null) return null
  if (v instanceof Date) return v
  if (typeof v === 'number') {
    // Excel 内部序列号转日期（参考 SheetJS）
    try {
      const d = XLSX.SSF.parse_date_code(v)
      if (d) return new Date(d.y, d.m - 1, d.d)
    } catch (e) {
      return null
    }
  }
  // 字符串尝试解析
  const s = String(v).trim()
  const t = Date.parse(s)
  if (!isNaN(t)) return new Date(t)
  // 尝试常见 yyyy/mm/dd 或 yyyy-mm-dd 带时分秒的情况
  const m = s.match(/(\d{4})[./-](\d{1,2})[./-](\d{1,2})/)
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]))
  return null
}

// 根据意图类型模糊匹配表头索引（返回列的数字索引或 -1）
function findHeaderIndex(headers, type) {
  // 先构建去换行的规范化头名，用于精确匹配
  const norms = headers.map(h => (h == null ? '' : String(h).replace(/\r?\n/g, '').trim()))

  // 各类型的候选精确表头（考虑括号、中文/英文括号等常见变体）
  const amountNames = ['承揽合同额（万元）', '承揽合同额(万元)', '承揽合同额（元）', '承揽合同额(元)', '承揽合同额']
  const timeNames = ['承揽确定时间（中标、竞谈或合同签订）', '承揽确定时间(中标、竞谈或合同签订)', '承揽确定时间', '承揽时间']
  const qtyNames = ['台/套数', '台/套', '台数', '套数', '台/套数']
  const unitNames = ['承揽单位', '单位', '承揽单位（单位）']

  // 优先做去换行后的精确匹配
  for (let i = 0; i < norms.length; i++) {
    const h = norms[i]
    if (type === 'amount' && amountNames.includes(h)) return i
    if (type === 'time' && timeNames.includes(h)) return i
    if (type === 'qty' && qtyNames.includes(h)) return i
    if (type === 'unit' && unitNames.includes(h)) return i
  }

  // 若精确匹配未命中，则退回到模糊匹配以提高容错性
  for (let i = 0; i < norms.length; i++) {
    const h = norms[i]
    if (type === 'amount') {
      if (/承揽/.test(h) && /(合同|金额|额)/.test(h)) return i
      if (/(金额|合同|承揽)/.test(h) && /(万元|元|金额|合同|额)/.test(h)) return i
      if (/金额/.test(h)) return i
    }
    if (type === 'time') {
      if (/承揽.*确定|确定.*时间|中标|竞谈|合同签订|时间/.test(h)) return i
      if (/时间/.test(h)) return i
    }
    if (type === 'qty') {
      if (/(台|套)/.test(h)) return i
      if (/数量/.test(h)) return i
    }
    if (type === 'unit') {
      if (/承揽单位|单位|承揽人/.test(h)) return i
      if (/单位/.test(h)) return i
    }
  }

  return -1
}

// 导出为文本文件
function exportTxt() {
  const txt = resultText.value || ''
  const det = JSON.stringify(details.value || {}, null, 2)
  const content = txt + '\n\n解析明细：\n' + det
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = '竞争对手承揽统计与分析.txt'
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}
</script>

<style scoped>
input[type="file"] { margin-top:8px }
</style>
