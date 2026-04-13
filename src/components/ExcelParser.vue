<template>
  <div>
    <label>上传 Excel 文件（.xlsx/.xls）：</label>
    <input type="file" accept=".xlsx,.xls" @change="onFileChange" />

    <div v-if="message" style="margin-top:12px;padding:12px;border:1px solid #ddd;background:#fafafa;">
      <strong>结果：</strong>
      <div style="margin-top:8px;">{{ message }}</div>
      <div style="margin-top:8px;color:#666;">解析明细（若某表格缺失或未识别列，会显示缺失说明）：</div>
      <ul>
        <li v-for="(d,k) in details" :key="k">{{ k }}：{{ d }}</li>
      </ul>
      <div style="margin-top:8px;">
        <!-- 导出为 TXT 按钮，仅当有结果时显示 -->
        <button @click="exportTxt">导出为 TXT</button>
      </div>
    </div>
  </div>
</template>

<script setup>
// 使用中文注释，说明非直观逻辑及关键步骤
import { ref, watch } from 'vue'
import { useStore } from 'vuex'
import * as XLSX from 'xlsx'

const store = useStore()
const message = ref(store.state.excelParserResult.message || '')
const details = ref(store.state.excelParserResult.details || {})

// 中文注释：统一监听页面结果并写入 Vuex，配合 store 持久化实现刷新不丢失。
watch(
  [message, details],
  () => {
    store.commit('setExcelParserResult', {
      message: message.value,
      details: details.value
    })
  },
  { deep: true }
)

// 读取文件并解析
function onFileChange(e) {
  const file = e.target.files && e.target.files[0]
  if (!file) return
  const reader = new FileReader()
  reader.onload = (ev) => {
    try {
      const data = new Uint8Array(ev.target.result)
      const workbook = XLSX.read(data, { type: 'array' })
      // 按需解析三张子表（优先按表名匹配）
      const newSign = parseSheet(workbook, '新签', '新签', '预计新签金额（元）')
      const revenue = parseSheet(workbook, '营收', '预计确认收入金额*（不含税）（元）')
      const receipt = parseSheet(workbook, '回款', '预计回款金额*（元）')

      // 将解析结果转换为所需数字（取最大序号作为项目数，金额求和并除以10000)
      const aa = roundInt(getCountFromResult(newSign))
      const bb = roundInt(getAmountWanFromResult(newSign))

      const cc = roundInt(getCountFromResult(revenue))
      const dd = roundInt(getAmountWanFromResult(revenue))

      const ee = roundInt(getCountFromResult(receipt))
      const ff = roundInt(getAmountWanFromResult(receipt))

      // 生成最终句子，按要求保留整数（四舍五入）
      message.value = `预计实现新签项目${aa}个，合同额${bb}万元；预计实现营业收入项目${cc}个，金额${dd}万元；预计货款回笼项目${ee}个，金额${ff}万元。具体详情见附件三：《区域指挥部指标预计情况表》。`

      // 填充解析明细用于调试与展示
      details.value = {
        '新签（原始）': summarizeForDisplay(newSign),
        '营收（原始）': summarizeForDisplay(revenue),
        '回款（原始）': summarizeForDisplay(receipt)
      }
    } catch (err) {
      message.value = '解析失败：' + err.message
    }
  }
  reader.readAsArrayBuffer(file)
}

// 导出当前结果为纯文本文件（TXT）
function exportTxt() {
  // 组成导出文本：先为结果句子，再附带解析明细，方便核对
  const lines = []
  if (message.value) lines.push(message.value)
  lines.push('\n解析明细：')
  const det = details.value || {}
  for (const k of Object.keys(det)) {
    lines.push(`${k}：${det[k]}`)
  }

  const content = lines.join('\n')
  // 创建 Blob 并触发浏览器下载
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  // 文件名使用简短描述
  a.href = url
  a.download = '区域指挥部指标预计情况表.txt'
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

// 将解析结果转为便于展示的摘要
function summarizeForDisplay(res) {
  if (!res) return '未找到表格或无法解析'
  return `序号最大值=${res.maxIndex ?? '-'}，金额合计（元）=${res.sumAmount ?? '-'}，行数=${res.rowCount ?? 0}`
}

// 从解析结果取项目数（取序号最大值）
function getCountFromResult(res) {
  if (!res || res.maxIndex == null || isNaN(res.maxIndex)) return 0
  return res.maxIndex
}

// 从解析结果取金额（元）并转换为万元
function getAmountWanFromResult(res) {
  if (!res || res.sumAmount == null || isNaN(res.sumAmount)) return 0
  return res.sumAmount / 10000
}

// 四舍五入并取整数
function roundInt(v) {
  return Math.round(Number(v) || 0)
}

// 解析单个工作表，支持按表名查找并智能识别“序号”和“金额”列
function parseSheet(workbook, sheetName, preferredAmountHeader) {
  const ws = workbook.Sheets[sheetName]
  if (!ws) return null
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 })
  if (!rows || rows.length === 0) return null

  // 假定第一行是表头
  const header = rows[0].map(h => (h == null ? '' : String(h).trim()))

  // 找到“序号”列索引
  const seqIdx = header.findIndex(h => h === '序号')

  // 尝试先用首选金额列名匹配（比如“预计新签金额（元）”），否则匹配任意包含“金额”的列
  let amtIdx = -1
  if (preferredAmountHeader) {
    amtIdx = header.findIndex(h => h === preferredAmountHeader)
  }
  if (amtIdx === -1) {
    amtIdx = header.findIndex(h => /金额/.test(h))
  }

  // 遍历数据行，收集序号和金额
  const seqs = []
  let sumAmount = 0
  let rowCount = 0
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r]
    if (!row || row.length === 0) continue
    rowCount++
    // 序号列解析为数值（可能为空或非数字）
    if (seqIdx >= 0) {
      const v = row[seqIdx]
      const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''))
      if (!isNaN(n)) seqs.push(n)
    }
    // 金额列解析为数值并累加（忽略非数字）
    if (amtIdx >= 0) {
      const av = row[amtIdx]
      const an = parseFloat(String(av).replace(/[^0-9.\-\.]/g, ''))
      if (!isNaN(an)) sumAmount += an
    }
  }

  const maxIndex = seqs.length ? Math.max(...seqs.map(x => Math.round(x))) : null
  return { maxIndex, sumAmount, rowCount }
}
</script>

<style scoped>
input[type="file"] { margin-top:8px }
</style>
