<template>
  <section class="summary-page">
    <h3>逾期压降汇总</h3>

    <div class="actions">
      <button @click="buildSummary">获取并计算</button>
    </div>

    <div v-if="resultText" class="result-panel">{{ resultText }}</div>
  </section>
</template>

<script setup>
// 中文注释：本页面仅读取 Vuex 中三个页面的现有结果，不重复解析 Excel。
import { ref, onMounted } from 'vue'
import { useStore } from 'vuex'

const INITIAL_OVERDUE = 8136

const store = useStore()
const resultText = ref(store.state.overdueCompressionSummaryResult?.resultText || '')

onMounted(() => {
  // 中文注释：进入页面时若三项数据已齐全，自动计算并回显。
  const preview = collectSourceValues(false)
  if (preview) {
    resultText.value = composeText(preview)
    persistResult()
  }
})

function buildSummary() {
  const values = collectSourceValues(true)
  if (!values) return

  resultText.value = composeText(values)
  persistResult()
}

function collectSourceValues(showAlertWhenMissing) {
  const completedText = store.state.overdueReceiptMonthlyResult?.resultText || ''
  const stockText = store.state.eastRegionOverdueStockResult?.resultText || ''
  const newText = store.state.eastRegionNewOverdueResult?.resultText || ''

  const cc = extractFirstNumber(completedText)
  const stock = extractFirstNumber(stockText)
  const added = extractFirstNumber(newText)

  const missingPages = []
  if (Number.isNaN(cc)) missingPages.push('完成逾期')
  if (Number.isNaN(stock)) missingPages.push('存量逾期')
  if (Number.isNaN(added)) missingPages.push('新增逾期')

  if (missingPages.length) {
    if (showAlertWhenMissing) {
      // 中文注释：按需求，任一页面结果无法获取时弹窗提示。
      window.alert(`无法获取以下页面结果：${missingPages.join('、')}。请先在对应页面完成计算。`)
    }
    return null
  }

  return { cc, stock, added }
}

function composeText(values) {
  const aa = roundHalfUp(values.stock + values.added)
  const bb = roundHalfUp(INITIAL_OVERDUE - aa)
  const rate = (INITIAL_OVERDUE - aa) / INITIAL_OVERDUE
  const percentText = `${roundHalfUp(Math.abs(rate * 100), 2).toFixed(2)}%`

  const descBb = bb >= 0 ? `较期初下降${Math.abs(bb)}万元` : `较期初增长${Math.abs(bb)}万元`
  const descXx = rate >= 0 ? `下降${percentText}` : `增长${percentText}`
  const cc = roundHalfUp(values.cc)

  return `（四）逾期压降：期初逾期8136万元，当前逾期${aa}万元, ${descBb}, ${descXx}，本月完成逾期回款${cc}万元。`
}

function extractFirstNumber(text) {
  const raw = String(text == null ? '' : text)
  const match = raw.match(/-?\d+(?:\.\d+)?/)
  if (!match) return NaN
  const num = Number.parseFloat(match[0])
  return Number.isNaN(num) ? NaN : num
}

function roundHalfUp(value, decimals = 0) {
  const factor = 10 ** decimals
  const scaled = Number(value) * factor
  if (Number.isNaN(scaled)) return 0

  // 中文注释：负数按绝对值后再四舍五入，保证与常规“四舍五入”一致。
  const rounded = scaled >= 0 ? Math.round(scaled) : -Math.round(Math.abs(scaled))
  return rounded / factor
}

function persistResult() {
  store.commit('setOverdueCompressionSummaryResult', {
    resultText: resultText.value
  })
}
</script>

<style scoped>
.summary-page {
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.actions {
  display: flex;
  align-items: center;
  gap: 8px;
}

.result-panel {
  padding: 12px;
  border: 1px solid #ddd;
  background: #fafafa;
  line-height: 1.8;
}
</style>
