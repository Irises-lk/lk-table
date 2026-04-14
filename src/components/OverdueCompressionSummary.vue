<template>
  <section class="summary-page">
    <h3>逾期压降汇总</h3>

    <!-- <div class="actions">
      <button @click="buildSummary">获取并计算</button>
    </div> -->

    <div v-if="resultText" class="result-panel">{{ resultText }}</div>
  </section>
</template>

<script setup>
// 中文注释：本页面仅读取 Vuex 中三个页面的现有结果，不重复解析 Excel。
import { ref, onMounted, watch } from 'vue'
import { useStore } from 'vuex'

const DEFAULT_INITIAL_OVERDUE = 8136
const props = defineProps({
  autoBuildKey: { type: Number, default: 0 }
})

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

watch(
  () => props.autoBuildKey,
  () => {
    if (!props.autoBuildKey) return
    // 中文注释：统一生成后延迟轮询，确保依赖页面先完成异步解析再汇总。
    tryAutoBuild(0)
  }
)

function tryAutoBuild(retryCount) {
  const values = collectSourceValues(false)
  if (values) {
    resultText.value = composeText(values)
    persistResult()
    return
  }

  if (retryCount >= 100) return
  window.setTimeout(() => {
    tryAutoBuild(retryCount + 1)
  }, 120)
}

function collectSourceValues(showAlertWhenMissing) {
  const completedText = store.state.overdueReceiptMonthlyResult?.resultText || ''
  const stockText = store.state.eastRegionOverdueStockResult?.resultText || ''
  const newText = store.state.eastRegionNewOverdueResult?.resultText || ''
  const stockRawAmount = store.state.eastRegionOverdueStockResult?.rawAmount
  const newRawAmount = store.state.eastRegionNewOverdueResult?.rawAmount
  const initialOverdueRaw = store.state.eastRegionOverdueStockResult?.initialOverdue

  const cc = extractFirstNumber(completedText)
  // 中文注释：优先使用结构化数值，避免从文本抽数引入精度或匹配误差。
  const stock = Number.isFinite(stockRawAmount) ? Number(stockRawAmount) : extractAmountValue(stockText)
  const added = Number.isFinite(newRawAmount) ? Number(newRawAmount) : extractAmountValue(newText)
  const initialOverdue = Number.isFinite(initialOverdueRaw) ? Number(initialOverdueRaw) : DEFAULT_INITIAL_OVERDUE

  const missingPages = []
  if (Number.isNaN(cc)) missingPages.push('完成逾期')
  if (Number.isNaN(stock)) missingPages.push('存量逾期')
  if (Number.isNaN(added)) missingPages.push('新增逾期')
  if (Number.isNaN(initialOverdue) || initialOverdue === 0) missingPages.push('期初逾期')

  if (missingPages.length) {
    if (showAlertWhenMissing) {
      // 中文注释：按需求，任一页面结果无法获取时弹窗提示。
      window.alert(`无法获取以下页面结果：${missingPages.join('、')}。请先在对应页面完成计算。`)
    }
    return null
  }

  return { cc, stock, added, initialOverdue }
}

function composeText(values) {
  // 中文注释：百分比应基于原始小数值计算，避免先取整导致精度丢失。
  const currentOverdueRaw = values.stock + values.added
  const initialOverdue = values.initialOverdue
  console.log('values',values);
  
console.log('initialOverdue',initialOverdue);

  const aa = roundHalfUp(currentOverdueRaw)
  const bb = roundHalfUp(initialOverdue - currentOverdueRaw)

  const rate = (initialOverdue - currentOverdueRaw) / initialOverdue
  const percentText = `${roundHalfUp(Math.abs(rate * 100), 2).toFixed(2)}%`

  const descBb = bb >= 0 ? `较期初下降${Math.abs(bb)}万元` : `较期初增长${Math.abs(bb)}万元`
  const descXx = rate >= 0 ? `下降${percentText}` : `增长${percentText}`
  const cc = roundHalfUp(values.cc)
  const initialDisplay = roundHalfUp(initialOverdue)

  return `（四）逾期压降：期初逾期${initialDisplay}万元，当前逾期${aa}万元, ${descBb}, ${descXx}，本月完成逾期回款${cc}万元。`
}

function extractFirstNumber(text) {
  const raw = String(text == null ? '' : text)
  const match = raw.match(/-?\d+(?:\.\d+)?/)
  if (!match) return NaN
  const num = Number.parseFloat(match[0])
  return Number.isNaN(num) ? NaN : num
}

function extractAmountValue(text) {
  const raw = String(text == null ? '' : text)
  // 中文注释：优先匹配“为 xxx 万元”这种主值片段，避免匹配到句子中的其他数字。
  const amountMatch = raw.match(/为\s*(-?\d+(?:\.\d+)?)\s*万元/)
  if (amountMatch) {
    const amount = Number.parseFloat(amountMatch[1])
    if (!Number.isNaN(amount)) return amount
  }

  return extractFirstNumber(raw)
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
</style>
