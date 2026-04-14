<template>
  <div class="app-shell">
    <h2>华东区域经营月报一键生成</h2>

    <section class="upload-panel">
      <div class="upload-row">
        <label for="allFiles">一次性上传全部源文件（支持多选）：</label>
        <input id="allFiles" type="file" accept=".xlsx,.xls" multiple @change="onAllFilesChange" />
      </div>

      <details class="rules-config">
        <summary>匹配规则配置（可编辑）</summary>
        <p class="rules-tip">中文注释：每行表示一组匹配规则；同一行使用 + 连接多个关键字表示“同时包含”。</p>
        <div class="rules-grid">
          <div v-for="item in ruleMeta" :key="item.key" class="rule-item">
            <label :for="`rule-${item.key}`">{{ item.title }}</label>
            <textarea
              :id="`rule-${item.key}`"
              v-model="editableRules[item.key]"
              rows="2"
              :placeholder="item.placeholder"
            />
          </div>
        </div>
        <div class="rules-actions">
          <button type="button" @click="resetRules">恢复默认规则</button>
        </div>
      </details>

      <button class="generate-btn" @click="generateAll">一键生成全部页面结果</button>

      <ul class="match-list">
        <li v-for="item in fileMatchView" :key="item.title">
          <span class="item-title">{{ item.title }}：</span>
          <span :class="item.matched ? 'ok' : 'warn'">{{ item.fileName }}</span>
        </li>
      </ul>
    </section>


    <section class="page-block">
      <LedgerReportGenerator :external-file="matchedFiles.ledger" :generate-key="generateKey" :hide-uploader="true" />
    </section>

    <section class="page-block">
      <OverdueReceiptMonthlyReport :external-file="matchedFiles.ledger" :generate-key="generateKey" :hide-uploader="true" />
    </section>

    <section class="page-block">
      <EastRegionOverdueStock :external-file="matchedFiles.stockOverdue" :generate-key="generateKey" :hide-uploader="true" />
    </section>

    <section class="page-block">
      <EastRegionNewOverdue :external-file="matchedFiles.newOverdue" :generate-key="generateKey" :hide-uploader="true" />
    </section>

    <section class="page-block">
      <MainInHandContractAmount :external-file="matchedFiles.inHand" :generate-key="generateKey" :hide-uploader="true" />
    </section>

    <section class="page-block">
      <ProjectTrackingSummary :external-file="matchedFiles.tracking" :generate-key="generateKey" :hide-uploader="true" />
    </section>

    <section class="page-block">
      <StageDynamicsSummary :external-file="matchedFiles.stageDynamics" :generate-key="generateKey" :hide-uploader="true" />
    </section>

    <section class="page-block">
      <BidImplementationSummary
        :ledger-external-file="matchedFiles.ledger"
        :in-hand-external-file="matchedFiles.inHand"
        :generate-key="generateKey"
        :hide-uploader="true"
      />
    </section>

    <section class="page-block">
      <OutboundProjectFollowup :external-file="matchedFiles.outbound" :generate-key="generateKey" :hide-uploader="true" />
    </section>

    <section class="page-block">
      <CompetitorAnalysisStrict :external-file="matchedFiles.competitor" :generate-key="generateKey" :hide-uploader="true" />
    </section>
    <section class="page-block">
      <NextMonthForecastSummary :external-file="matchedFiles.forecast" :generate-key="generateKey" :hide-uploader="true" />
    </section>

    <section class="page-block">
      <OverdueCompressionSummary :auto-build-key="generateKey" />
    </section>
  </div>
</template>

<script setup>
import { computed, ref, watch } from 'vue'
import NextMonthForecastSummary from './components/NextMonthForecastSummary.vue'
import LedgerReportGenerator from './components/LedgerReportGenerator.vue'
import OverdueReceiptMonthlyReport from './components/OverdueReceiptMonthlyReport.vue'
import EastRegionOverdueStock from './components/EastRegionOverdueStock.vue'
import EastRegionNewOverdue from './components/EastRegionNewOverdue.vue'
import MainInHandContractAmount from './components/MainInHandContractAmount.vue'
import ProjectTrackingSummary from './components/ProjectTrackingSummary.vue'
import StageDynamicsSummary from './components/StageDynamicsSummary.vue'
import BidImplementationSummary from './components/BidImplementationSummary.vue'
import OutboundProjectFollowup from './components/OutboundProjectFollowup.vue'
import CompetitorAnalysisStrict from './components/CompetitorAnalysisStrict.vue'
import OverdueCompressionSummary from './components/OverdueCompressionSummary.vue'

const selectedFiles = ref([])
const generateKey = ref(0)
const RULE_STORAGE_KEY = 'lkgzl_rule_editor_v1'

// 中文注释：规则元数据用于渲染可编辑配置区。
const ruleMeta = [
  { key: 'forecast', title: '次月指标预计汇总', placeholder: '月份预计指标表\n预计指标表' },
  { key: 'competitor', title: '竞争对手承揽统计（含严格版）', placeholder: '竞争对手经营承揽情况' },
  { key: 'ledger', title: '业绩台账（含完成逾期/中标实施）', placeholder: '业绩台账' },
  { key: 'stockOverdue', title: '存量逾期', placeholder: '存量逾期' },
  { key: 'newOverdue', title: '新增逾期', placeholder: '新增逾期' },
  { key: 'inHand', title: '在手主办合同总额（含中标实施）', placeholder: '在手合同台账\n在手合同' },
  { key: 'tracking', title: '信息跟踪阶段统计', placeholder: '项目信息更新核对' },
  { key: 'stageDynamics', title: '三阶段动态统计', placeholder: '重点项目进度信息表' },
  { key: 'outbound', title: '中资外带项目信息跟进', placeholder: '属地企业带出去项目信息跟进反馈表\n带出去项目信息跟进反馈表' }
]

const defaultEditableRules = {
  forecast: '月份预计指标表\n预计指标表',
  competitor: '竞争对手经营承揽情况',
  ledger: '业绩台账',
  stockOverdue: '存量逾期',
  newOverdue: '新增逾期',
  inHand: '在手合同台账\n在手合同',
  tracking: '项目信息更新核对',
  stageDynamics: '重点项目进度信息表',
  outbound: '属地企业带出去项目信息跟进反馈表\n带出去项目信息跟进反馈表'
}

const editableRules = ref(loadEditableRules())

const parsedRules = computed(() => {
  const output = {}
  for (const item of ruleMeta) {
    output[item.key] = parseRuleText(editableRules.value[item.key])
  }
  return output
})

watch(
  editableRules,
  () => {
    // 中文注释：规则编辑后写入本地存储，避免刷新丢失。
    localStorage.setItem(RULE_STORAGE_KEY, JSON.stringify(editableRules.value))
  },
  { deep: true }
)

const matchedFiles = computed(() => {
  return {
    forecast: pickByRules(selectedFiles.value, parsedRules.value.forecast),
    competitor: pickByRules(selectedFiles.value, parsedRules.value.competitor),
    ledger: pickByRules(selectedFiles.value, parsedRules.value.ledger),
    stockOverdue: pickByRules(selectedFiles.value, parsedRules.value.stockOverdue),
    newOverdue: pickByRules(selectedFiles.value, parsedRules.value.newOverdue),
    inHand: pickByRules(selectedFiles.value, parsedRules.value.inHand),
    tracking: pickByRules(selectedFiles.value, parsedRules.value.tracking),
    stageDynamics: pickByRules(selectedFiles.value, parsedRules.value.stageDynamics),
    outbound: pickByRules(selectedFiles.value, parsedRules.value.outbound)
  }
})

const fileMatchView = computed(() => {
  return [
    { title: '次月指标预计汇总', fileName: fileNameOrMissing(matchedFiles.value.forecast), matched: !!matchedFiles.value.forecast },
    { title: '竞争对手承揽统计（含严格版）', fileName: fileNameOrMissing(matchedFiles.value.competitor), matched: !!matchedFiles.value.competitor },
    { title: '业绩台账/完成逾期/中标实施阶段（业绩台账）', fileName: fileNameOrMissing(matchedFiles.value.ledger), matched: !!matchedFiles.value.ledger },
    { title: '存量逾期', fileName: fileNameOrMissing(matchedFiles.value.stockOverdue), matched: !!matchedFiles.value.stockOverdue },
    { title: '新增逾期', fileName: fileNameOrMissing(matchedFiles.value.newOverdue), matched: !!matchedFiles.value.newOverdue },
    { title: '在手主办合同总额/中标实施阶段（在手合同）', fileName: fileNameOrMissing(matchedFiles.value.inHand), matched: !!matchedFiles.value.inHand },
    { title: '信息跟踪阶段统计', fileName: fileNameOrMissing(matchedFiles.value.tracking), matched: !!matchedFiles.value.tracking },
    { title: '三阶段动态统计', fileName: fileNameOrMissing(matchedFiles.value.stageDynamics), matched: !!matchedFiles.value.stageDynamics },
    { title: '中资外带项目信息跟进', fileName: fileNameOrMissing(matchedFiles.value.outbound), matched: !!matchedFiles.value.outbound }
  ]
})

function onAllFilesChange(event) {
  const files = Array.from((event.target && event.target.files) || [])
  selectedFiles.value = files
}

function generateAll() {
  // 中文注释：通过递增 key 触发各子组件的 watch，从而统一开始解析。
  generateKey.value += 1
}

function fileNameOrMissing(file) {
  return file ? file.name : '未匹配到文件'
}

function pickByRules(files, rules) {
  if (!Array.isArray(files) || !files.length) return null

  for (const file of files) {
    const normalized = normalizeText(file.name)
    for (const keywords of rules) {
      const hit = keywords.every((keyword) => normalized.includes(normalizeText(keyword)))
      if (hit) return file
    }
  }
  return null
}

function normalizeText(text) {
  return String(text == null ? '' : text).replace(/\s+/g, '').replace(/[()（）\[\]【】_.-]/g, '').toLowerCase()
}

function parseRuleText(text) {
  const lines = String(text == null ? '' : text)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line)

  // 中文注释：每行支持“关键词A+关键词B”语法，表示文件名需同时包含多个关键字。
  return lines.map((line) => {
    return line
      .split('+')
      .map((part) => part.trim())
      .filter((part) => part)
  })
}

function loadEditableRules() {
  try {
    const raw = localStorage.getItem(RULE_STORAGE_KEY)
    if (!raw) return { ...defaultEditableRules }
    const parsed = JSON.parse(raw)
    return {
      ...defaultEditableRules,
      ...(parsed && typeof parsed === 'object' ? parsed : {})
    }
  } catch (error) {
    return { ...defaultEditableRules }
  }
}

function resetRules() {
  editableRules.value = { ...defaultEditableRules }
}
</script>

<style scoped>
.app-shell {
  max-width: 1080px;
  margin: 24px auto;
  padding: 0 12px 20px;
  font-family: 'Microsoft YaHei', 'PingFang SC', sans-serif;
  background: linear-gradient(180deg, #f6faf7 0%, #f8f9fc 100%);
}

.app-shell h2 {
  margin: 0 0 14px;
  padding: 10px 14px;
  border-radius: 10px;
  background: linear-gradient(135deg, #166534 0%, #1f9d65 100%);
  color: #fff;
  letter-spacing: 0.5px;
}

.upload-panel {
  padding: 14px;
  border: 1px solid #d8e7de;
  border-radius: 12px;
  background: #ffffff;
  box-shadow: 0 6px 18px rgba(17, 56, 36, 0.06);
  margin-bottom: 18px;
}

.upload-row {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
}

.rules-config {
  margin-top: 10px;
  border: 1px solid #e6e6e6;
  background: #fff;
  padding: 8px;
}

.rules-config summary {
  cursor: pointer;
  font-weight: 600;
}

.rules-tip {
  margin: 8px 0;
  color: #666;
  font-size: 13px;
}

.rules-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
  gap: 8px;
}

.rule-item {
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.rule-item textarea {
  width: 100%;
  resize: vertical;
  min-height: 52px;
  padding: 6px;
}

.rules-actions {
  margin-top: 8px;
}

.generate-btn {
  margin-top: 10px;
  min-width: 210px;
  padding: 8px 16px;
  border: 0;
  border-radius: 8px;
  color: #fff;
  background: linear-gradient(135deg, #0f766e 0%, #0ea5a0 100%);
  cursor: pointer;
}

.generate-btn:hover {
  filter: brightness(1.05);
}

.match-list {
  margin-top: 10px;
  padding-left: 18px;
}

.item-title {
  color: #333;
}

.ok {
  color: #146b2e;
}

.warn {
  color: #a41515;
}

.page-block {
  margin-bottom: 14px;
  padding: 12px;
  border: 1px solid #e2e8f0;
  border-radius: 12px;
  background: #fff;
  box-shadow: 0 4px 14px rgba(0, 0, 0, 0.04);
}

/* 中文注释：使用深度选择器统一子组件页面视觉，减少逐个组件维护成本。 */
.page-block :deep(h3) {
  margin: 0 0 10px;
  padding-left: 10px;
  border-left: 4px solid #0ea5a0;
  font-size: 18px;
  color: #0f172a;
}

.page-block :deep(.desc) {
  margin: 0 0 10px;
  color: #4b5563;
}

.page-block :deep(.upload-row),
.page-block :deep(.actions) {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
  margin-bottom: 10px;
}

.page-block :deep(input[type='file']) {
  padding: 6px;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  background: #f8fafc;
}

.page-block :deep(button) {
  padding: 6px 12px;
  border: 1px solid #cbd5e1;
  border-radius: 8px;
  background: #f8fafc;
  cursor: pointer;
}

.page-block :deep(button:hover) {
  background: #eef2ff;
}

.page-block :deep(.panel),
.page-block :deep(.result-panel),
.page-block :deep(.result-text) {
  padding: 12px;
  border: 1px solid #cce6dc;
  border-radius: 10px;
  background: #f3fbf7;
  color: #14532d;
  line-height: 1.8;
}

.page-block :deep(.error-panel),
.page-block :deep(.error-text) {
  padding: 12px;
  border: 1px solid #fecaca;
  border-radius: 10px;
  background: #fef2f2;
  color: #991b1b;
  line-height: 1.8;
}

.page-block :deep(.content),
.page-block :deep(.text-block),
.page-block :deep(.result-only) {
  margin-top: 6px;
  white-space: pre-wrap;
}
</style>