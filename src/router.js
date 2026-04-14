import { createRouter, createWebHistory } from 'vue-router'
import ExcelParser from './components/ExcelParser.vue'
import CompetitorAnalysis from './components/CompetitorAnalysis.vue'
import LedgerReportGenerator from './components/LedgerReportGenerator.vue'
import OverdueReceiptMonthlyReport from './components/OverdueReceiptMonthlyReport.vue'
import EastRegionOverdueStock from './components/EastRegionOverdueStock.vue'
import EastRegionNewOverdue from './components/EastRegionNewOverdue.vue'
import OverdueCompressionSummary from './components/OverdueCompressionSummary.vue'
import MainInHandContractAmount from './components/MainInHandContractAmount.vue'
import ProjectTrackingSummary from './components/ProjectTrackingSummary.vue'
import StageDynamicsSummary from './components/StageDynamicsSummary.vue'
import BidImplementationSummary from './components/BidImplementationSummary.vue'
import OutboundProjectFollowup from './components/OutboundProjectFollowup.vue'
import CompetitorAnalysisStrict from './components/CompetitorAnalysisStrict.vue'
import NextMonthForecastSummary from './components/NextMonthForecastSummary.vue'
// 中文注释：新增路由用于竞争对手承揽统计与分析页面

const routes = [
  {
    path: '/',
    name: 'Home',
    component: ExcelParser
  }
  ,
  {
    path: '/competitor',
    name: 'CompetitorAnalysis',
    component: CompetitorAnalysis
  },
  {
    path: '/ledger-report',
    name: 'LedgerReportGenerator',
    component: LedgerReportGenerator
  },
  {
    path: '/overdue-receipt-monthly-report',
    name: 'OverdueReceiptMonthlyReport',
    component: OverdueReceiptMonthlyReport
  },
  {
    path: '/east-region-overdue-stock',
    name: 'EastRegionOverdueStock',
    component: EastRegionOverdueStock
  },
  {
    path: '/east-region-new-overdue',
    name: 'EastRegionNewOverdue',
    component: EastRegionNewOverdue
  },
  {
    path: '/overdue-compression-summary',
    name: 'OverdueCompressionSummary',
    component: OverdueCompressionSummary
  },
  {
    path: '/main-in-hand-contract-amount',
    name: 'MainInHandContractAmount',
    component: MainInHandContractAmount
  },
  {
    path: '/project-tracking-summary',
    name: 'ProjectTrackingSummary',
    component: ProjectTrackingSummary
  },
  {
    path: '/stage-dynamics-summary',
    name: 'StageDynamicsSummary',
    component: StageDynamicsSummary
  },
  {
    path: '/bid-implementation-summary',
    name: 'BidImplementationSummary',
    component: BidImplementationSummary
  },
  {
    path: '/outbound-project-followup',
    name: 'OutboundProjectFollowup',
    component: OutboundProjectFollowup
  },
  {
    path: '/competitor-analysis-strict',
    name: 'CompetitorAnalysisStrict',
    component: CompetitorAnalysisStrict
  },
  {
    path: '/next-month-forecast-summary',
    name: 'NextMonthForecastSummary',
    component: NextMonthForecastSummary
  }
]

const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes
})

export default router