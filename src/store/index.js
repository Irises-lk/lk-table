import { createStore } from 'vuex'

const STORAGE_KEY = 'lkgzl_report_results_v1'

// 中文注释：从 localStorage 读取历史结果，保证刷新后仍能回显。
function loadPersistedState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return {}
    const parsed = JSON.parse(raw)
    return parsed && typeof parsed === 'object' ? parsed : {}
  } catch (error) {
    return {}
  }
}

const persisted = loadPersistedState()

const store = createStore({
  state() {
    return {
      excelParserResult: persisted.excelParserResult || { message: '', details: {} },
      competitorAnalysisResult: persisted.competitorAnalysisResult || { resultText: '', details: {} },
      ledgerReportResult: persisted.ledgerReportResult || { statusText: '', sheetRecognition: {}, rowData: null, resultText: '' },
      overdueReceiptMonthlyResult: persisted.overdueReceiptMonthlyResult || { statusText: '', sheetRecognition: {}, resultText: '' },
      eastRegionOverdueStockResult: persisted.eastRegionOverdueStockResult || { resultText: '', errorText: '' },
      eastRegionNewOverdueResult: persisted.eastRegionNewOverdueResult || { resultText: '', errorText: '' },
      overdueCompressionSummaryResult: persisted.overdueCompressionSummaryResult || { resultText: '' },
      mainInHandContractAmountResult: persisted.mainInHandContractAmountResult || { resultText: '', errorText: '' },
      projectTrackingSummaryResult: persisted.projectTrackingSummaryResult || { resultText: '', errorText: '' },
      stageDynamicsSummaryResult: persisted.stageDynamicsSummaryResult || { resultText: '', errorText: '' },
      bidImplementationSummaryResult: persisted.bidImplementationSummaryResult || { resultText: '', errorText: '' },
      outboundProjectFollowupResult: persisted.outboundProjectFollowupResult || { resultText: '', errorText: '' },
      competitorAnalysisStrictResult: persisted.competitorAnalysisStrictResult || { resultText: '', errorText: '' },
      nextMonthForecastResult: persisted.nextMonthForecastResult || { resultText: '', errorText: '' },
      inHandThreeSheetTableResult: persisted.inHandThreeSheetTableResult || {
        statusText: '',
        errorText: '',
        data: { targetMonth: '', newSignRows: [], revenueRows: [], receiptRows: [] }
      }
    }
  },
  mutations: {
    setExcelParserResult(state, payload) {
      state.excelParserResult = {
        message: payload && payload.message ? payload.message : '',
        details: payload && payload.details ? payload.details : {}
      }
    },
    setCompetitorAnalysisResult(state, payload) {
      state.competitorAnalysisResult = {
        resultText: payload && payload.resultText ? payload.resultText : '',
        details: payload && payload.details ? payload.details : {}
      }
    },
    setLedgerReportResult(state, payload) {
      state.ledgerReportResult = {
        statusText: payload && payload.statusText ? payload.statusText : '',
        sheetRecognition: payload && payload.sheetRecognition ? payload.sheetRecognition : {},
        rowData: payload && payload.rowData ? payload.rowData : null,
        resultText: payload && payload.resultText ? payload.resultText : ''
      }
    },
    setOverdueReceiptMonthlyResult(state, payload) {
      state.overdueReceiptMonthlyResult = {
        statusText: payload && payload.statusText ? payload.statusText : '',
        sheetRecognition: payload && payload.sheetRecognition ? payload.sheetRecognition : {},
        resultText: payload && payload.resultText ? payload.resultText : ''
      }
    },
    setEastRegionOverdueStockResult(state, payload) {
      state.eastRegionOverdueStockResult = {
        resultText: payload && payload.resultText ? payload.resultText : '',
        errorText: payload && payload.errorText ? payload.errorText : ''
      }
    },
    setEastRegionNewOverdueResult(state, payload) {
      state.eastRegionNewOverdueResult = {
        resultText: payload && payload.resultText ? payload.resultText : '',
        errorText: payload && payload.errorText ? payload.errorText : ''
      }
    },
    setOverdueCompressionSummaryResult(state, payload) {
      state.overdueCompressionSummaryResult = {
        resultText: payload && payload.resultText ? payload.resultText : ''
      }
    },
    setMainInHandContractAmountResult(state, payload) {
      state.mainInHandContractAmountResult = {
        resultText: payload && payload.resultText ? payload.resultText : '',
        errorText: payload && payload.errorText ? payload.errorText : ''
      }
    },
    setProjectTrackingSummaryResult(state, payload) {
      state.projectTrackingSummaryResult = {
        resultText: payload && payload.resultText ? payload.resultText : '',
        errorText: payload && payload.errorText ? payload.errorText : ''
      }
    },
    setStageDynamicsSummaryResult(state, payload) {
      state.stageDynamicsSummaryResult = {
        resultText: payload && payload.resultText ? payload.resultText : '',
        errorText: payload && payload.errorText ? payload.errorText : ''
      }
    },
    setBidImplementationSummaryResult(state, payload) {
      state.bidImplementationSummaryResult = {
        resultText: payload && payload.resultText ? payload.resultText : '',
        errorText: payload && payload.errorText ? payload.errorText : ''
      }
    },
    setOutboundProjectFollowupResult(state, payload) {
      state.outboundProjectFollowupResult = {
        resultText: payload && payload.resultText ? payload.resultText : '',
        errorText: payload && payload.errorText ? payload.errorText : ''
      }
    },
    setCompetitorAnalysisStrictResult(state, payload) {
      state.competitorAnalysisStrictResult = {
        resultText: payload && payload.resultText ? payload.resultText : '',
        errorText: payload && payload.errorText ? payload.errorText : ''
      }
    },
    setNextMonthForecastResult(state, payload) {
      state.nextMonthForecastResult = {
        resultText: payload && payload.resultText ? payload.resultText : '',
        errorText: payload && payload.errorText ? payload.errorText : ''
      }
    },
    setInHandThreeSheetTableResult(state, payload) {
      state.inHandThreeSheetTableResult = {
        statusText: payload && payload.statusText ? payload.statusText : '',
        errorText: payload && payload.errorText ? payload.errorText : '',
        data: payload && payload.data
          ? payload.data
          : { targetMonth: '', newSignRows: [], revenueRows: [], receiptRows: [] }
      }
    }
  }
})

// 中文注释：每次状态变更后统一落盘，避免路由切换或刷新导致数据丢失。
store.subscribe((_mutation, state) => {
  const toPersist = {
    excelParserResult: state.excelParserResult,
    competitorAnalysisResult: state.competitorAnalysisResult,
    ledgerReportResult: state.ledgerReportResult,
    overdueReceiptMonthlyResult: state.overdueReceiptMonthlyResult,
    eastRegionOverdueStockResult: state.eastRegionOverdueStockResult,
    eastRegionNewOverdueResult: state.eastRegionNewOverdueResult,
    overdueCompressionSummaryResult: state.overdueCompressionSummaryResult,
    mainInHandContractAmountResult: state.mainInHandContractAmountResult,
    projectTrackingSummaryResult: state.projectTrackingSummaryResult,
    stageDynamicsSummaryResult: state.stageDynamicsSummaryResult,
    bidImplementationSummaryResult: state.bidImplementationSummaryResult,
    outboundProjectFollowupResult: state.outboundProjectFollowupResult,
    competitorAnalysisStrictResult: state.competitorAnalysisStrictResult,
    nextMonthForecastResult: state.nextMonthForecastResult,
    inHandThreeSheetTableResult: state.inHandThreeSheetTableResult
  }

  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(toPersist))
  } catch (error) {
    // 中文注释：本地存储不可用时静默降级，不影响页面计算。
  }
})

export default store
