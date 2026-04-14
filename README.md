# lk-table
刘柯表格

## Word 模板自动填充

页面已支持将各模块结果动态注入 Word 模板并下载。

### 使用步骤

1. 先完成 Excel 上传并点击“一键生成全部页面结果”。
2. 在页面顶部上传 Word 模板（必须是 .docx 格式）。
3. 点击“生成 Word 月报”，系统会下载填充后的文档。

### 模板怎么改

在 Word 模板中，把需要替换的位置改成占位符，格式为花括号。

示例：

- 今日日期：{reportDate}
- 业绩台账段落：{ledgerReportText}
- 逾期压降段落：{overdueCompressionText}

### 可用占位符

- reportDate
- nextMonthForecastText
- competitorStrictText
- ledgerReportText
- overdueReceiptText
- stockOverdueText
- newOverdueText
- inHandContractText
- projectTrackingText
- stageDynamicsText
- bidImplementationText
- outboundFollowupText
- overdueCompressionText
- competitorAnalysisText
- excelParserText

### 注意事项

1. 仅支持 .docx，不支持 .doc。
2. 占位符必须完整，不能拆分成多个文本片段。
3. 建议先在模板中用纯文本输入占位符，再做字体和段落样式调整。
