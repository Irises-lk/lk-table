<template>
  <div style="max-width:900px;margin:24px auto;font-family:Arial,Helvetica,sans-serif;">
    <h2>区域指挥部指标预计情况解析</h2>
    <!-- 导航：主页为旧的 ExcelParser，新增页面为竞争对手承揽统计 -->
    <nav style="margin-bottom:12px">
      <router-link to="/">指标解析</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/competitor">竞争对手承揽统计</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/ledger-report">业绩台账文本生成</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/overdue-receipt-monthly-report">完成逾期</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/east-region-overdue-stock">存量逾期</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/east-region-new-overdue">新增逾期</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/overdue-compression-summary">逾期压降汇总</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/main-in-hand-contract-amount">在手主办合同总额</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/project-tracking-summary">信息跟踪阶段统计</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/stage-dynamics-summary">三阶段动态统计</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/bid-implementation-summary">中标实施阶段汇总</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/outbound-project-followup">中资外带项目跟进</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/competitor-analysis-strict">竞争对手统计严格版</router-link>
      <span style="margin:0 8px">|</span>
      <router-link to="/next-month-forecast-summary">次月指标预计汇总</router-link>
    </nav>
    <router-view />
  </div>
</template>

<script setup>
// 采用路由视图渲染子页面，组件在 router 中注册
</script>