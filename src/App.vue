<template>
  <div id="app">
    <ExcelViewer />
  </div>
</template>

<script>
import ExcelViewer from './components/ExcelViewer.vue'

export default {
  name: 'App',
  components: {
    ExcelViewer
  }
}
// 上述问题没有得到解决，问题描述详细为在列合并单元格所合并的单元格一行最后会在最后一列多一个单元格，例如第6行的第3列、第7行的第3列、第8行的第3列合并成一个单元格后，在网页上显示在第7行与第8行最后一列右侧多了一个单元格，而第6行单元格最后一列正常没有多出单元格
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
