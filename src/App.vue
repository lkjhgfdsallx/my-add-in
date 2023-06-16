<script setup>
import { useRouter } from 'vue-router'
import { onMounted } from 'vue'

onMounted(() => {
  const urlParams = new URLSearchParams(window.location.search)
  const link = urlParams.get('1ink') // 获取名为 "1ink" 的参数值

  console.log(link) // 输出参数值
  useRouter().push(`/${link}`)
})
</script>

<template>
  <div id="app">
    <router-view />
  </div>
</template>

<style></style>
