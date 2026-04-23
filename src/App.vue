<script setup>
import { RouterView } from 'vue-router'
import { NConfigProvider, NGlobalStyle } from 'naive-ui'
</script>

<template>
  <NConfigProvider>
    <NGlobalStyle />
    <RouterView />
  </NConfigProvider>
</template>
