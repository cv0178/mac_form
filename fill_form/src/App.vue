<script setup>
const navigationItems = [
  { to: '/', icon: '✦', label: '开始填表' },
  { to: '/knowledge', icon: '▣', label: '我的资料库' },
  { to: '/history', icon: '◷', label: '填写记录' },
  { to: '/settings', icon: '⚙', label: '设置' },
];
</script>
<template><div class="app-shell"><aside class="sidebar"><div class="brand"><span class="brand-mark">✎</span><span>智填助手</span></div><nav class="navigation" aria-label="主导航"><RouterLink v-for="item in navigationItems" :key="item.to" :to="item.to" class="nav-item"><span class="nav-icon" aria-hidden="true">{{ item.icon }}</span>{{ item.label }}</RouterLink></nav><div class="local-notice"><span aria-hidden="true">⌂</span><span>所有资料仅保存在此设备</span></div></aside><div class="main-scroll"><RouterView /></div></div></template>
