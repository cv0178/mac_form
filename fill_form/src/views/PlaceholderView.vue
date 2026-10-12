<script setup>
defineProps({ title: String, description: String });
</script>
<template><main class="content"><section class="coming-soon"><span>✦</span><p class="page-kicker">COMING SOON</p><h1>{{ title }}</h1><p class="page-description">{{ description }}</p><RouterLink class="primary-button" to="/">返回开始填表</RouterLink></section></main></template>
