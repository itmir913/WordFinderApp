<script setup>
import { watch, nextTick, ref } from 'vue';
import { useAppStore } from '../../stores/app.js';

const store = useAppStore();
const logEl = ref(null);

watch(() => store.logs.length, async () => {
  await nextTick();
  if (logEl.value) logEl.value.scrollTop = logEl.value.scrollHeight;
});
</script>

<template>
  <div class="h-full p-2.5">
    <div
      ref="logEl"
      class="h-full bg-[#212529] text-[#C1C9D2] rounded-lg p-3 overflow-y-auto font-mono text-base leading-6"
    >
      <div v-for="(line, i) in store.logs" :key="i">{{ line }}</div>
      <div v-if="store.logs.length === 0" class="text-[#6C757D]">로그가 여기 표시됩니다.</div>
    </div>
  </div>
</template>
