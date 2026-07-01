<script setup>
import { useAppStore } from '../stores/app.js';

const store = useAppStore();

function start() {
  if (!store.keywordCount) { alert('검색할 단어가 담긴 CSV를 먼저 입력하세요.'); return; }
  if (!store.fileCount) { alert('처리할 PDF, Excel 파일을 추가하세요.'); return; }
  store.startProcessing();
}
</script>

<template>
  <div class="flex justify-end gap-2 shrink-0">
    <button
      class="h-10 px-8 font-bold rounded-md text-white transition-colors"
      :class="store.isProcessing
        ? 'bg-[#A9E3D4] cursor-not-allowed'
        : 'bg-success hover:bg-success-hover cursor-pointer'"
      :disabled="store.isProcessing"
      @click="start"
    >
      ▶ 처리 시작
    </button>
    <button
      class="h-10 px-8 font-bold rounded-md text-white transition-colors"
      :class="!store.isProcessing
        ? 'bg-[#FFE3E3] text-[#FFA8A8] cursor-not-allowed'
        : 'bg-danger hover:bg-danger-hover cursor-pointer'"
      :disabled="!store.isProcessing"
      @click="store.stopProcessing()"
    >
      ⛔ 중지
    </button>
  </div>
</template>
