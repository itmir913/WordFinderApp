<script setup>
import { useAppStore } from '../stores/app.js';

const store = useAppStore();
</script>

<template>
  <div class="bg-white border border-[#DEE2E6] rounded-lg px-4 pt-5 pb-3 relative shrink-0">
    <span class="absolute -top-[11px] left-3 px-2 bg-white text-[#495057] text-xs font-bold">
      Step 1. 검색 기준 설정 (단어 목록)
    </span>

    <div class="flex items-center gap-2">
      <span class="text-[#495057] whitespace-nowrap text-[13px]">단어 목록 CSV:</span>
      <input
        :value="store.csvPath"
        readonly
        placeholder="검색할 단어들이 담긴 CSV 파일을 등록하세요"
        class="flex-1 border border-[#CED4DA] rounded px-2.5 py-1.5 text-[13px] bg-app-bg text-[#495057] cursor-default outline-none"
      />
      <button
        class="px-3 py-1.5 bg-[#E9ECEF] text-[#495057] font-bold text-[13px] rounded hover:bg-[#DEE2E6] transition-colors whitespace-nowrap"
        :disabled="store.isProcessing"
        @click="store.selectCsv()"
      >
        CSV 파일 불러오기
      </button>
    </div>

    <div class="flex justify-end mt-1.5">
      <span
        class="text-xs font-bold"
        :class="store.keywordCount > 0 ? 'text-primary' : 'text-danger'"
      >
        {{ store.keywordCount > 0 ? `검사할 단어 수: ${store.keywordCount}개` : '검색 단어 미등록' }}
      </span>
    </div>
  </div>
</template>
