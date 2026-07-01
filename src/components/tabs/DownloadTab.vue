<script setup>
import { onMounted, computed } from 'vue';
import { useAppStore } from '../../stores/app.js';

const store = useAppStore();

onMounted(() => {
  store.fetchLatestVersion();
});

const versionStatus = computed(() => {
  if (store.latestVersion === null) return 'loading';
  if (store.latestVersion === '') return 'error';
  return store.latestVersion === store.version ? 'latest' : 'outdated';
});
</script>

<template>
  <div class="h-full overflow-y-auto px-6 py-5 text-base text-[#343A40] leading-relaxed">

    <!-- ── 프로그램 정보 ── -->
    <section class="border border-[#DEE2E6] rounded-xl p-6">
      <h2 class="text-xl font-bold text-[#2C3E50] mb-4">ℹ️ 프로그램 정보</h2>

      <div class="grid grid-cols-[auto_1fr] gap-x-6 gap-y-2 items-baseline">
        <span class="font-bold text-[#495057]">버전</span>
        <span class="text-primary font-bold">{{ store.version }}</span>

        <span class="font-bold text-[#495057]">제작자</span>
        <span>운양고등학교 이종환T</span>

        <span class="font-bold text-[#495057]">이메일</span>
        <a
          href="#"
          class="text-[#3498DB] underline"
          @click.prevent="store.openUrl('mailto:hello@luminousky.com')"
        >hello@luminousky.com</a>

        <span class="font-bold text-[#495057]">공식 홈페이지</span>
        <a
          href="#"
          class="text-[#3498DB] underline"
          @click.prevent="store.openUrl('https://luminousky.com/teacher-utility-kit/neis-wordfinder/')"
        >https://luminousky.com/teacher-utility-kit/neis-wordfinder/</a>

        <span class="font-bold text-[#495057]">GitHub</span>
        <a
          href="#"
          class="text-[#3498DB] underline"
          @click.prevent="store.openUrl('https://github.com/itmir913/WordFinderApp')"
        >https://github.com/itmir913/WordFinderApp</a>
      </div>

      <div class="bg-app-bg rounded-lg px-4 py-3 mt-5 text-[#495057]">
        본 프로그램은 <strong>PolyForm Noncommercial License 1.0.0</strong>을 따릅니다.<br />
        개인 및 학교 등 교육 현장에서 자유롭게 사용하실 수 있습니다.<br />
        <strong class="text-danger">※ 상업적 이용, 유료 판매 및 이를 활용한 모든 영리 활동은 엄격히 금지됩니다.</strong>
      </div>
    </section>

    <!-- ── 최신버전 다운로드 ── -->
    <section class="border border-[#DEE2E6] rounded-xl p-6 mt-5">
      <h2 class="text-xl font-bold text-[#2C3E50] mb-4">⬇️ 최신버전 다운로드</h2>

      <!-- 버전 상태 -->
      <div class="flex items-center gap-3 mb-4">
        <span class="text-[#495057]">현재 버전: <strong class="text-primary">{{ store.version }}</strong></span>
        <span
          v-if="versionStatus === 'loading'"
          class="text-app-muted"
        >🔄 확인 중...</span>
        <span
          v-else-if="versionStatus === 'latest'"
          class="font-bold text-success"
        >✅ 최신버전입니다.</span>
        <span
          v-else-if="versionStatus === 'outdated'"
          class="font-bold text-danger"
        >⚠️ 최신버전이 아닙니다. (최신: {{ store.latestVersion }})</span>
        <span
          v-else
          class="text-app-muted"
        >버전 정보를 불러올 수 없습니다.</span>
      </div>

      <p class="text-app-muted mb-5">
        기능 개선 및 버그가 수정된 최신버전은 GitHub Release 페이지에서 다운로드할 수 있습니다.
      </p>

      <button
        class="w-full bg-success hover:bg-success-hover text-white font-bold text-lg py-4 rounded-lg shadow transition-colors"
        @click="store.openUrl('https://github.com/itmir913/WordFinderApp/releases/latest/download/WordFinderApp.zip')"
      >
        📥 최신버전 다운로드하기
      </button>

      <div class="bg-[#E7F1FF] border-l-4 border-primary px-4 py-3 mt-5">
        <p class="font-bold text-[#0056B3] mb-2">💡 업데이트 방법</p>
        <ul class="list-disc pl-5 space-y-1 text-[#343A40]">
          <li>위 버튼을 클릭하여 최신버전의 압축 파일(ZIP)을 다운로드합니다.</li>
          <li>기존 프로그램을 삭제하고, 새로 다운로드한 파일을 압축 해제하여 사용하세요.</li>
        </ul>
      </div>
    </section>

  </div>
</template>
