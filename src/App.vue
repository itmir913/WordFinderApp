<script setup>
import { onMounted } from 'vue';
import { useAppStore } from './stores/app.js';
import { getCurrentWindow } from '@tauri-apps/api/window';
import TitleBar from './components/TitleBar.vue';
import CsvSection from './components/CsvSection.vue';
import ActionBar from './components/ActionBar.vue';
import GuideTab from './components/tabs/GuideTab.vue';
import FileTab from './components/tabs/FileTab.vue';
import LogTab from './components/tabs/LogTab.vue';
import DownloadTab from './components/tabs/DownloadTab.vue';

const store = useAppStore();

const TABS = [
  { label: '📖 사용 방법', component: GuideTab },
  { label: '📁 검색 대상 파일', component: FileTab },
  { label: '📜 시스템 로그', component: LogTab },
  { label: '⬇️ 최신버전 다운로드', component: DownloadTab },
];

onMounted(async () => {
  await store.loadDefaultCsv();

  // Tauri 파일 드래그앤드롭 이벤트 등록
  const appWindow = getCurrentWindow();
  await appWindow.onDragDropEvent((event) => {
    if (event.payload.type === 'drop') {
      const paths = event.payload.paths ?? [];
      const csvPaths = paths.filter(p => p.toLowerCase().endsWith('.csv'));
      const filePaths = paths.filter(p => /\.(pdf|xlsx)$/i.test(p));

      if (csvPaths.length === 1) {
        store.loadCsvFromPath(csvPaths[0]);
      } else if (csvPaths.length > 1) {
        store.addLog('⚠️ CSV 파일은 한 번에 하나만 등록 가능합니다.');
      }
      if (filePaths.length > 0) {
        store.addFiles(filePaths);
        store.activeTab = 1;
      }
    }
  });
});
</script>

<template>
  <div class="flex flex-col h-screen w-full bg-app-bg text-[#333333] text-base select-none">
    <TitleBar />
    <div class="flex flex-col flex-1 min-h-0 gap-2.5 px-4 pb-3 pt-4">
      <CsvSection />

      <!-- 탭 영역 -->
      <div class="flex flex-col flex-1 min-h-0 border border-[#E9ECEF] rounded-lg overflow-hidden bg-white">
        <!-- 탭 네비게이션 -->
        <div class="flex border-b border-[#E9ECEF] bg-app-bg shrink-0">
          <button
            v-for="(tab, i) in TABS"
            :key="i"
            class="px-5 py-2 font-bold text-base border-t border-x border-[#E9ECEF] rounded-t-lg -mb-px transition-colors"
            :class="store.activeTab === i
              ? 'bg-white text-[#212529] border-b-white'
              : 'bg-app-bg text-app-muted hover:bg-[#F1F3F5]'"
            @click="store.activeTab = i"
          >
            {{ tab.label }}
          </button>
        </div>

        <!-- 탭 콘텐츠 -->
        <div class="flex-1 min-h-0 overflow-auto">
          <component :is="TABS[store.activeTab].component" />
        </div>
      </div>

      <ActionBar />
    </div>
  </div>
</template>
