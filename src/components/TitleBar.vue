<script setup>
import { getCurrentWindow } from '@tauri-apps/api/window';
import { Info, Minus, X } from '@lucide/vue';
import { useAppStore } from '../stores/app.js';

const store = useAppStore();
const appWindow = getCurrentWindow();

function minimize() { appWindow.minimize(); }
function close() { appWindow.close(); }

function showAbout() {
  alert(`학교생활기록부 일괄 점검 프로그램\n버전: ${store.version}\n제작자: 운양고등학교 이종환T`);
}

function onDragStart(e) {
  if (e.button === 0) {
    appWindow.startDragging();
  }
}
</script>

<template>
  <div
    data-tauri-drag-region
    class="flex items-center h-10 px-4 shrink-0 bg-app-bg border-b border-app-border cursor-move"
    @mousedown="onDragStart"
  >
    <span class="font-bold text-[15px] text-[#343A40] pointer-events-none">
      🔍 학교생활기록부 일괄 점검 프로그램 {{ store.version }}
    </span>

    <div class="ml-auto flex">
      <button
        class="w-[45px] h-10 flex items-center justify-center text-app-muted hover:bg-[#E9ECEF] hover:text-[#343A40] transition-colors cursor-default"
        @mousedown.stop
        @click="showAbout"
      >
        <Info :size="18" />
      </button>
      <button
        class="w-[45px] h-10 flex items-center justify-center text-app-muted hover:bg-[#E9ECEF] hover:text-[#343A40] transition-colors cursor-default"
        @mousedown.stop
        @click="minimize"
      >
        <Minus :size="16" />
      </button>
      <button
        class="w-[45px] h-10 flex items-center justify-center text-app-muted hover:bg-danger hover:text-white transition-colors cursor-default"
        @mousedown.stop
        @click="close"
      >
        <X :size="16" />
      </button>
    </div>
  </div>
</template>
