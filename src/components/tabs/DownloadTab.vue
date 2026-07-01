<script setup>
import { useAppStore } from '../../stores/app.js';
const store = useAppStore();
</script>

<template>
  <div class="h-full overflow-y-auto px-6 py-5 text-[14px] text-[#343A40] leading-relaxed">
    <h2 class="text-2xl font-bold text-[#2C3E50] border-b-2 border-[#E9ECEF] pb-2 mb-4">
      ⬇️ 최신 버전 다운로드 및 업데이트
    </h2>

    <p class="mt-4">
      현재 실행 중인 버전: <strong class="text-primary">{{ store.version }}</strong>
    </p>
    <p class="text-[13px] text-app-muted mt-1">
      기능 개선 및 버그가 수정된 최신 버전은 GitHub Release 페이지에서 다운로드 가능합니다.
    </p>

    <div class="bg-app-bg border border-[#DEE2E6] rounded-lg p-8 mt-6 text-center">
      <h3 class="text-lg font-bold text-[#343A40] mb-2">최신 버전 프로그램 다운로드하기</h3>
      <p class="text-[13px] text-app-muted mb-6">
        아래 버튼을 클릭하면 GitHub에서 가장 최신 배포 파일을 다운로드합니다.
      </p>
      <button
        class="bg-success hover:bg-success-hover text-white font-bold text-lg px-10 py-4 rounded-lg shadow transition-colors"
        @click="store.openUrl('https://github.com/itmir913/WordFinderApp/releases/latest/download/WordFinderApp.zip')"
      >
        📥 최신 버전 프로그램 다운로드하기
      </button>
    </div>

    <div class="bg-[#E7F1FF] border-l-4 border-primary px-4 py-3 mt-6 text-[13px]">
      <p class="font-bold text-[#0056B3] mb-2">💡 업데이트 방법</p>
      <ul class="list-disc pl-5 space-y-1 text-[#343A40]">
        <li>위 <strong>다운로드 버튼</strong>을 클릭하여 최신 버전의 압축 파일(ZIP)을 다운로드합니다.</li>
        <li>기존에 사용하시던 프로그램을 삭제하시고 새로 다운로드받은 파일을 압축 해제하여 사용하시면 됩니다.</li>
      </ul>
    </div>
  </div>
</template>
