import { defineStore } from 'pinia';
import { invoke } from '@tauri-apps/api/core';
import { getVersion } from '@tauri-apps/api/app';
import { open } from '@tauri-apps/plugin-dialog';
import { openUrl } from '@tauri-apps/plugin-opener';

export const useAppStore = defineStore('app', {
  state: () => ({
    version: '',

    // CSV / 키워드
    csvPath: '',
    keywords: [],

    // 파일 목록
    files: [], // { id, name, path, status: '대기'|'처리중'|'성공'|'실패' }
    nextId: 1,

    // 로그
    logs: [],

    // 처리 상태
    isProcessing: false,
    stopRequested: false,

    // 활성 탭
    activeTab: 0,

    // Web Worker
    _worker: null,

    // 최신 버전 (null: 미조회, '': 조회 실패)
    latestVersion: null,
  }),

  getters: {
    keywordCount: (state) => state.keywords.length,
    fileCount: (state) => state.files.length,
  },

  actions: {
    // ── 앱 초기화 ─────────────────────────────────────
    async init() {
      this.version = await getVersion();
      await this.loadDefaultCsv();
    },

    // ── CSV ──────────────────────────────────────────
    async loadCsvFromPath(path) {
      try {
        const content = await invoke('read_file_bytes', { path });
        this._parseCsvBytes(content, path);
      } catch (e) {
        this.addLog(`❌ CSV 로드 실패: ${e}`);
        throw e;
      }
    },

    async selectCsv() {
      const path = await open({
        filters: [{ name: 'CSV 파일', extensions: ['csv'] }],
        multiple: false,
      });
      if (path) await this.loadCsvFromPath(path);
    },

    async loadDefaultCsv() {
      try {
        const content = await invoke('load_default_csv');
        this._parseCsvString(content, '(기본값)');
      } catch {
        // default.csv 없으면 조용히 무시
      }
    },

    _parseCsvBytes(bytes, path) {
      const decoder = new TextDecoder('utf-8');
      const text = decoder.decode(new Uint8Array(bytes));
      this._parseCsvString(text, path);
    },

    _parseCsvString(text, path) {
      const lines = text.split(/\r?\n/).slice(1); // 헤더 제거
      const kws = [...new Set(
        lines.map(l => l.split(',')[0].trim()).filter(Boolean)
      )];
      this.keywords = kws;
      this.csvPath = path;
      this.addLog(`✅ CSV 로드 완료: ${kws.length}개 | ${path}`);
    },

    // ── 파일 목록 ─────────────────────────────────────
    addFiles(paths) {
      const existing = new Set(this.files.map(f => f.path));
      let added = 0;
      for (const path of paths) {
        const ext = path.split('.').pop().toLowerCase();
        if (!['pdf', 'xlsx'].includes(ext)) continue;
        if (existing.has(path)) continue;
        const name = path.replace(/\\/g, '/').split('/').pop();
        this.files.push({ id: this.nextId++, name, path, status: '대기' });
        added++;
      }
      if (added) this.addLog(`📂 ${added}개 파일 추가됨.`);
    },

    async selectFiles() {
      const paths = await open({
        filters: [{ name: '지원 파일', extensions: ['pdf', 'xlsx'] }],
        multiple: true,
      });
      if (paths) this.addFiles(Array.isArray(paths) ? paths : [paths]);
    },

    removeFile(id) {
      const idx = this.files.findIndex(f => f.id === id);
      if (idx !== -1) {
        this.addLog(`🗑 '${this.files[idx].name}' 파일 제외.`);
        this.files.splice(idx, 1);
      }
    },

    clearFiles() {
      this.files = [];
      this.addLog('🗑 파일 목록 전체 초기화.');
    },

    updateFileStatus(id, status) {
      const file = this.files.find(f => f.id === id);
      if (file) file.status = status;
    },

    // ── 처리 ─────────────────────────────────────────
    async startProcessing() {
      if (this.keywords.length === 0 || this.files.length === 0) return;

      this.isProcessing = true;
      this.stopRequested = false;
      this.files.forEach(f => (f.status = '대기'));
      this.addLog('🚀 처리 시작');
      this.activeTab = 2; // 로그 탭으로 이동

      this._worker = new Worker(
        new URL('../workers/processor.worker.js', import.meta.url),
        { type: 'module' }
      );

      this._worker.onmessage = async (e) => {
        const msg = e.data;
        switch (msg.type) {
          case 'progress':
            this.updateFileStatus(msg.id, msg.status);
            break;
          case 'log':
            this.addLog(msg.message);
            break;
          case 'result':
            try {
              await invoke('write_file_bytes', { path: msg.outputPath, data: Array.from(msg.data) });
              this.updateFileStatus(msg.id, '성공');
              this.addLog(`✅ 완료 → ${msg.outputPath}`);
            } catch (e) {
              this.updateFileStatus(msg.id, '실패');
              this.addLog(`❌ 저장 실패 [${msg.name}]: ${e}`);
            }
            break;
          case 'error':
            this.updateFileStatus(msg.id, '실패');
            this.addLog(`❌ 오류 [${msg.name}]: ${msg.message}`);
            break;
          case 'done':
            this.addLog('✅ 모든 처리 완료.');
            this._cleanup();
            break;
        }
      };

      // 각 파일 바이트 읽기 → 워커로 전송
      for (const file of this.files) {
        if (this.stopRequested) {
          this.addLog('⛔ 사용자에 의해 중지되었습니다.');
          break;
        }
        try {
          const bytes = await invoke('read_file_bytes', { path: file.path });
          const ext = file.name.split('.').pop().toLowerCase();
          const outputPath = file.path.replace(/[^/\\]+$/, `output_${file.name}`);
          const data = new Uint8Array(bytes);
          this._worker.postMessage(
            { type: 'process', id: file.id, name: file.name, ext, outputPath, data, keywords: [...this.keywords] },
            [data.buffer]
          );
        } catch (e) {
          this.updateFileStatus(file.id, '실패');
          this.addLog(`❌ 파일 읽기 실패 [${file.name}]: ${e}`);
        }
      }

      this._worker.postMessage({ type: 'flush' });
    },

    stopProcessing() {
      this.stopRequested = true;
      this.addLog('⛔ 중지 요청...');
    },

    _cleanup() {
      this._worker?.terminate();
      this._worker = null;
      this.isProcessing = false;
      this.stopRequested = false;
    },

    // ── 로그 ─────────────────────────────────────────
    addLog(message) {
      const now = new Date().toLocaleTimeString('ko-KR');
      this.logs.push(`[${now}] ${message}`);
    },

    // ── 최신 버전 확인 ────────────────────────────────
    async fetchLatestVersion() {
      try {
        const res = await fetch('https://api.github.com/repos/itmir913/WordFinderApp/releases/latest');
        const data = await res.json();
        this.latestVersion = data.tag_name ?? '';
      } catch {
        this.latestVersion = '';
      }
    },

    // ── 외부 링크 ─────────────────────────────────────
    async openUrl(url) {
      await openUrl(url);
    },
  },
});
