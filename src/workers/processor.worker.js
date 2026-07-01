import * as pdfjsLib from 'pdfjs-dist';
import * as XLSX from 'xlsx';

// PDF.js: 워커 안에서 실행이므로 내부 워커 비활성화
pdfjsLib.GlobalWorkerOptions.workerSrc = '';

// 처리 대기 큐 (flush 신호 전까지 순차 처리)
const queue = [];
let flushed = false;

self.onmessage = async (e) => {
  const msg = e.data;
  if (msg.type === 'process') {
    queue.push(msg);
    if (flushed) await processNext();
  } else if (msg.type === 'flush') {
    flushed = true;
    await drainQueue();
  }
};

async function drainQueue() {
  while (queue.length > 0) await processNext();
  self.postMessage({ type: 'done' });
}

async function processNext() {
  const task = queue.shift();
  if (!task) return;
  try {
    if (task.ext === 'pdf') await processPdf(task);
    else if (task.ext === 'xlsx') await processExcel(task);
    else self.postMessage({ type: 'error', id: task.id, name: task.name, message: `지원하지 않는 형식: .${task.ext}` });
  } catch (e) {
    self.postMessage({ type: 'error', id: task.id, name: task.name, message: String(e) });
  }
}

// ── PDF 처리 ──────────────────────────────────────────────────────
async function processPdf({ id, name, outputPath, data, keywords }) {
  self.postMessage({ type: 'progress', id, status: '처리중' });
  self.postMessage({ type: 'log', message: `▶ PDF 처리 시작: ${name}` });

  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const pattern = buildPattern(keywords);
  const found = [];

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const text = content.items.map(i => i.str).join(' ');
    const matches = [...text.matchAll(pattern)].map(m => m[0]);
    if (matches.length > 0) {
      found.push({ page: p, keywords: [...new Set(matches)] });
    }
  }

  // 결과를 Excel 보고서로 출력
  const rows = found.flatMap(({ page, keywords: kws }) =>
    kws.map(kw => ({ 페이지: page, 발견된단어: kw }))
  );
  const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ 페이지: '-', 발견된단어: '탐지된 단어 없음' }]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '탐지결과');
  const xlsxPath = outputPath.replace(/\.pdf$/i, '_결과.xlsx');
  const xlsxBytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

  self.postMessage(
    { type: 'result', id, name, outputPath: xlsxPath, data: new Uint8Array(xlsxBytes) },
    [new Uint8Array(xlsxBytes).buffer]
  );
  self.postMessage({ type: 'log', message: `✅ PDF 완료 | 탐지 ${rows.length}건 → ${xlsxPath}` });
}

// ── Excel 처리 ────────────────────────────────────────────────────
async function processExcel({ id, name, outputPath, data, keywords }) {
  self.postMessage({ type: 'progress', id, status: '처리중' });
  self.postMessage({ type: 'log', message: `▶ Excel 처리 시작: ${name}` });

  const pattern = buildPattern(keywords);
  const wb = XLSX.read(data, { type: 'array' });
  const wsName = wb.SheetNames[0];
  const ws = wb.Sheets[wsName];

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const header = rows[0] ?? [];

  // 결과 컬럼 추가 (기존 열 이후)
  const detectedCol = '발견여부';
  const keywordsCol = '발견된 단어';
  header.push(detectedCol, keywordsCol);

  const resultRows = [header];
  let detectedCount = 0;

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const cellText = row.join(' ');
    const matches = [...cellText.matchAll(pattern)].map(m => m[0]);
    const unique = [...new Set(matches)];
    const detected = unique.length > 0;
    if (detected) detectedCount++;
    resultRows.push([...row, detected ? 'TRUE' : 'FALSE', unique.join(', ')]);
  }

  const newWs = XLSX.utils.aoa_to_sheet(resultRows);
  wb.Sheets[wsName] = newWs;
  const outBytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

  self.postMessage(
    { type: 'result', id, name, outputPath, data: new Uint8Array(outBytes) },
    [new Uint8Array(outBytes).buffer]
  );
  self.postMessage({
    type: 'log',
    message: `✅ Excel 완료 | 전체 ${resultRows.length - 1}행 중 탐지 ${detectedCount}행 → ${outputPath}`,
  });
}

// ── 공통 ──────────────────────────────────────────────────────────
function buildPattern(keywords) {
  const escaped = [...keywords].sort((a, b) => b.length - a.length).map(escape);
  return new RegExp(escaped.join('|'), 'gi');
}

function escape(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
