import * as pdfjsLib from 'pdfjs-dist';
import * as XLSX from 'xlsx';
import { PDFDocument, PDFName, PDFNumber } from 'pdf-lib';
import pdfjsWorkerSrc from 'pdfjs-dist/build/pdf.worker.mjs?url';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorkerSrc;

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

  const pattern = buildPattern(keywords);

  // Step 1: pdfjs-dist로 텍스트 위치 추출
  // data.slice()로 복사 — pdfjs가 내부적으로 버퍼를 소비할 수 있으므로
  const pdf = await pdfjsLib.getDocument({ data: data.slice() }).promise;
  const pageHighlights = [];
  let totalFound = 0;

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const rects = findKeywordRects(content.items, pattern);
    if (rects.length > 0) {
      pageHighlights.push({ pageIndex: p - 1, rects });
      totalFound += rects.length;
    }
  }

  // Step 2: pdf-lib으로 하이라이트 어노테이션 추가
  const pdfDoc = await PDFDocument.load(data);

  for (const { pageIndex, rects } of pageHighlights) {
    const page = pdfDoc.getPage(pageIndex);
    for (const rect of rects) {
      addHighlightAnnotation(pdfDoc, page, rect);
    }
  }

  const outBytes = await pdfDoc.save();
  const resultData = new Uint8Array(outBytes);

  self.postMessage(
    { type: 'result', id, name, outputPath, data: resultData },
    [resultData.buffer]
  );
  self.postMessage({ type: 'log', message: `✅ PDF 완료 | 탐지 ${totalFound}건 → ${outputPath}` });
}

// 텍스트 아이템 목록에서 키워드 위치를 계산해 highlight rect 배열 반환
function findKeywordRects(items, pattern) {
  const segments = items
    .filter(item => item.str)
    .map(item => ({
      text: item.str,
      x: item.transform[4],
      y: item.transform[5],
      width: item.width,
      fontSize: Math.abs(item.transform[3]),
    }));

  if (segments.length === 0) return [];

  // 전체 텍스트 문자열 + 위치 맵 구성
  let fullText = '';
  const posMap = []; // [{segIdx, charIdx, isGap}]

  for (let i = 0; i < segments.length; i++) {
    const seg = segments[i];
    for (let j = 0; j < seg.text.length; j++) {
      posMap.push({ segIdx: i, charIdx: j });
      fullText += seg.text[j];
    }
    if (i < segments.length - 1) {
      posMap.push({ segIdx: i, charIdx: -1, isGap: true });
      fullText += ' ';
    }
  }

  const rects = [];

  for (const match of fullText.matchAll(pattern)) {
    const start = match.index;
    const end = start + match[0].length;

    // 매치 범위 내 문자를 segment 별로 그룹화
    const segGroups = new Map();
    for (let ci = start; ci < end; ci++) {
      const pos = posMap[ci];
      if (!pos || pos.isGap) continue;
      const { segIdx, charIdx } = pos;
      if (!segGroups.has(segIdx)) {
        segGroups.set(segIdx, { minChar: charIdx, maxChar: charIdx });
      } else {
        const g = segGroups.get(segIdx);
        g.minChar = Math.min(g.minChar, charIdx);
        g.maxChar = Math.max(g.maxChar, charIdx);
      }
    }

    for (const [segIdx, { minChar, maxChar }] of segGroups) {
      const seg = segments[segIdx];
      const charCount = seg.text.length || 1;
      // 글자 비율로 x 범위 계산 (비례 추정)
      const xStart = seg.x + seg.width * (minChar / charCount);
      const xEnd   = seg.x + seg.width * ((maxChar + 1) / charCount);

      rects.push({
        x: xStart,
        y: seg.y - seg.fontSize * 0.1,
        w: xEnd - xStart,
        h: seg.fontSize * 1.2,
      });
    }
  }

  return rects;
}

// pdf-lib으로 Highlight 어노테이션 추가
function addHighlightAnnotation(pdfDoc, page, { x, y, w, h }) {
  const { context } = pdfDoc;

  const annot = context.obj({
    Type: PDFName.of('Annot'),
    Subtype: PDFName.of('Highlight'),
    Rect: context.obj([x, y, x + w, y + h]),
    // QuadPoints: 좌상-우상-좌하-우하 순 (PDF 스펙)
    QuadPoints: context.obj([
      x,     y + h,
      x + w, y + h,
      x,     y,
      x + w, y,
    ]),
    C: context.obj([1, 0.9, 0]), // 노란색 (R G B, 0~1)
    F: PDFNumber.of(4),          // Print 플래그
  });

  const annotRef = context.register(annot);

  const annots = page.node.get(PDFName.of('Annots'));
  if (annots) {
    annots.push(annotRef);
  } else {
    page.node.set(PDFName.of('Annots'), context.obj([annotRef]));
  }
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

  header.push('발견여부', '발견된 단어');

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
  const resultData = new Uint8Array(outBytes);

  self.postMessage(
    { type: 'result', id, name, outputPath, data: resultData },
    [resultData.buffer]
  );
  self.postMessage({
    type: 'log',
    message: `✅ Excel 완료 | 전체 ${resultRows.length - 1}행 중 탐지 ${detectedCount}행 → ${outputPath}`,
  });
}

// ── 공통 ──────────────────────────────────────────────────────────
function buildPattern(keywords) {
  const escaped = [...keywords].sort((a, b) => b.length - a.length).map(escapeRegex);
  return new RegExp(escaped.join('|'), 'gi');
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
