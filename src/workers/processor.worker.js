import * as pdfjsLib from 'pdfjs-dist';
import * as XLSX from 'xlsx';
import { PDFDocument, PDFName, PDFNumber, PDFHexString } from 'pdf-lib';
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
const CONSEC_SPACE_PATTERN = /(?<! ) {2,5}(?! )/g;
const CONSEC_SPACE_LABEL = '연속 공백';

async function processPdf({ id, name, outputPath, data, keywords, detectConsecutiveSpaces }) {
  self.postMessage({ type: 'progress', id, status: '처리중' });
  self.postMessage({ type: 'log', message: `▶ PDF 처리 시작: ${name}` });

  const pattern = buildPattern(keywords);

  // Step 1: pdfjs-dist로 텍스트 위치 추출
  // data.slice()로 복사 — pdfjs가 내부적으로 버퍼를 소비할 수 있으므로
  const pdf = await pdfjsLib.getDocument({ data: data.slice() }).promise;
  const pageHighlights = []; // { pageIndex, rects, keywords }
  let totalFound = 0;

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const { rects, matchedKeywords } = findKeywordRectsAndKeywords(content.items, pattern, detectConsecutiveSpaces);
    if (rects.length > 0) {
      pageHighlights.push({ pageIndex: p - 1, rects, keywords: matchedKeywords });
      totalFound += rects.length;
    }
  }

  // Step 2: pdf-lib으로 하이라이트 + 북마크 추가
  const pdfDoc = await PDFDocument.load(data);

  for (const { pageIndex, rects } of pageHighlights) {
    const page = pdfDoc.getPage(pageIndex);
    for (const rect of rects) {
      addHighlightAnnotation(pdfDoc, page, rect);
    }
  }

  // 북마크: 페이지별·키워드별 1개 (중복 제거)
  const outlineItems = pageHighlights.flatMap(({ pageIndex, keywords }) =>
    [...keywords].map(kw => ({ title: `P${pageIndex + 1}: ${kw}`, pageIndex }))
  );
  if (outlineItems.length > 0) addOutlines(pdfDoc, outlineItems);

  const outBytes = await pdfDoc.save();
  const resultData = new Uint8Array(outBytes);

  self.postMessage(
    { type: 'result', id, name, outputPath, data: resultData },
    [resultData.buffer]
  );
  self.postMessage({ type: 'log', message: `✅ PDF 완료 | 탐지 ${totalFound}건, 북마크 ${outlineItems.length}개 → ${outputPath}` });
}

// 텍스트 아이템 목록에서 키워드 위치와 매칭된 키워드 Set을 반환
function findKeywordRectsAndKeywords(items, pattern, detectConsecutiveSpaces = false) {
  const segments = items
    .filter(item => item.str)
    .map(item => ({
      text: item.str,
      x: item.transform[4],
      y: item.transform[5],
      width: item.width,
      fontSize: Math.abs(item.transform[3]),
    }));

  if (segments.length === 0) return { rects: [], matchedKeywords: new Set() };

  // 전체 텍스트 문자열 + 위치 맵 구성
  let fullText = '';
  const posMap = [];

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
  const matchedKeywords = new Set();

  for (const match of fullText.matchAll(pattern)) {
    matchedKeywords.add(match[0]);

    const start = match.index;
    const end = start + match[0].length;

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
      const xStart = seg.x + seg.width * (minChar / charCount);
      const xEnd   = seg.x + seg.width * ((maxChar + 1) / charCount);
      const padX = seg.fontSize * 0.15;
      const padY = seg.fontSize * 0.15;

      rects.push({
        x: xStart - padX,
        y: seg.y - seg.fontSize * 0.2 - padY,
        w: (xEnd - xStart) + padX * 2,
        h: seg.fontSize * 1.4 + padY,
      });
    }
  }

  // 연속 공백 검사: 세그먼트별로 직접 검사 (인위적 gap 공백과 분리)
  if (detectConsecutiveSpaces) {
    for (const seg of segments) {
      CONSEC_SPACE_PATTERN.lastIndex = 0;
      for (const match of seg.text.matchAll(CONSEC_SPACE_PATTERN)) {
        matchedKeywords.add(CONSEC_SPACE_LABEL);
        const minChar = match.index;
        const maxChar = match.index + match[0].length - 1;
        const charCount = seg.text.length || 1;
        const xStart = seg.x + seg.width * (minChar / charCount);
        const xEnd   = seg.x + seg.width * ((maxChar + 1) / charCount);
        const padX = seg.fontSize * 0.15;
        const padY = seg.fontSize * 0.15;
        rects.push({
          x: xStart - padX,
          y: seg.y - seg.fontSize * 0.2 - padY,
          w: (xEnd - xStart) + padX * 2,
          h: seg.fontSize * 1.4 + padY,
        });
      }
    }
  }

  return { rects, matchedKeywords };
}

// pdf-lib으로 PDF 북마크(Outlines) 추가
function addOutlines(pdfDoc, items) {
  // items: [{ title: string, pageIndex: number }]
  const { context, catalog } = pdfDoc;

  // 한글 등 유니코드 타이틀 → UTF-16 BE HexString (PDF 스펙)
  function pdfTitle(str) {
    const bytes = [0xFE, 0xFF];
    for (let i = 0; i < str.length; i++) {
      const c = str.charCodeAt(i);
      bytes.push((c >> 8) & 0xFF, c & 0xFF);
    }
    return PDFHexString.of(bytes.map(b => b.toString(16).padStart(2, '0')).join(''));
  }

  // 각 북마크 딕셔너리 생성 (Parent는 나중에 설정)
  const itemDicts = items.map(({ title, pageIndex }) => {
    const pageRef = pdfDoc.getPage(pageIndex).ref;
    return context.obj({
      Title: pdfTitle(title),
      Dest: context.obj([pageRef, PDFName.of('XYZ'), null, null, null]),
    });
  });

  const itemRefs = itemDicts.map(d => context.register(d));

  // 루트 Outlines 딕셔너리
  const rootDict = context.obj({
    Type: PDFName.of('Outlines'),
    First: itemRefs[0],
    Last: itemRefs[itemRefs.length - 1],
    Count: PDFNumber.of(items.length),
  });
  const rootRef = context.register(rootDict);

  // 각 아이템에 Parent / Prev / Next 연결
  for (let i = 0; i < itemDicts.length; i++) {
    itemDicts[i].set(PDFName.of('Parent'), rootRef);
    if (i > 0) itemDicts[i].set(PDFName.of('Prev'), itemRefs[i - 1]);
    if (i < itemDicts.length - 1) itemDicts[i].set(PDFName.of('Next'), itemRefs[i + 1]);
  }

  // 카탈로그에 등록 + 북마크 패널 자동 열기
  catalog.set(PDFName.of('Outlines'), rootRef);
  catalog.set(PDFName.of('PageMode'), PDFName.of('UseOutlines'));
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
async function processExcel({ id, name, outputPath, data, keywords, detectConsecutiveSpaces }) {
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
    if (detectConsecutiveSpaces) {
      const hasConsec = row.some(cell => CONSEC_SPACE_PATTERN.test(String(cell)));
      // RegExp.test()은 lastIndex를 변경하므로 초기화
      CONSEC_SPACE_PATTERN.lastIndex = 0;
      if (hasConsec) unique.push(CONSEC_SPACE_LABEL);
    }
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
