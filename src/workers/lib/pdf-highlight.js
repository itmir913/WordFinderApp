// PDF 하이라이트/북마크 생성 로직 (순수 함수 — Worker 밖에서도 테스트 가능)
import { PDFName, PDFNumber, PDFHexString } from 'pdf-lib';

// 텍스트 아이템 목록에서 키워드 위치와 원본 키워드 Set을 반환
export function findKeywordRectsAndKeywords(items, pattern, kwMap) {
  const segments = items
    .filter(item => item.str)
    .map(item => {
      // 텍스트 행렬 [a b c d e f]: (a,b)=진행 방향, (c,d)=글자 위쪽 방향
      // 회전된 텍스트(예: [0 s -s 0 e f])에서도 올바른 크기/방향을 얻는다.
      const [a, b, c, d, e, f] = item.transform;
      const advLen = Math.hypot(a, b);
      const fontSize = Math.hypot(c, d) || Math.abs(item.height) || 0;
      return {
        text: item.str,
        x: e,
        y: f,
        width: item.width,
        fontSize,
        // 진행 방향 단위벡터
        ux: advLen ? a / advLen : 1,
        uy: advLen ? b / advLen : 0,
        // 위쪽 단위벡터
        vx: fontSize ? c / fontSize : 0,
        vy: fontSize ? d / fontSize : 1,
      };
    });

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
    const originalKw = kwMap.get(match[0].toLowerCase()) ?? match[0];

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
      if (!(seg.fontSize > 0)) continue; // 글자 크기를 못 구하면 하이라이트 생략

      const charCount = seg.text.length || 1;
      const advStart = seg.width * (minChar / charCount) - seg.fontSize * 0.08;
      const advEnd   = seg.width * ((maxChar + 1) / charCount) + seg.fontSize * 0.08;
      const above = seg.fontSize * 0.88;  // 베이스라인 위(어센더)
      const below = seg.fontSize * 0.26;  // 베이스라인 아래(디센더)

      // 텍스트 방향을 따라 네 꼭짓점 계산 (회전 텍스트 대응)
      const at = (adv, up) => ({
        x: seg.x + seg.ux * adv + seg.vx * up,
        y: seg.y + seg.uy * adv + seg.vy * up,
      });

      rects.push([
        at(advStart, above),  // 좌상
        at(advEnd,   above),  // 우상
        at(advStart, -below), // 좌하
        at(advEnd,   -below), // 우하
      ]);
      // rect 생성 성공 시에만 키워드 등록 (rect 없는 매칭이 북마크에 포함되는 것 방지)
      matchedKeywords.add(originalKw);
    }
  }

  return { rects, matchedKeywords };
}

// pdf-lib으로 PDF 북마크(Outlines) 추가
export function addOutlines(pdfDoc, items) {
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

  const itemDicts = items.map(({ title, pageIndex }) => {
    const pageRef = pdfDoc.getPage(pageIndex).ref;
    return context.obj({
      Title: pdfTitle(title),
      Dest: context.obj([pageRef, PDFName.of('XYZ'), null, null, null]),
    });
  });

  const itemRefs = itemDicts.map(d => context.register(d));

  const rootDict = context.obj({
    Type: PDFName.of('Outlines'),
    First: itemRefs[0],
    Last: itemRefs[itemRefs.length - 1],
    Count: PDFNumber.of(items.length),
  });
  const rootRef = context.register(rootDict);

  for (let i = 0; i < itemDicts.length; i++) {
    itemDicts[i].set(PDFName.of('Parent'), rootRef);
    if (i > 0) itemDicts[i].set(PDFName.of('Prev'), itemRefs[i - 1]);
    if (i < itemDicts.length - 1) itemDicts[i].set(PDFName.of('Next'), itemRefs[i + 1]);
  }

  catalog.set(PDFName.of('Outlines'), rootRef);
  catalog.set(PDFName.of('PageMode'), PDFName.of('UseOutlines'));
}

export const HIGHLIGHT_RGB = [1, 0.9, 0]; // 노란색 (R G B, 0~1)

// 뷰어가 외형(AP)을 자동 생성하지 않아도 보이도록 Appearance Stream을 직접 생성
function buildHighlightAppearance(context, quad, rect) {
  const n = v => v.toFixed(2);
  const [tl, tr, bl, br] = quad;
  const ops = [
    '/GS gs',
    `${HIGHLIGHT_RGB.join(' ')} rg`,
    `${n(tl.x)} ${n(tl.y)} m`,
    `${n(tr.x)} ${n(tr.y)} l`,
    `${n(br.x)} ${n(br.y)} l`,
    `${n(bl.x)} ${n(bl.y)} l`,
    'f',
  ].join('\n');

  const gsRef = context.register(context.obj({
    Type: 'ExtGState',
    BM: PDFName.of('Multiply'), // 아래 글자가 비쳐 보이도록
    ca: 1,
    CA: 1,
  }));

  return context.register(context.flateStream(ops, {
    Type: 'XObject',
    Subtype: 'Form',
    BBox: context.obj(rect),
    Resources: context.obj({ ExtGState: context.obj({ GS: gsRef }) }),
  }));
}

// pdf-lib으로 Highlight 어노테이션 추가
// quad: [좌상, 우상, 좌하, 우하] (PDF 스펙 QuadPoints 순서)
export function addHighlightAnnotation(pdfDoc, page, quad) {
  const { context } = pdfDoc;

  const xs = quad.map(p => p.x);
  const ys = quad.map(p => p.y);
  const rect = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];

  const annot = context.obj({
    Type: PDFName.of('Annot'),
    Subtype: PDFName.of('Highlight'),
    Rect: context.obj(rect),
    QuadPoints: context.obj(quad.flatMap(p => [p.x, p.y])),
    C: context.obj(HIGHLIGHT_RGB),
    F: PDFNumber.of(4), // Print 플래그
    AP: context.obj({ N: buildHighlightAppearance(context, quad, rect) }),
  });

  const annotRef = context.register(annot);

  // /Annots 가 간접참조(예: /Annots 12 0 R)인 PDF도 있으므로 반드시 lookup 사용
  const annots = page.node.Annots();
  if (annots) {
    annots.push(annotRef);
  } else {
    page.node.set(PDFName.of('Annots'), context.obj([annotRef]));
  }
}
