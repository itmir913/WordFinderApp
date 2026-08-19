// 검색 단어 매칭 로직 (순수 함수 — Worker 밖에서도 테스트 가능)

export const CONSEC_SPACE_LABEL = '연속 공백';

// 공백이 2~5칸 연달아 있는가.
//
// 예전에는 lookbehind를 쓴 정규식 리터럴 하나로 판정했다. lookbehind는
// Safari 16.4(macOS 13.3) 이상에서만 파싱되고, 리터럴이라 지원하지 않는
// WKWebView에서는 이 모듈 자체가 SyntaxError로 로드에 실패한다 — 연속 공백
// 검사만이 아니라 PDF·Excel 처리가 통째로 죽는다.
// 지금은 / +/g로 최대 길이 공백 덩어리를 훑어 같은 판정을 한다.
//
// 6칸 이상은 표 정렬용 여백으로 보고 넘어간다 — 예전 동작 그대로다.
export function hasConsecutiveSpaces(text) {
  for (const run of String(text).matchAll(/ +/g)) {
    if (run[0].length >= 2 && run[0].length <= 5) return true;
  }
  return false;
}

export function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// 키워드 배열 → { pattern, kwMap }
// kwMap: lowercase(keyword) → 원본 keyword (북마크/발견단어 표시용)
//
// 긴 단어를 먼저 정렬하는 이유: 정규식 교대(|)는 앞쪽 우선이라
// ['토익', '토익스피킹'] 순서면 '토익스피킹'이 영원히 '토익'으로만 잡힌다.
export function buildPattern(keywords) {
  const sorted = [...keywords].sort((a, b) => b.length - a.length);
  const escaped = sorted.map(escapeRegex);
  const pattern = new RegExp(escaped.join('|'), 'gi');
  const kwMap = new Map(sorted.map(kw => [kw.toLowerCase(), kw]));
  return { pattern, kwMap };
}

// 셀 한 줄에서 발견된 원본 키워드를 가나다순으로 반환 (Excel 처리용)
export function findKeywordsInRow(cells, pattern, kwMap, detectConsecutiveSpaces = false) {
  const foundSet = new Set();

  // 셀 단위 독립 매칭
  for (const cell of cells) {
    for (const match of String(cell).matchAll(pattern)) {
      foundSet.add(kwMap.get(match[0].toLowerCase()) ?? match[0]);
    }
  }

  if (detectConsecutiveSpaces && cells.some(hasConsecutiveSpaces)) {
    foundSet.add(CONSEC_SPACE_LABEL);
  }

  return [...foundSet].sort((a, b) => a.localeCompare(b, 'ko'));
}
