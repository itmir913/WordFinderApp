import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import {
  buildPattern,
  escapeRegex,
  findKeywordsInRow,
  hasConsecutiveSpaces,
  CONSEC_SPACE_LABEL,
} from '../src/workers/lib/keywords.js';

const matches = (text, keywords) => {
  const { pattern, kwMap } = buildPattern(keywords);
  return [...text.matchAll(pattern)].map(m => kwMap.get(m[0].toLowerCase()) ?? m[0]);
};

describe('escapeRegex', () => {
  it('정규식 메타문자를 리터럴로 만든다', () => {
    expect(new RegExp(escapeRegex('C++')).test('C++')).toBe(true);
    expect(new RegExp(escapeRegex('a.b')).test('axb')).toBe(false);
  });

  it('메타문자가 든 단어를 정상 탐지한다', () => {
    expect(matches('C++ 수업을 들었다', ['C++'])).toEqual(['C++']);
    expect(matches('(주)회사', ['(주)'])).toEqual(['(주)']);
  });
});

describe('buildPattern', () => {
  it('대소문자를 구분하지 않는다', () => {
    expect(matches('toeic TOEIC ToEiC', ['TOEIC'])).toEqual(['TOEIC', 'TOEIC', 'TOEIC']);
  });

  it('원본 표기를 kwMap으로 되돌린다', () => {
    const { kwMap } = buildPattern(['TOEIC']);
    expect(kwMap.get('toeic')).toBe('TOEIC');
  });

  // 회귀: 짧은 단어가 먼저 오면 긴 단어가 영원히 잘린 채로만 잡힌다.
  it('긴 단어를 짧은 단어보다 먼저 매칭한다 (등록 순서 무관)', () => {
    expect(matches('토익스피킹 응시', ['토익', '토익스피킹'])).toEqual(['토익스피킹']);
    expect(matches('토익스피킹 응시', ['토익스피킹', '토익'])).toEqual(['토익스피킹']);
  });

  it('한 단어가 여러 번 나오면 모두 잡는다', () => {
    expect(matches('토익 그리고 토익', ['토익'])).toHaveLength(2);
  });
});

describe('findKeywordsInRow', () => {
  const { pattern, kwMap } = buildPattern(['토익', 'TOEFL']);

  it('여러 셀에 걸친 발견 단어를 중복 없이 모은다', () => {
    // 한글/라틴 혼합 정렬 순서는 플랫폼 ICU에 따라 달라지므로 집합으로만 본다.
    const found = findKeywordsInRow(['토익 준비', '토익 재응시', 'TOEFL'], pattern, kwMap);
    expect([...found].sort()).toEqual(['TOEFL', '토익']);
  });

  it('발견 단어가 없으면 빈 배열', () => {
    expect(findKeywordsInRow(['평범한 내용'], pattern, kwMap)).toEqual([]);
  });

  it('숫자·null 셀도 문자열로 취급한다', () => {
    expect(() => findKeywordsInRow([1, null, undefined, ''], pattern, kwMap)).not.toThrow();
  });

  it('가나다순으로 정렬한다', () => {
    const { pattern: p, kwMap: k } = buildPattern(['하늘', '가방', '나무']);
    expect(findKeywordsInRow(['하늘 나무 가방'], p, k)).toEqual(['가방', '나무', '하늘']);
  });

  describe('연속 공백 검사', () => {
    it('옵션이 꺼져 있으면 검사하지 않는다', () => {
      expect(findKeywordsInRow(['앞    뒤'], pattern, kwMap, false)).toEqual([]);
    });

    it('2~5칸 연속 공백을 잡는다', () => {
      expect(findKeywordsInRow(['앞  뒤'], pattern, kwMap, true)).toEqual([CONSEC_SPACE_LABEL]);
      expect(findKeywordsInRow(['앞     뒤'], pattern, kwMap, true)).toEqual([CONSEC_SPACE_LABEL]);
    });

    it('공백 1칸은 잡지 않는다', () => {
      expect(findKeywordsInRow(['앞 뒤'], pattern, kwMap, true)).toEqual([]);
    });

    it('6칸 이상은 잡지 않는다 (표 정렬용 여백)', () => {
      expect(findKeywordsInRow(['앞      뒤'], pattern, kwMap, true)).toEqual([]);
    });

    // 회귀: 예전 구현은 모듈 수준 /g 정규식에 test()를 써서 lastIndex가
    // 전진했다. 초기화를 빠뜨리면 두 번째 행부터 결과가 어긋난다.
    it('여러 행을 연속 검사해도 결과가 흔들리지 않는다', () => {
      for (let i = 0; i < 5; i++) {
        expect(findKeywordsInRow(['앞  뒤'], pattern, kwMap, true)).toEqual([CONSEC_SPACE_LABEL]);
      }
    });
  });
});

describe('hasConsecutiveSpaces', () => {
  it.each([
    ['앞 뒤', false, '1칸'],
    ['앞  뒤', true, '2칸'],
    ['앞   뒤', true, '3칸'],
    ['앞     뒤', true, '5칸'],
    ['앞      뒤', false, '6칸 — 표 정렬용 여백'],
    ['앞뒤', false, '공백 없음'],
    ['  앞뒤', true, '문장 맨 앞'],
    ['앞뒤  ', true, '문장 맨 뒤'],
    ['앞      뒤  중간', true, '6칸 뒤에 2칸이 또 있으면 잡는다'],
  ])('%s → %s (%s)', (text, expected) => {
    expect(hasConsecutiveSpaces(text)).toBe(expected);
  });

  it('문자열이 아닌 값도 받는다', () => {
    expect(hasConsecutiveSpaces(123)).toBe(false);
    expect(hasConsecutiveSpaces(null)).toBe(false);
    expect(hasConsecutiveSpaces(undefined)).toBe(false);
  });

  // 회귀: lookbehind는 Safari 16.4(macOS 13.3) 미만 WKWebView에서
  // 파싱 시점 SyntaxError다. 정규식 리터럴이라 그 엔진에서는 이 모듈이
  // 통째로 로드에 실패하고, PDF·Excel 처리가 전부 죽는다.
  it('워커 로직에 lookbehind 정규식이 없다', () => {
    for (const file of ['keywords.js', 'pdf-highlight.js']) {
      const source = readFileSync(
        fileURLToPath(new URL(`../src/workers/lib/${file}`, import.meta.url)),
        'utf8',
      );
      expect(source, `${file}에 lookbehind가 있다`).not.toMatch(/\(\?<[=!]/);
    }
  });
});
