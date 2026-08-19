// default.csv와 프로그램 내장 단어 목록이 어긋나지 않는지 지킨다.
//
// 이 검사가 없어서 실제로 사고가 났다. `유의어 목록 업데이트`(647a756)가
// default.csv만 고치는 바람에 내장 목록이 57개 뒤처졌고, 아무도 몰랐다.
// Windows 포터블 zip은 exe 옆에 default.csv를 같이 넣어 주지만 macOS dmg는
// 그렇지 않아서, macOS 사용자만 조용히 옛 목록을 쓰고 있었다.
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { decodeCsvBytes, parseKeywords } from '../src/lib/csv.js';

const FILES = {
  'default.csv': '../default.csv',
  '내장 목록': '../src-tauri/src/resources/embedded_keywords.csv',
};

// 파일이 없으면 건너뛰지 않고 실패한다. 건너뛰게 만들면 경로가 바뀐 날
// 검사가 조용히 사라진다.
const keywordsOf = relative =>
  parseKeywords(decodeCsvBytes(readFileSync(fileURLToPath(new URL(relative, import.meta.url)))).text);

describe('단어 목록 동기화', () => {
  const lists = Object.fromEntries(
    Object.entries(FILES).map(([label, path]) => [label, keywordsOf(path)]),
  );

  it.each(Object.keys(FILES))('%s에 단어가 들어 있다', label => {
    expect(lists[label].length).toBeGreaterThan(100);
  });

  it('default.csv와 내장 목록의 단어가 완전히 같다', () => {
    const [a, b] = [lists['default.csv'], lists['내장 목록']];
    const onlyInDefault = a.filter(k => !b.includes(k));
    const onlyInEmbedded = b.filter(k => !a.includes(k));

    const hint = [
      '두 목록이 어긋났다. default.csv를 내장 목록으로 복사해 맞춘다:',
      '  cp default.csv src-tauri/src/resources/embedded_keywords.csv',
      onlyInDefault.length ? `내장 목록에 없는 단어(${onlyInDefault.length}): ${onlyInDefault.slice(0, 10).join(', ')}` : '',
      onlyInEmbedded.length ? `default.csv에 없는 단어(${onlyInEmbedded.length}): ${onlyInEmbedded.slice(0, 10).join(', ')}` : '',
    ].filter(Boolean).join('\n');

    expect(onlyInDefault, hint).toEqual([]);
    expect(onlyInEmbedded, hint).toEqual([]);
    expect(b, hint).toEqual(a); // 순서까지 같아야 한다
  });
});
