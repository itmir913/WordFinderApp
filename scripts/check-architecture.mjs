#!/usr/bin/env node
// CLAUDE.md의 아키텍처 규칙을 기계로 강제한다.
// 사람이 리뷰에서 놓치기 쉬운 것만 골랐다 — 규칙을 늘릴 거면 여기에 더한다.
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

const ROOT = fileURLToPath(new URL('..', import.meta.url));

// fs.globSync는 Node 22에서 아직 실험적이라 쓰지 않는다.
function filesUnder(dir, extensions) {
  return readdirSync(join(ROOT, dir), { recursive: true, withFileTypes: true })
    .filter(e => e.isFile() && extensions.some(ext => e.name.endsWith(ext)))
    .map(e => `${join(e.parentPath, e.name).slice(ROOT.length)}`.replaceAll('\\', '/'));
}

const RULES = [
  {
    name: 'invoke()는 스토어에서만 호출한다',
    why: 'Tauri IPC를 한 곳에 모아 두어야 권한·에러 처리를 한 번만 손본다.',
    dir: 'src',
    extensions: ['.vue', '.js'],
    skip: ['src/stores/app.js'],
    pattern: /(^|[^.\w])invoke\s*\(/,
  },
  {
    name: '인라인 style 속성 금지',
    why: 'Tailwind 유틸리티 클래스로만 스타일링한다.',
    dir: 'src',
    extensions: ['.vue'],
    pattern: /<[^>]*\sstyle\s*=\s*["']/,
  },
  {
    name: 'Vue 컴포넌트에서 워커를 직접 만들지 않는다',
    why: '워커 수명 관리는 스토어 책임이다.',
    dir: 'src/components',
    extensions: ['.vue'],
    pattern: /new\s+Worker\s*\(/,
  },
  {
    // 이건 스타일 문제가 아니다. 정규식 리터럴의 lookbehind는 파싱 시점
    // SyntaxError라, 지원하지 않는 엔진에서는 그 모듈이 통째로 로드에 실패한다.
    name: '정규식 lookbehind 금지 (?<= / ?<!)',
    why: 'Safari 16.4(macOS 13.3) 미만 WKWebView에서 모듈 전체가 SyntaxError로 죽는다.',
    dir: 'src',
    extensions: ['.vue', '.js'],
    pattern: /\(\?<[=!]/,
  },
  {
    name: '삼켜지는 에러 금지 (빈 catch)',
    why: 'CLAUDE.md — Silent error handling 금지.',
    dir: 'src',
    extensions: ['.vue', '.js'],
    pattern: /catch\s*(?:\([^)]*\))?\s*\{\s*\}/,
  },
];

let failed = 0;

for (const rule of RULES) {
  const files = filesUnder(rule.dir, rule.extensions)
    .filter(f => !(rule.skip ?? []).includes(f));

  // 경로가 바뀌어 규칙이 조용히 무력화되는 것을 막는다.
  if (files.length === 0) {
    console.error(`✖ 규칙 "${rule.name}"이 검사할 파일을 하나도 못 찾았다 (${rule.dir}).`);
    failed++;
    continue;
  }

  for (const file of files) {
    readFileSync(join(ROOT, file), 'utf8').split(/\r?\n/).forEach((line, i) => {
      if (rule.pattern.test(line)) {
        console.error(`✖ ${file}:${i + 1}  ${rule.name}`);
        console.error(`    ${line.trim()}`);
        console.error(`    ↳ ${rule.why}`);
        failed++;
      }
    });
  }
}

if (failed > 0) {
  console.error(`\n아키텍처 규칙 위반 ${failed}건.`);
  process.exit(1);
}
console.log(`✔ 아키텍처 규칙 ${RULES.length}개 통과 (검사 파일 ${filesUnder('src', ['.vue', '.js']).length}개)`);
