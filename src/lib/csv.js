// 검색 단어 CSV 읽기 (순수 함수 — 스토어 밖에서도 테스트 가능)

// 한국 윈도우의 Excel은 CSV를 CP949(EUC-KR)로 저장한다. UTF-8로만 읽으면
// 단어가 통째로 깨진 채 "로드 완료"가 뜨고 아무것도 안 잡힌다.
//
// fatal:true가 핵심이다. 이게 없으면 TextDecoder가 깨진 바이트를 U+FFFD로
// 조용히 바꿔치기해서 폴백이 영원히 안 걸린다.
const ENCODINGS = ['utf-8', 'euc-kr'];

export function decodeCsvBytes(bytes) {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);

  for (const encoding of ENCODINGS) {
    try {
      return { text: new TextDecoder(encoding, { fatal: true }).decode(data), encoding };
    } catch {
      // 이 인코딩으로는 못 읽는다 — 다음 후보로 넘어간다.
    }
  }

  // 어느 쪽으로도 못 읽으면 UTF-8로 최대한 복구해서라도 넘긴다.
  // 여기까지 왔다는 것 자체를 호출자가 로그로 남길 수 있게 encoding을 알려 준다.
  return { text: new TextDecoder('utf-8').decode(data), encoding: 'unknown' };
}

// 첫 줄(헤더)을 버리고 각 줄의 첫 칸을 단어로 쓴다. 중복·빈 줄은 제거.
export function parseKeywords(text) {
  const lines = text.replace(/^﻿/, '').split(/\r?\n/).slice(1);
  return [...new Set(lines.map(line => line.split(',')[0].trim()).filter(Boolean))];
}
