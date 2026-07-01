# GLOBAL RULES

## ARCHITECTURE
- Components must NOT call `invoke()` directly — use store actions only
- All Tauri IPC calls live in `src/stores/app.js`
- File I/O (read/write) handled by Rust commands in `src-tauri/src/lib.rs`
- Heavy processing (PDF parsing, Excel parsing) runs in `src/workers/processor.worker.js`

## TECH STACK
- Frontend: Vue 3 (Composition API) + Tailwind CSS v4 + Pinia
- Desktop: Tauri 2
- PDF: pdfjs-dist (Web Worker)
- Excel: SheetJS/xlsx (Web Worker)
- Icons: @lucide/vue
- Font: Pretendard

## CSS / STYLING
- **Tailwind CSS utility classes only** — no `style="..."` inline attributes
- No `<style>` blocks with hardcoded values (scoped Tailwind @apply is acceptable)
- Custom theme tokens defined in `src/assets/main.css` under `@theme`

## COMPONENT STRUCTURE
```
src/
  assets/main.css          # Tailwind v4 entry + @theme tokens
  stores/app.js            # Single Pinia store (state + all actions)
  workers/
    processor.worker.js    # PDF + Excel processing (Web Worker)
  components/
    TitleBar.vue
    CsvSection.vue
    ActionBar.vue
    tabs/
      GuideTab.vue
      FileTab.vue
      LogTab.vue
      DownloadTab.vue
  App.vue                  # Root layout only
  main.js
```

## CONVENTIONS
- Rust commands: snake_case
- Vue components: PascalCase
- Store actions: camelCase
- Files: kebab-case
- Commit messages: Korean or English, concise

## GIT / COMMIT RULES
- **Co-Authored-By / Co-Worked / Cowork 문구 삽입 금지**: 커밋 메시지에 Claude 관련 문구 일절 포함하지 않는다.
- 커밋 메시지: 한국어 또는 영어, 간결하게 작성

## PROHIBITED
- Inline CSS (`style="..."`)
- Business logic in Vue components
- Direct `invoke()` calls outside the store
- Silent error handling
