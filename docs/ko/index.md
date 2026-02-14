---
layout: home

hero:
  name: SheetKit
  text: 고성능 SpreadsheetML 라이브러리
  tagline: 네이티브 속도로 Excel (.xlsx) 파일을 읽고 씁니다. Rust로 작성되었으며, TypeScript 바인딩을 함께 제공합니다.
  image: /logo.svg
  actions:
    - theme: brand
      text: 시작하기
      link: /ko/getting-started
    - theme: alt
      text: API 레퍼런스
      link: /ko/api-reference/
    - theme: alt
      text: GitHub
      link: https://github.com/Nebu1eto/sheetkit

features:
  - icon: ⚡
    title: 네이티브 성능
    details: Rust로 작성되어 높은 처리량과 낮은 메모리 사용량을 제공합니다. 대용량 스프레드시트 처리에 적합합니다.
  - icon: 📊
    title: 완전한 SpreadsheetML 지원
    details: 스타일, 차트, 이미지, 수식, 조건부 서식, 데이터 유효성 검사, 피벗 테이블, 스트리밍 쓰기 등 다양한 기능을 지원합니다.
  - icon: 🔒
    title: 타입 안전한 듀얼 런타임
    details: Rust와 TypeScript 모두에서 타입 안전한 API를 제공합니다. napi-rs 기반의 Node.js 바인딩으로 동일한 라이브러리를 네이티브 또는 JavaScript에서 사용할 수 있습니다.
---
