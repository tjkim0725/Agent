# Tosca Commander MCP 입문 교육자료

Tosca를 많이 써 보지 않았고 MCP도 처음인 테스터를 위한 3시간 20분짜리 교육 자료입니다.

| 파일 | 내용 |
|------|------|
| `Tosca_Commander_MCP_교육자료.pptx` | 강의 슬라이드 33장 (발표자 노트 포함) |
| `Tosca_Commander_MCP_실습워크북.docx` | 실습 워크북 (A4, Lab 0~5, 프롬프트 템플릿 T00~T18, 부록) |
| `교육설계안.md` | 대상 분석, 학습 목표, 시간표, 슬라이드·워크북 구성, 출처 |
| `검증보고서.md` | 제작 후 검증한 항목과 결과 |
| `src/` | 두 파일을 다시 만드는 스크립트와 공통 콘텐츠 |

## 기준

- Tosca Commander 2026.1 LTS Patch 1 이상 (MCP Server 애드인)
- Tricentis commander-mcp 스킬팩 1.2.1 (github.com/Tricentis/mcp-skills, `Tosca/Commander/MCP`)

## 고치는 방법

프롬프트, 템플릿, 도구 설명은 `src/content.js` 한 곳에만 있습니다. 여기를 고친 뒤 두 파일을 다시 만들면 슬라이드와 워크북의 문구가 같이 바뀝니다.

```bash
cd src
npm install
npm run build
python3 verify.py ../Tosca_Commander_MCP_교육자료.pptx ../Tosca_Commander_MCP_실습워크북.docx <mcp-skills>/Tosca/Commander/MCP/claude/skills/commander-mcp
```

슬라이드 색은 테마 색으로 들어가 있어서, PowerPoint의 디자인 → 테마 색에서 바꾸면 전체가 함께 바뀝니다.
