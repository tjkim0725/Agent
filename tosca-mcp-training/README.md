# Tosca Commander MCP 입문 교육자료

Tosca를 많이 써 보지 않았고 MCP도 처음인 테스터를 위한 3시간 20분짜리 교육 자료입니다. AI 클라이언트는 VS Code + GitHub Copilot만 다룹니다. 실습(Lab) 부분은 내용을 채워 넣을 수 있도록 틀만 들어 있습니다.

| 파일 | 내용 |
|------|------|
| `Tosca_Commander_MCP_교육자료.pptx` | 강의 슬라이드 46장 (발표자 노트 포함, 제공받은 디자인 템플릿 적용) |
| `Tosca_Commander_MCP_실습워크북.docx` | 실습 워크북 (A4, Lab 1~6 틀, 프롬프트 템플릿 T00~T18, 부록) |
| `교육설계안.md` | 대상 분석, 학습 목표, 시간표, 슬라이드·워크북 구성, 출처 |
| `검증보고서.md` | 제작 후 검증한 항목과 결과 |
| `src/` | 두 파일을 다시 만드는 스크립트와 공통 콘텐츠 |

## 기준

- Tosca Commander 2026.1 LTS Patch 1 이상 (MCP Server 애드인)
- Tricentis commander-mcp 스킬팩 1.2.1 (github.com/Tricentis/mcp-skills, `Tosca/Commander/MCP`, VS Code용 팩)
- VS Code + GitHub Copilot (메뉴 이름은 2026년 10월 VS Code 문서 기준)

## 실습 내용 채우기

슬라이드 18~32와 워크북 2.5, 3장에 대괄호 `[ ]`로 된 칸이 있습니다. 두 가지 방법 중 하나로 채웁니다.

- 파일에서 바로 고치기: PowerPoint와 Word에서 대괄호 칸을 지우고 내용을 넣습니다. 슬라이드와 워크북을 따로 고쳐야 합니다.
- `src/content.js`에서 고친 뒤 다시 빌드하기: `labs` 배열의 `labFrame(n)`을 내용을 채운 객체로 바꿉니다(파일 안의 예시 참고). 항목은 `title`, `time`, `goal`, `before`, `steps`, `prompts`, `tools`, `checks`, `record`, `stuck`이고, 슬라이드와 워크북이 같이 바뀝니다. Lab 수는 `LAB_COUNT`, 휴식 위치는 `BREAK_AFTER`로 정합니다. 대상 앱과 모듈 칸(슬라이드 18, 워크북 2.5)은 각 빌드 스크립트에 직접 들어 있습니다.

## 고치는 방법

프롬프트, 템플릿, 도구 설명은 `src/content.js` 한 곳에만 있습니다. 여기를 고친 뒤 두 파일을 다시 만들면 슬라이드와 워크북의 문구가 같이 바뀝니다.

```bash
cd src
pip install -r requirements.txt   # python-pptx
npm install                       # docx
npm run build
python3 verify.py ../Tosca_Commander_MCP_교육자료.pptx ../Tosca_Commander_MCP_실습워크북.docx <mcp-skills>/Tosca/Commander/MCP/claude/skills/commander-mcp
```

## 슬라이드 디자인

슬라이드는 제공받은 디자인 템플릿(`01_Tosca_Customer_4H_DemoWebShop.pptx`)의 마스터·레이아웃·테마·이미지를 그대로 씁니다. 원본 슬라이드 내용은 쓰지 않았습니다.

- `src/template/tricentis_base.pptx`: 원본 템플릿에서 슬라이드를 모두 빼고 디자인 요소만 남긴 파일 (`src/make_base.py`로 생성)
- `src/build_pptx.py`: 이 파일 위에 슬라이드를 만들며, 제목·3단 규칙선·표·번호 사각형·LAB 화면·각주·꼬리말은 원본 템플릿의 위치·크기·색을 따릅니다

템플릿이 바뀌면 `python3 make_base.py <새 템플릿.pptx> template/tricentis_base.pptx`로 기반 파일을 다시 만든 뒤 빌드합니다.
