// PPT와 워크북이 함께 쓰는 교육 콘텐츠.
// 프롬프트·템플릿 문구는 여기서만 고친다. 두 산출물의 문구가 어긋나지 않게 하기 위함.

const MCP_URL = "http://127.0.0.1:46248/mcp";

const meta = {
  course: "Tosca Commander MCP 입문",
  subtitle: "AI 어시스턴트로 테스트케이스 만들기",
  baseline: "Tosca Commander 2026.1 LTS Patch 1 이상 · commander-mcp 스킬팩 1.2.1 기준",
  duration: "3시간 20분 (휴식 10분 포함)",
  mcpUrl: MCP_URL,
  app: "https://demowebshop.tricentis.com",
};

// 시간표 (슬라이드 2, 워크북 표지 뒤)
const agenda = [
  { no: "00", title: "오프닝과 Tosca 용어", time: "10분" },
  { no: "01", title: "Tosca MCP 이해", time: "25분" },
  { no: "02", title: "실습: 테스트케이스 만들기", time: "90분" },
  { no: "03", title: "활용사례와 프롬프트 작성법", time: "30분" },
  { no: "04", title: "프롬프트 템플릿", time: "25분" },
  { no: "05", title: "정리와 Q&A", time: "10분" },
];

const terms = [
  { term: "Workspace", ko: "워크스페이스", desc: "Tosca 작업 공간. 내 PC에만 있는 싱글 유저와, 공통 리포지토리를 여럿이 함께 쓰는 멀티 유저가 있다." },
  { term: "Module", ko: "모듈", desc: "테스트할 화면의 기술 정보. XScan으로 화면을 스캔해서 만든다. 화면 요소 하나하나는 ModuleAttribute라고 부른다." },
  { term: "TestCase", ko: "테스트케이스", desc: "모듈을 순서대로 끌어다 놓아 만든 테스트 절차." },
  { term: "TestStep", ko: "테스트스텝", desc: "테스트케이스의 한 단계. 스텝 안에 넣는 값은 TestStepValue." },
  { term: "ActionMode", ko: "액션모드", desc: "값을 어떻게 쓸지 정한다. Input(입력), Verify(검증), Buffer(저장 후 재사용), WaitOn(나타날 때까지 대기) 등." },
  { term: "ScratchBook", ko: "스크래치북", desc: "테스트케이스를 시험 삼아 돌려 보는 곳. 결과가 정식 실행 기록으로 남지 않는다." },
  { term: "ExecutionList", ko: "실행 목록", desc: "정식 실행 단위. 실행할 때마다 결과(ExecutionLog)가 남는다." },
  { term: "Check out / Check in", ko: "체크아웃 / 체크인", desc: "멀티 유저에서 객체를 잠그고(체크아웃) 수정한 뒤 공통 리포지토리에 반영(체크인)하는 절차." },
];

// 도구 32종 (commander-mcp 1.2.1 tools-catalog 기준)
const toolGroups = [
  {
    key: "read", name: "조회", icon: "search",
    summary: "워크스페이스 구조, 객체, 속성을 읽는다",
    tools: [
      ["get_workspace_info", "열린 워크스페이스 정보(이름, 경로, 싱글/멀티 유저, 프로젝트 루트 ID)"],
      ["get_object_info", "경로나 ID로 객체를 찾고, 필요하면 바로 아래 자식 목록까지 조회"],
      ["list_assets", "TestCase, XModule, TCFolder 같은 타입으로 하위 폴더까지 검색"],
      ["get_attributes", "객체의 속성 값과 수정 가능 여부 조회"],
      ["get_current_selection", "Commander 화면에서 지금 선택된 객체 조회 (사용자가 '선택한 것'을 말할 때만)"],
      ["get_automation_object", "모듈(XModule) 내용을 직렬화된 형태로 가져오기"],
    ],
  },
  {
    key: "edit", name: "편집", icon: "edit",
    summary: "속성 값과 설정을 바꾸고 객체를 정리한다",
    tools: [
      ["set_attribute", "속성 하나 수정 (이름, 설명, 값, ActionMode 등)"],
      ["list_test_configuration_parameters", "TCP(Test Configuration Parameter) 조회"],
      ["create_test_configuration_parameters", "TCP 설정 (해당 객체의 로컬 TCP 전체를 교체)"],
      ["set_object_selection", "Commander 화면에서 객체를 선택해 사용자에게 보여주기"],
      ["delete_objects", "객체 삭제. 되돌릴 수 없음"],
    ],
  },
  {
    key: "author", name: "작성·작업", icon: "build",
    summary: "테스트케이스를 만들고 메뉴 작업을 실행한다",
    tools: [
      ["create_test_cases", "테스트케이스 생성 또는 수정 (스텝, 요구사항 연결 포함)"],
      ["execute_drop_task", "드래그 앤 드롭 동작 (모듈을 테스트케이스에 끌어다 놓기 등)"],
      ["import_automation_objects", "직렬화된 자동화 객체를 워크스페이스로 가져오기"],
      ["list_available_tasks", "오른쪽 클릭 메뉴에 있는 작업 목록 조회"],
      ["execute_task", "오른쪽 클릭 메뉴 작업 실행 (Checkout 등)"],
    ],
  },
  {
    key: "run", name: "실행·저장", icon: "play",
    summary: "검증 실행을 돌리고 변경 내용을 저장한다",
    tools: [
      ["execute_test_suite", "ScratchBook 검증 실행 또는 ExecutionList 실행 시작"],
      ["execute_test_suite_status", "실행 상태 확인 (끝날 때까지 반복 조회)"],
      ["save_workspace", "변경 내용 저장"],
      ["check_in_all", "체크아웃한 객체를 모두 체크인 (멀티 유저)"],
      ["update_all", "공통 리포지토리의 최신 내용 받아오기 (멀티 유저)"],
    ],
  },
  {
    key: "api", name: "API", icon: "api",
    summary: "API 연결과 요청/응답 모듈을 만든다",
    tools: [
      ["list_api_connections", "API 연결 목록 조회 (비밀번호, 토큰은 반환하지 않음)"],
      ["create_api_connection", "재사용할 API 연결 생성"],
      ["create_api_module", "API 요청/응답 모듈 생성. 실제 엔드포인트를 호출해 응답까지 받아 둔다"],
      ["update_api_module", "기존 API 모듈의 헤더, 파라미터 등 수정"],
    ],
  },
  {
    key: "di", name: "Data Integrity", icon: "db",
    summary: "DB 연결과 데이터 비교 테스트를 만든다 (라이선스 필요)",
    tools: [
      ["di_connection", "DI 연결 생성, 수정, 삭제, 조회"],
      ["get_di_connection_schema", "연결된 DB의 테이블·컬럼 정보 조회 (Dev/Staging 태그 연결만)"],
      ["run_di_sql_statement", "DI 연결로 SQL 실행 후 결과 조회"],
      ["create_di_db_expert_module", "SQL 결과를 검증하는 DB Expert 테스트케이스 생성"],
      ["create_di_row_by_row_comparison", "원본과 대상 데이터를 행 단위로 비교하는 테스트케이스 생성"],
      ["create_di_load_caching_db_from_customization", "SAP 테이블을 캐싱 DB로 적재하는 테스트케이스 생성"],
      ["parse_di_lineage_mapping", "데이터 거버넌스 도구에서 내보낸 매핑 CSV 해석"],
    ],
  },
];

// 클라이언트별 연결 설정
const clients = [
  {
    name: "VS Code + GitHub Copilot",
    where: "프로젝트 폴더의 .vscode/mcp.json",
    note: "Copilot Chat을 Agent 모드로 바꿔야 도구가 보인다.",
    code: `{
  "servers": {
    "tosca-commander": {
      "type": "http",
      "url": "${MCP_URL}"
    }
  }
}`,
  },
  {
    name: "Cursor",
    where: "Settings → MCP → Add server (또는 .cursor/mcp.json)",
    note: "서버 이름은 tosca-commander로 맞추면 스킬팩 안내와 같아진다.",
    code: `{
  "mcpServers": {
    "tosca-commander": {
      "url": "${MCP_URL}"
    }
  }
}`,
  },
  {
    name: "Claude Code",
    where: "터미널에서 한 번 실행",
    note: "등록 후 /mcp 명령으로 연결 상태를 볼 수 있다.",
    code: `claude mcp add --transport http tosca-commander ${MCP_URL}`,
  },
  {
    name: "Claude Desktop",
    where: "설정 → 개발자 → 구성 편집 (claude_desktop_config.json)",
    note: "로컬 HTTP 서버를 mcp-remote로 중계한다. Node.js가 설치돼 있어야 한다.",
    code: `{
  "mcpServers": {
    "tosca-commander": {
      "command": "npx",
      "args": ["-y", "mcp-remote", "${MCP_URL}"]
    }
  }
}`,
  },
];

// 실습용으로 강사가 미리 준비하는 모듈
const practiceModules = [
  { module: "OpenUrl", where: "표준 모듈", attrs: "Url", note: "브라우저를 열고 주소로 이동" },
  { module: "Header", where: "Modules/MCP실습", attrs: "Log in, Log out, Account", note: "상단 메뉴. Account는 로그인 후 이메일이 보이는 링크" },
  { module: "Login", where: "Modules/MCP실습", attrs: "Email, Password, Log in, Error Message", note: "로그인 화면" },
  { module: "CloseBrowser", where: "표준 모듈", attrs: "Title", note: "브라우저 닫기" },
];

// Lab 3에서 만들어질 스텝 (기대 결과)
const lab3Steps = [
  ["1", "OpenUrl", "Url", "https://demowebshop.tricentis.com", "Input"],
  ["2", "Header", "Log in", "{CLICK}", "Input"],
  ["3", "Login", "Email", "[실습 이메일]", "Input"],
  ["", "Login", "Password", "[실습 비밀번호]", "Input"],
  ["", "Login", "Log in", "{CLICK}", "Input"],
  ["4", "Header", "Account", "[실습 이메일]", "Verify"],
  ["5", "Header", "Log out", "{CLICK}", "Input"],
  ["6", "CloseBrowser", "Title", "Demo Web Shop*", "Input"],
];

const labs = [
  {
    id: "Lab 0", title: "연결 확인", time: "10분",
    goal: "AI 어시스턴트가 Commander의 MCP Server에 붙었는지 확인한다.",
    before: [
      "Tosca Commander를 실행하고 실습 워크스페이스를 연다.",
      "AI 클라이언트에 tosca-commander 서버를 등록한다 (2장 참고).",
      "VS Code라면 Copilot Chat을 Agent 모드로 바꾸고, 도구 목록에 tosca-commander가 켜져 있는지 본다.",
    ],
    prompts: [
      { label: "0-1", text: "Tosca Commander MCP에서 쓸 수 있는 도구 목록을 보여줘. 도구 이름과 한 줄 설명을 표로 정리해줘." },
      { label: "0-2", text: "지금 열려 있는 Tosca 워크스페이스 정보를 알려줘. 이름, 경로, 싱글 유저인지 멀티 유저인지 포함해서." },
    ],
    tools: ["get_workspace_info"],
    checks: [
      "도구 목록에 create_test_cases와 execute_test_suite가 있다",
      "워크스페이스 이름이 Commander 제목 표시줄과 같다",
      "모드(single / multi)를 아래 기록란에 적었다",
    ],
    record: ["워크스페이스 이름", "모드 (single / multi)", "도구 개수"],
    stuck: [
      ["연결 거부, 응답 없음", "Commander가 켜져 있고 워크스페이스가 열려 있는지, 주소의 포트가 46248인지 확인한다. Patch 1이면 Tosca MCP Server Addin이 켜져 있는지도 본다."],
      ["도구가 하나도 안 보임", "클라이언트에서 서버를 켰는지, VS Code라면 Agent 모드인지 확인한다. 설정을 바꿨다면 클라이언트를 다시 시작한다."],
    ],
  },
  {
    id: "Lab 1", title: "워크스페이스 둘러보기", time: "10분",
    goal: "조회만 하는 요청으로 워크스페이스 구조와 실습 모듈을 확인한다.",
    before: ["이 Lab은 아무것도 바꾸지 않는다. 결과에 나오는 긴 ID(surrogateId)는 몰라도 된다."],
    prompts: [
      { label: "1-1", text: "TestCases 폴더 아래 구조를 2단계 깊이까지 트리로 보여줘. 각 항목은 이름과 타입만 표시해줘." },
      { label: "1-2", text: "Modules/MCP실습 폴더에 있는 모듈(XModule)을 모두 찾아서, 모듈별 속성(ModuleAttribute) 이름을 표로 정리해줘." },
    ],
    tools: ["get_object_info", "list_assets", "get_attributes"],
    checks: [
      "Header, Login 모듈이 목록에 있다",
      "Login 모듈에 Email, Password, Log in 속성이 있다",
      "Commander에서 아무것도 바뀌지 않았다",
    ],
    record: ["TestCases 아래 1단계 폴더 수", "Login 모듈 속성 개수"],
    stuck: [
      ["MCP실습 폴더가 없다고 함", "Commander 트리에서 실제 폴더 이름을 확인하고, 보이는 그대로 경로를 다시 적는다. 띄어쓰기도 같아야 한다."],
    ],
  },
  {
    id: "Lab 2", title: "수동 테스트케이스 만들기", time: "15분",
    goal: "글로 된 테스트 절차를 Tosca 수동 테스트케이스로 옮긴다.",
    before: ["[내 이름]은 본인 이름으로 바꾼다. 폴더가 없으면 AI가 중간 폴더를 만든다."],
    prompts: [
      {
        label: "2-1",
        text: `TestCases/MCP실습/[내 이름] 폴더에 '로그인 실패_수동' 수동 테스트케이스를 만들어줘.
설명: 틀린 비밀번호로 로그인하면 오류 문구가 나오는지 확인한다.
스텝 (지시 / 기대 결과):
1. 브라우저에서 https://demowebshop.tricentis.com 에 접속한다 / 메인 화면이 열린다
2. 상단의 Log in 링크를 클릭한다 / 로그인 화면이 열린다
3. 실습 계정 이메일과 틀린 비밀번호를 넣고 Log in 버튼을 클릭한다 / 'Login was unsuccessful' 문구가 보인다
작업 상태(TestCaseWorkState)는 IN_WORK로 해줘.
만들기 전에 어떤 순서로 작업할지 먼저 보여줘.`,
      },
      { label: "2-2", text: "좋아, 그 순서대로 진행해줘. 끝나면 만든 테스트케이스의 스텝을 다시 읽어서 보여줘." },
    ],
    tools: ["create_test_cases", "list_available_tasks", "execute_task (Manual XTestStep)", "set_attribute", "save_workspace"],
    checks: [
      "작업 순서를 먼저 받고 나서 진행을 승인했다",
      "경로가 TestCases/MCP실습/[내 이름]/로그인 실패_수동 이다",
      "스텝 3개에 지시와 기대 결과가 모두 들어갔다",
      "Commander에서 테스트케이스를 직접 열어 확인했다",
    ],
    record: ["만든 테스트케이스 경로", "스텝 수"],
    stuck: [
      ["기대 결과가 빠짐", "\"3개 스텝 모두 기대 결과를 넣었는지 확인하고, 빠진 것만 채워줘\"라고 다시 요청한다."],
      ["Commander 화면에 대화상자가 뜨고 멈춤", "대화상자를 닫고 \"이름과 설명은 속성 변경(set_attribute)으로 해줘\"라고 지시한다. 스킬팩을 설치하면 대부분 예방된다."],
    ],
  },
  {
    id: "Lab 3", title: "자동 테스트케이스 만들기", time: "25분",
    goal: "기존 모듈을 재사용해 로그인 자동화 테스트케이스를 만든다. 계획을 먼저 받고 승인하는 흐름을 익힌다.",
    before: [
      "[실습 이메일], [실습 비밀번호]는 강사가 나눠 준 계정으로 바꾼다.",
      "실무에서는 비밀번호를 평문 값으로 두지 않는다. 오늘은 실습이라 그대로 넣는다.",
    ],
    prompts: [
      {
        label: "3-1",
        text: `TestCases/MCP실습/[내 이름] 폴더에 '로그인 성공_자동' 자동화 테스트케이스를 만들려고 해. 아직 만들지 말고 계획만 보여줘.
- 모듈은 Modules/MCP실습 폴더와 표준 모듈(OpenUrl, CloseBrowser)에서만 찾아. 새로 스캔하지 마.
- 시나리오:
  1. https://demowebshop.tricentis.com 열기
  2. Header 모듈의 Log in 클릭
  3. Login 모듈에 Email [실습 이메일], Password [실습 비밀번호] 입력 후 Log in 클릭
  4. Header 모듈의 Account 값이 [실습 이메일]과 같은지 Verify
  5. Log out 클릭 후 브라우저 닫기
- 스텝별로 모듈, 속성, 값, ActionMode를 표로 보여주고, 맞는 모듈을 못 찾은 스텝은 따로 표시해줘.`,
      },
      { label: "3-2", text: "계획대로 만들어줘. 다 만들면 테스트케이스의 스텝 구조를 다시 읽어서 표로 보여줘." },
    ],
    tools: ["list_assets", "get_object_info", "create_test_cases", "set_attribute"],
    checks: [
      "계획 표를 먼저 받았고, 모든 스텝에 모듈이 매칭됐다",
      "4번 스텝(Account)의 ActionMode가 Verify다",
      "결과 표가 계획 표(기대 스텝)와 같다",
      "Commander에서 스텝과 값을 직접 확인했다",
    ],
    record: ["만든 테스트케이스 경로", "계획과 달라진 부분"],
    stuck: [
      ["모듈을 못 찾았다고 함", "모듈 폴더 경로를 다시 확인해 준다. 정말 없는 모듈이면 AI는 만들 수 없다. 사람이 XScan으로 스캔한 뒤 다시 요청한다."],
      ["값이 엉뚱한 속성에 들어감", "\"3번 스텝의 Email 값만 [실습 이메일]로 고쳐줘\"처럼 스텝 번호와 속성을 짚어서 고친다."],
    ],
  },
  {
    id: "Lab 4", title: "ScratchBook 검증 실행과 수정", time: "20분",
    goal: "만든 테스트케이스를 검증 실행하고, 일부러 낸 실패를 AI와 함께 분석해서 고친다.",
    before: [
      "실행이 시작되면 브라우저가 열린다. 끝날 때까지 마우스와 키보드를 건드리지 않는다.",
      "실행은 한 번에 하나만 된다. 앞 실행이 끝나기 전에 다음 요청을 보내지 않는다.",
    ],
    prompts: [
      { label: "4-1", text: "'로그인 성공_자동' 테스트케이스를 ScratchBook에서 검증 실행해줘. 끝나면 스텝별 통과/실패를 표로 보여줘." },
      { label: "4-2", text: "같은 테스트케이스에서 Account의 Verify 값을 wrong@example.com 으로 바꾸고 다시 검증 실행해줘." },
      { label: "4-3", text: "실패한 스텝과 메시지를 알려줘. 원인이 테스트 데이터 문제인지 애플리케이션 문제인지 근거와 함께 설명해줘. 아직 고치지는 마." },
      { label: "4-4", text: "Verify 값을 원래 [실습 이메일]로 되돌리고 다시 검증 실행해줘. 통과하면 결과 표를 보여줘." },
    ],
    tools: ["save_workspace", "execute_test_suite (is_validation_run=true)", "execute_test_suite_status", "set_attribute"],
    checks: [
      "4-1 실행이 모두 통과했다",
      "4-2 실행에서 Account 검증 스텝이 실패했다",
      "AI가 원인을 기대값(테스트 데이터) 문제로 설명했다",
      "4-4 실행이 다시 통과했다",
    ],
    record: ["1차 실행 결과", "2차 실행에서 실패한 스텝", "3차 실행 결과"],
    stuck: [
      ["\"another run in progress\" 같은 메시지", "앞 실행이 아직 안 끝났다. 기다렸다가 \"이전 실행 결과를 먼저 보여줘\"라고 요청한다."],
      ["응답이 끊기거나 시간 초과", "다시 실행하지 않는다. Commander 쪽 실행은 계속되고 있을 수 있다. 잠시 뒤 결과만 다시 물어본다."],
    ],
  },
  {
    id: "Lab 5", title: "저장과 정리", time: "10분",
    goal: "작업 결과를 저장하고, 멀티 유저 환경에서는 무엇이 더 필요한지 확인한다.",
    before: ["오늘 실습 워크스페이스는 싱글 유저다. 5-2는 멀티 유저일 때 쓰는 문장으로 읽기만 한다."],
    prompts: [
      { label: "5-1", text: "오늘 만들거나 바꾼 객체를 경로와 함께 표로 정리하고, 워크스페이스를 저장해줘." },
      { label: "5-2", text: "(멀티 유저일 때) 바뀐 객체 목록을 먼저 보여줘. 내가 확인하면 저장하고 체크인해줘. 체크인 코멘트는 'MCP 실습: 로그인 테스트 추가'로 남겨줘." },
    ],
    tools: ["save_workspace", "check_in_all (멀티 유저)", "update_all (멀티 유저)"],
    checks: [
      "저장 완료 응답을 받았다",
      "정리된 표에 Lab 2, Lab 3에서 만든 테스트케이스가 있다",
      "Commander를 닫았다가 다시 열어도 테스트케이스가 남아 있다",
    ],
    record: ["저장한 시각", "정리된 객체 수"],
    stuck: [
      ["check_in_all이 오류를 냄", "싱글 유저 워크스페이스에서는 체크인이 없다. 저장(save_workspace)만 하면 된다."],
    ],
  },
];

// AI에게 맡기지 않는 일 (슬라이드 11, 워크북 1.4)
const notForAI = [
  ["LuScanLine", "화면 스캔(XScan)", "AI는 이미 있는 모듈만 재사용한다. 새 화면은 사람이 XScan으로 스캔한다."],
  ["LuKeyRound", "비밀번호와 접속 정보", "민감한 값은 사람이 넣는다. AI가 만든 DI 연결은 'Allow AI queries'가 꺼진 채로 생긴다."],
  ["LuTrash2", "삭제 결정", "삭제는 되돌릴 수 없다. 대상 목록을 먼저 보고 따로 승인한다."],
  ["LuBug", "결함 판단", "실패가 앱 결함인지는 사람이 판단한다. 테스트를 고쳐서 결함을 덮지 않는다."],
  ["LuMonitorOff", "Commander 밖의 일", "Commander가 닫혀 있으면 동작하지 않는다. 야간 배치나 CI는 TCShell 같은 방식을 쓴다."],
  ["LuCirclePause", "동시 실행", "실행은 워크스페이스 전체에서 한 번에 하나만 된다."],
];

// 문제 해결 (부록 B, 슬라이드 22)
const troubleshooting = [
  ["연결 거부, 서버 응답 없음", "Commander가 꺼져 있거나 워크스페이스가 닫혀 있음", "Commander 실행 → 워크스페이스 열기 → 포트 46248 확인 (디버그 빌드는 8080, 직접 바꿨다면 그 값)"],
  ["도구 목록이 비어 있음", "클라이언트에서 서버가 꺼져 있음", "서버 활성화, VS Code는 Agent 모드, 설정 저장 후 클라이언트 재시작"],
  ["\"Unknown task ...\" 응답", "대상 또는 상위 폴더가 체크아웃되지 않음 (멀티 유저)", "\"상위 폴더를 Checkout 한 다음 다시 시도해줘\""],
  ["모듈을 못 찾음", "모듈 위치를 모르거나 실제로 모듈이 없음", "모듈 폴더 경로를 알려 준다. 없으면 XScan으로 스캔 후 재요청"],
  ["Commander가 대화상자를 띄우고 멈춤", "이름·설명 변경을 메뉴 작업으로 실행함", "대화상자 닫기 → \"속성 변경(set_attribute)으로 해줘\". 스킬팩 설치로 예방"],
  ["실행 결과가 안 돌아옴, 시간 초과", "실행이 길어 클라이언트 요청 시간이 지남", "재실행하지 말고 \"마지막 실행 결과를 확인해줘\""],
  ["다른 실행이 진행 중이라는 메시지", "실행은 워크스페이스 전체에서 한 번에 하나", "앞 실행이 끝난 뒤 요청"],
  ["바꾼 내용이 사라짐", "저장하지 않고 Commander를 닫음", "작업 끝에 항상 저장 요청 (save_workspace)"],
];

// Part 3 활용사례
const useCases = [
  {
    no: "1", title: "수동 테스트 문서 옮기기", icon: "doc",
    body: "엑셀이나 워드로 관리하던 수동 테스트 절차를 Tosca 수동 테스트케이스로 등록한다.",
    human: "스텝 누락, 기대 결과 문구",
    template: "T05",
  },
  {
    no: "2", title: "자동화 초안 만들기", icon: "build",
    body: "시나리오 문장과 기존 모듈로 자동 테스트케이스 초안을 만들고 ScratchBook에서 바로 돌려 본다.",
    human: "모듈 매칭, ActionMode, 테스트 데이터",
    template: "T06",
  },
  {
    no: "3", title: "테스트케이스 읽어 주기", icon: "read",
    body: "다른 사람이 만든 테스트케이스의 목적, 스텝, 데이터를 문장으로 요약해 인수인계 자료로 쓴다.",
    human: "업무 의미가 맞는지",
    template: "T07",
  },
  {
    no: "4", title: "실패 원인 분석", icon: "search",
    body: "실패한 실행을 읽어 원인을 분류하고, 두 실행을 비교하거나 자주 흔들리는 테스트를 찾는다.",
    human: "결함 등록 여부, 확신도 낮은 항목",
    template: "T10 · T12 · T13",
  },
  {
    no: "5", title: "반복 편집", icon: "edit",
    body: "여러 테스트케이스의 상태값 변경, TCP 설정, 스텝 추가처럼 손이 많이 가는 편집을 맡긴다.",
    human: "대상 범위, 체크인 전 변경 목록",
    template: "T03 · T04",
  },
  {
    no: "6", title: "API·데이터 테스트", icon: "db",
    body: "API 요청 모듈을 실제 응답까지 받아 만들고, Data Integrity로 DB 간 데이터 비교 테스트를 만든다.",
    human: "접속 정보와 비밀번호는 직접 입력",
    template: "T17",
  },
];

const promptElements = [
  { key: "대상", q: "어디서?", desc: "폴더 경로, 테스트케이스 이름, 모듈 위치를 Commander에 보이는 그대로 적는다.", ex: "TestCases/MCP실습/홍길동 폴더에" },
  { key: "작업", q: "무엇을?", desc: "만들기, 고치기, 실행, 분석 중 하나. 한 번에 한 가지만 시킨다.", ex: "'로그인 성공_자동' 테스트케이스를 만들어줘" },
  { key: "조건", q: "어떻게?", desc: "쓸 모듈 범위, 값, ActionMode, 하지 말아야 할 것을 적는다.", ex: "모듈은 Modules/MCP실습에서만 찾고, 새로 스캔하지 마" },
  { key: "확인", q: "어떻게 받을까?", desc: "계획 먼저, 표로, 승인 후 실행처럼 결과 받는 방식을 정한다.", ex: "만들기 전에 스텝 계획을 표로 보여줘" },
];

const promptPairs = [
  {
    bad: "로그인 테스트 만들어줘",
    good: "TestCases/MCP실습/홍길동 폴더에 '로그인 성공_자동' 테스트케이스를 만들어줘. 모듈은 Modules/MCP실습에서만 찾고, 만들기 전에 스텝 계획을 표로 보여줘.",
    why: "위치와 모듈 범위가 없으면 AI가 아무 폴더에나 만들거나 엉뚱한 모듈을 고른다.",
  },
  {
    bad: "어제 실패한 거 고쳐줘",
    good: "Execution/야간 회귀 ExecutionList의 마지막 실행에서 실패한 테스트케이스를 찾아, 실패 스텝과 메시지를 표로 정리해줘. 원인 추정과 확신도도 적어 주고, 수정은 하지 마.",
    why: "분석과 수정을 한 번에 시키면 실제 결함을 테스트 수정으로 덮어 버릴 수 있다.",
  },
  {
    bad: "필요 없는 테스트케이스 정리해줘",
    good: "TestCases/임시 폴더에서 이름이 '복사본'으로 끝나는 테스트케이스 목록만 보여줘. 삭제는 목록을 확인한 다음 따로 요청할게.",
    why: "삭제는 되돌릴 수 없다. 목록 확인과 삭제를 나눠서 요청한다.",
  },
];

const tips = [
  ["경로는 보이는 그대로", "TestCases/MCP실습/홍길동처럼 Commander 트리에 보이는 이름 그대로 적는다."],
  ["모듈 찾을 곳을 정해 준다", "\"Modules/MCP실습에서만\"처럼 범위를 주면 엉뚱한 모듈을 쓰지 않는다."],
  ["바꾸는 일은 계획부터", "\"아직 실행하지 말고 계획만 보여줘\"를 붙인다."],
  ["한 번에 한 가지", "만들기, 실행, 분석을 한 문장에 몰아넣지 않는다."],
  ["결과 모양을 정한다", "표, 스텝 번호, 통과/실패처럼 받고 싶은 형식을 적는다."],
  ["삭제·체크인 전에는 목록", "대상 목록을 먼저 보고 나서 따로 승인한다."],
  ["분석과 수정은 따로", "\"원인만 설명해줘\" → 확인 → \"이렇게 고쳐줘\" 순서로 간다."],
];

// 프롬프트 템플릿. T01~T18은 commander-mcp 스킬팩 reference/workflows-index.md 순서를 따른다.
const preamble = {
  id: "T00",
  title: "공통 머리말",
  when: "대화를 시작할 때 한 번 붙여 넣는다. 이후 요청은 T01~T18을 쓴다.",
  prompt: `앞으로 Tosca Commander MCP(tosca-commander)로 작업할 거야. 아래 규칙을 지켜줘.
1. 객체는 내가 적은 경로로 찾고, Commander 화면에서 선택된 객체에 기대지 마.
2. 무언가를 만들거나 바꾸기 전에는 쓸 도구와 대상을 표로 보여주고 내 승인을 기다려.
3. 삭제, 체크인, Checkout Tree는 따로 한 번 더 확인받아.
4. UI 모듈을 새로 스캔하지 마. 필요한 모듈이 없으면 없다고 알려줘.
5. 이름이나 설명을 바꿀 때는 메뉴 작업 말고 속성 변경(set_attribute)으로 해.
6. 작업이 끝나면 바뀐 객체와 경로를 정리하고, 저장했는지 알려줘.`,
};

const templates = [
  {
    id: "T01", title: "워크스페이스 구조 파악", workflow: "inspect-workspace", group: "조회",
    when: "처음 받은 워크스페이스에서 무엇이 어디 있는지 알고 싶을 때",
    prompt: "[폴더 경로] 아래 구조를 [깊이]단계까지 트리로 보여줘. 각 항목은 이름과 타입만 표시하고, [타입] 개수도 함께 알려줘. 아무것도 바꾸지 마.",
    example: "TestCases/주문 폴더 아래 구조를 2단계까지 트리로 보여줘. 각 항목은 이름과 타입만 표시하고, TestCase 개수도 함께 알려줘. 아무것도 바꾸지 마.",
    tools: "get_workspace_info, get_object_info, list_assets",
    caution: "항목이 많으면 AI가 여러 번 나눠 가져온다. 끝까지 다 가져왔는지 물어본다.",
  },
  {
    id: "T02", title: "테스트케이스 뼈대 만들기", workflow: "create-test-case", group: "작성",
    when: "이름과 위치가 정해진 테스트케이스를 여러 개 한 번에 만들 때",
    prompt: `[폴더 경로]에 아래 이름으로 테스트케이스를 만들어줘. 스텝은 비워 둬.
- [테스트케이스 이름 1]
- [테스트케이스 이름 2]
연결할 요구사항: [요구사항 경로 또는 없음]
만든 뒤 경로를 다시 읽어서 결과를 표로 보여줘.`,
    example: `TestCases/회원가입 폴더에 아래 이름으로 테스트케이스를 만들어줘. 스텝은 비워 둬.
- 가입_정상
- 가입_이메일 중복
- 가입_비밀번호 규칙 위반
연결할 요구사항: 없음
만든 뒤 경로를 다시 읽어서 결과를 표로 보여줘.`,
    tools: "get_object_info, create_test_cases",
    caution: "이미 있는 테스트케이스를 이 방식으로 고치면 스텝이 통째로 바뀐다. 스텝 하나 추가는 T03을 쓴다.",
  },
  {
    id: "T03", title: "기존 테스트케이스에 스텝 추가", workflow: "add-step-to-existing-test-case", group: "작성",
    when: "이미 있는 테스트케이스 끝에 스텝 하나를 덧붙일 때",
    prompt: "[테스트케이스 경로] 맨 뒤에 [모듈 경로] 모듈로 스텝을 하나 추가해줘. 값은 [속성]=[값] ([ActionMode])로 넣어줘. 기존 스텝은 건드리지 말고, 추가한 뒤 전체 스텝 목록을 보여줘.",
    example: "TestCases/주문/장바구니 담기 맨 뒤에 Modules/주문/Cart 모듈로 스텝을 하나 추가해줘. 값은 Quantity=2 (Verify)로 넣어줘. 기존 스텝은 건드리지 말고, 추가한 뒤 전체 스텝 목록을 보여줘.",
    tools: "get_object_info, execute_drop_task, get_attributes, set_attribute, save_workspace",
    caution: "\"기존 스텝은 건드리지 마\"를 꼭 넣는다. 빠지면 AI가 테스트케이스 전체를 다시 쓰는 방법을 고를 수 있다.",
  },
  {
    id: "T04", title: "TCP 설정", workflow: "configure-test-configuration-parameters", group: "편집",
    when: "URL, 브라우저 같은 실행 환경 값을 TCP로 관리할 때",
    prompt: `[대상 경로]의 Test Configuration Parameter를 상속된 값까지 먼저 보여줘. 그다음 로컬 TCP를 아래처럼 맞춰줘. 목록에 없는 기존 로컬 값은 [유지/삭제]해줘.
- [이름]=[값]
- [이름]=[값]`,
    example: `TestCases/MCP실습 폴더의 Test Configuration Parameter를 상속된 값까지 먼저 보여줘. 그다음 로컬 TCP를 아래처럼 맞춰줘. 목록에 없는 기존 로컬 값은 유지해줘.
- Browser=Chrome
- BaseUrl=https://demowebshop.tricentis.com`,
    tools: "get_object_info, list_test_configuration_parameters, create_test_configuration_parameters",
    caution: "쓰기 도구는 로컬 TCP 전체를 교체한다. '유지'를 골랐다면 결과 목록에 기존 값이 남아 있는지 확인한다.",
  },
  {
    id: "T05", title: "수동 테스트케이스 작성", workflow: "author-manual-test-case", group: "작성",
    when: "문서로 된 수동 테스트 절차를 Tosca에 옮길 때",
    prompt: `[폴더 경로]에 '[테스트케이스 이름]' 수동 테스트케이스를 만들어줘.
설명(목적, 사전 조건): [내용]
스텝 (지시 / 기대 결과):
1. [지시] / [기대 결과]
2. [지시] / [기대 결과]
작업 상태(TestCaseWorkState): [PLANNED / IN_WORK / COMPLETED]
만들기 전에 작업 순서를 먼저 보여줘.`,
    example: "Lab 2 프롬프트 2-1",
    tools: "create_test_cases, list_available_tasks, execute_task (Manual XTestStep), set_attribute, save_workspace",
    caution: "원본 문서의 스텝 수와 결과의 스텝 수가 같은지, 기대 결과가 빠지지 않았는지 본다.",
  },
  {
    id: "T06", title: "자동 테스트케이스 작성", workflow: "author-automated-test-case", group: "작성",
    when: "시나리오 문장으로 자동화 테스트케이스 초안을 만들 때",
    prompt: `[폴더 경로]에 '[테스트케이스 이름]' 자동화 테스트케이스를 만들려고 해. 아직 만들지 말고 계획만 보여줘.
- 모듈 범위: [모듈 폴더 경로]와 [표준 모듈]에서만 찾아. 새로 스캔하지 마.
- 시나리오:
  1. [동작]
  2. [동작]
- 검증: [어떤 화면 요소를 어떤 값으로 Verify]
- 재사용할 값: [버퍼 이름과 대상, 없으면 없음]
스텝별로 모듈, 속성, 값, ActionMode를 표로 보여주고, 맞는 모듈이 없는 스텝은 따로 표시해줘.`,
    example: "Lab 3 프롬프트 3-1 → 승인 후 3-2",
    tools: "list_assets, get_object_info, create_test_cases, set_attribute, execute_test_suite, execute_test_suite_status",
    caution: "모듈이 없는 스텝은 사람이 XScan으로 스캔해야 한다. 버퍼 값은 소비하는 쪽에서 {B[이름]}으로 참조된다.",
  },
  {
    id: "T07", title: "테스트케이스 설명", workflow: "explain-test-case", group: "조회",
    when: "남이 만든 테스트케이스를 이해하거나 인수인계 자료를 만들 때",
    prompt: "[테스트케이스 경로]가 무엇을 테스트하는지 설명해줘. 목적, 대상 애플리케이션, 사전 조건과 데이터, 스텝별 동작, 조건·반복·버퍼 같은 특이 로직, 자동화 상태 순서로 정리해줘. 아무것도 바꾸지 마.",
    example: "TestCases/주문/결제_카드가 무엇을 테스트하는지 설명해줘. 목적, 대상 애플리케이션, 사전 조건과 데이터, 스텝별 동작, 조건·반복·버퍼 같은 특이 로직, 자동화 상태 순서로 정리해줘. 아무것도 바꾸지 마.",
    tools: "get_object_info, get_attributes",
    caution: "스텝이 많으면 AI가 일부 값만 골라 읽는다. 전부 필요하면 \"모든 스텝의 값을 포함해줘\"를 붙인다.",
  },
  {
    id: "T08", title: "실행하고 결과 기다리기", workflow: "run-and-monitor-execution", group: "실행·분석",
    when: "테스트케이스를 검증 실행하거나 ExecutionList를 돌릴 때",
    prompt: "[테스트케이스 경로 또는 ExecutionList 경로]를 실행해줘. 테스트케이스면 ScratchBook 검증 실행으로, ExecutionList면 일반 실행으로 해줘. 실행 전에 워크스페이스를 저장하고, 끝날 때까지 기다렸다가 통과/실패/미실행 개수와 걸린 시간을 알려줘.",
    example: "Execution/스모크 테스트를 실행해줘. ExecutionList니까 일반 실행으로 해줘. 실행 전에 워크스페이스를 저장하고, 끝날 때까지 기다렸다가 통과/실패/미실행 개수와 걸린 시간을 알려줘.",
    tools: "get_object_info, save_workspace, execute_test_suite, execute_test_suite_status",
    caution: "실행은 한 번에 하나. 같은 ExecutionList를 다시 돌리면 이전 결과가 덮이므로, 비교가 필요하면 먼저 보관(Archive actual ExecutionLog)을 요청한다. Patch 버전에 따라 ExecutionList 실행은 지원되지 않을 수 있다.",
  },
  {
    id: "T09", title: "실행 결과 찾기", workflow: "locate-execution-results", group: "실행·분석",
    when: "어느 실행 결과를 봐야 할지 경로부터 찾아야 할 때",
    prompt: "Execution 폴더 아래에서 이름에 '[키워드]'가 들어간 ExecutionList를 찾아줘. 가장 최근 실행의 시작 시각, 통과/실패 개수, 결과가 있는 경로를 알려줘.",
    example: "Execution 폴더 아래에서 이름에 '야간'이 들어간 ExecutionList를 찾아줘. 가장 최근 실행의 시작 시각, 통과/실패 개수, 결과가 있는 경로를 알려줘.",
    tools: "get_object_info, get_attributes",
    caution: "실행 결과를 검색하는 전용 도구가 없어 AI가 폴더를 한 단계씩 따라간다. 3단계 안에서 못 찾으면 경로를 물어보는 것이 정상이다.",
  },
  {
    id: "T10", title: "실패 원인 분석", workflow: "analyze-execution-results", group: "실행·분석",
    when: "마지막 실행에서 실패한 테스트의 원인을 알고 싶을 때",
    prompt: "[ExecutionList 경로]의 마지막 실행에서 실패한 테스트케이스를 찾아줘. 각각 실패한 스텝, 오류 메시지, 원인 분류(애플리케이션 결함 / 테스트 데이터 / 타이밍 / 환경 / 모듈 불일치 / 기대값 불일치), 고칠 위치, 확신도(0~10)를 표로 정리해줘. 같은 원인은 묶어서 보여주고, 수정은 하지 마.",
    example: "Execution/야간 회귀의 마지막 실행에서 실패한 테스트케이스를 찾아줘. (이하 템플릿과 같음)",
    tools: "get_attributes, get_object_info",
    caution: "확신도가 낮은 항목은 사람이 로그를 직접 본다. 애플리케이션 결함이면 테스트를 고치지 말고 결함으로 등록한다.",
  },
  {
    id: "T11", title: "중단된 실행 진단", workflow: "diagnose-incomplete-run", group: "실행·분석",
    when: "실행이 끝까지 돌지 않았거나 미실행 항목이 많을 때",
    prompt: "[ExecutionList 경로]의 마지막 실행이 끝까지 돌지 않은 것 같아. 미실행(NotExecuted)과 상태 불명(Unknown) 항목이 왜 생겼는지 확인하고, 어느 테스트의 어느 스텝에서 멈췄는지, 다시 돌리기 전에 확인할 것을 알려줘.",
    example: "Execution/주간 전체 회귀의 마지막 실행이 끝까지 돌지 않은 것 같아. (이하 템플릿과 같음)",
    tools: "get_attributes, get_object_info",
    caution: "상태 불명(Unknown)은 통과도 실패도 아니다. 실행 당시 화면 상황은 AI가 볼 수 없어서 물어볼 수 있다.",
  },
  {
    id: "T12", title: "실행 이력 추세 분석", workflow: "analyze-execution-history", group: "실행·분석",
    when: "여러 번의 실행 기록에서 반복 실패나 흔들리는 테스트를 찾을 때",
    prompt: "[ExecutionList 경로]에 보관된 실행 기록을 보고 자주 실패하는 테스트케이스, 통과와 실패를 오가는(Flaky) 테스트케이스, 반복되는 오류 메시지를 정리해줘. 기간은 [기간]으로 한정해줘.",
    example: "Execution/야간 회귀에 보관된 실행 기록을 보고 (중략) 기간은 최근 2주로 한정해줘.",
    tools: "get_object_info, get_attributes",
    caution: "보관된 실행 기록(ExecutionLog)이 없으면 이력이 한 번뿐이다. 평소에 실행 결과를 보관해 두어야 쓸 수 있다.",
  },
  {
    id: "T13", title: "두 실행 비교", workflow: "compare-two-runs", group: "실행·분석",
    when: "배포 전후처럼 두 번의 실행 결과 차이를 볼 때",
    prompt: "[ExecutionList 경로]의 [첫 번째 실행]과 [두 번째 실행]을 비교해줘. 새로 실패한 것, 새로 통과한 것, 계속 실패하는 것을 나눠 표로 보여주고, 새로 실패한 것은 실패 스텝과 메시지도 적어줘.",
    example: "Execution/야간 회귀의 최신 실행과 바로 전 보관 실행을 비교해줘. (이하 템플릿과 같음)",
    tools: "get_object_info, get_attributes",
    caution: "같은 ExecutionList의 실행끼리만 비교한다. 실행 시간 차이는 같은 에이전트, 같은 환경일 때만 의미가 있다.",
  },
  {
    id: "T14", title: "결과 기반 수정과 재검증", workflow: "remediate-from-results", group: "실행·분석",
    when: "분석이 끝난 실패를 실제로 고치고 다시 돌려 볼 때",
    prompt: "[테스트케이스 경로]의 [스텝/값]을 [변경 내용]으로 고쳐줘. 원인은 [분석 결과 요약]이야. 고친 뒤 ScratchBook에서 검증 실행해서 통과하는지 보여줘. 원인이 UI 모듈 자체라면 고치지 말고 알려줘.",
    example: "TestCases/주문/결제_카드의 4번 스텝 금액 Verify 값을 12,000으로 고쳐줘. 원인은 가격 정책 변경으로 기대값이 바뀐 거야. (이하 템플릿과 같음)",
    tools: "get_attributes, set_attribute, create_test_cases, execute_test_suite, execute_test_suite_status, save_workspace",
    caution: "기대값이 정말 바뀐 경우에만 테스트를 고친다. 앱이 틀린 거라면 결함 등록이 먼저다.",
  },
  {
    id: "T15", title: "Commander 메뉴 작업 실행", workflow: "execute-task", group: "관리",
    when: "오른쪽 클릭 메뉴에 있는 작업을 AI에게 대신 시킬 때",
    prompt: "[대상 경로]에서 쓸 수 있는 오른쪽 클릭 메뉴 작업 목록을 보여줘. 그중 '[작업 이름]'을 실행하려고 하는데, 필요한 입력값이 있으면 먼저 물어봐줘.",
    example: "TestCases/MCP실습 폴더에서 쓸 수 있는 오른쪽 클릭 메뉴 작업 목록을 보여줘. 그중 'Checkout'을 실행하려고 하는데, 필요한 입력값이 있으면 먼저 물어봐줘.",
    tools: "get_object_info, list_available_tasks, execute_task",
    caution: "작업 이름은 목록에 나온 그대로(대소문자까지) 써야 한다. Scan, Rescan처럼 화면이 필요한 작업은 MCP로 하지 않는다.",
  },
  {
    id: "T16", title: "저장과 체크인", workflow: "save-and-checkin", group: "관리",
    when: "작업을 마무리하고 결과를 저장하거나 공유할 때",
    prompt: "이번 작업에서 바뀐 객체 목록을 경로와 함께 보여주고 저장해줘. (멀티 유저라면) 내가 목록을 확인한 뒤 체크인까지 해줘. 체크인 코멘트는 '[코멘트]'로 남겨줘.",
    example: "Lab 5 프롬프트 5-1, 5-2",
    tools: "save_workspace, check_in_all, update_all",
    caution: "체크인은 멀티 유저에서만 된다. 오래 편집할 때는 시작 전에 update_all로 최신 내용을 받아 두면 충돌이 줄어든다.",
  },
  {
    id: "T17", title: "Data Integrity 시작하기", workflow: "di-getting-started", group: "확장",
    when: "DB 데이터를 비교하거나 검증하는 테스트를 만들 때",
    prompt: "Data Integrity 연결 목록을 보여줘. 그중 [연결 이름]의 테이블 목록을 확인하고, [원본 테이블]과 [대상 테이블]의 건수를 비교하는 테스트케이스를 [폴더 경로]에 만들 계획을 보여줘.",
    example: "Data Integrity 연결 목록을 보여줘. 그중 DWH_STG의 테이블 목록을 확인하고, ORDERS_SRC와 ORDERS_TGT의 건수를 비교하는 테스트케이스를 TestCases/데이터검증 폴더에 만들 계획을 보여줘.",
    tools: "di_connection, get_di_connection_schema, run_di_sql_statement, create_di_db_expert_module, create_di_row_by_row_comparison",
    caution: "AI가 조회할 연결은 Connection Manager에서 'Allow AI queries'를 켜야 한다(AI가 만든 연결은 꺼진 상태로 생성). 스키마 조회는 Dev, Development, Staging, Stg 태그 연결만 된다. 비밀번호는 사람이 넣는다.",
  },
  {
    id: "T18", title: "DEX 에이전트 확인", workflow: "configure-dex-agent", group: "확장",
    when: "분산 실행(DEX)에 연결된 에이전트를 확인할 때",
    prompt: "Configurations에서 분산 실행 에이전트 목록을 새로고침해서 보여줘. TestConfiguration별로 매칭된 에이전트 수도 표로 정리해줘.",
    example: "템플릿 그대로 사용",
    tools: "get_workspace_info, list_assets, list_available_tasks, execute_task (Refresh Agents), get_object_info, get_attributes",
    caution: "멀티 유저 워크스페이스에서만 된다. 에이전트 설치, 서버 연결, 원격 데스크톱 설정은 MCP 범위 밖이다.",
  },
];

const references = [
  ["Tosca 2026.1 문서: Model Context Protocol integration", "docs.tricentis.com/tosca-2026.1/en-us/content/mcp/overview.htm"],
  ["Tosca 2026.1 문서: Use Tosca MCP Server", "docs.tricentis.com/tosca-2026.1/en-us/content/mcp/use_mcp_server.htm"],
  ["Tosca 2026.1 LTS 클라이언트 패치 노트", "docs.tricentis.com/tosca-2026.1/en-us/content/resources/webhelp/patch_notes_testsuite.htm"],
  ["Tricentis mcp-skills (Commander MCP 스킬팩)", "github.com/Tricentis/mcp-skills/tree/main/Tosca/Commander/MCP"],
  ["Model Context Protocol 소개", "modelcontextprotocol.io"],
];

module.exports = {
  meta, agenda, terms, toolGroups, clients, practiceModules, lab3Steps, labs,
  notForAI, troubleshooting, useCases, promptElements, promptPairs, tips, preamble, templates, references,
};
