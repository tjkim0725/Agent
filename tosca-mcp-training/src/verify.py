"""교육자료 검증 스크립트.

python3 verify.py <deck.pptx> <workbook.docx> <commander-mcp 스킬 폴더>

1. 도구 이름과 워크플로 파일 이름이 공식 스킬팩에 실제로 있는지
2. 도구 32종, 템플릿 18종 같은 개수가 맞는지
3. 실습 프롬프트가 슬라이드와 워크북에 똑같이 들어 있는지
4. 슬라이드마다 발표자 노트가 있는지
5. 이모지, 줄표(—), 과장된 표현 같은 기계적으로 쓴 티가 나는 표현이 없는지
"""
import json
import os
import re
import subprocess
import sys
import zipfile

deck, workbook, skill = sys.argv[1], sys.argv[2], sys.argv[3]
here = os.path.dirname(os.path.abspath(__file__))
content = json.loads(subprocess.check_output(
    ["node", "-e", "console.log(JSON.stringify(require(process.argv[1])))", os.path.join(here, "content.js")]))

fails, notes = [], []
def check(cond, msg):
    (notes if cond else fails).append(("OK   " if cond else "FAIL ") + msg)

# ---------- 원문 텍스트 ----------
def pptx_text(path):
    z = zipfile.ZipFile(path)
    slides = sorted([n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)], key=lambda n: int(re.findall(r"\d+", n)[0]))
    out = []
    for n in slides:
        xml = z.read(n).decode("utf-8")
        paras = re.findall(r"<a:p>(.*?)</a:p>", xml, flags=re.S)
        out.append("\n".join("".join(re.findall(r"<a:t>(.*?)</a:t>", p)) for p in paras))
    notes_parts = [n for n in z.namelist() if re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", n)]
    note_texts = ["".join(re.findall(r"<a:t>(.*?)</a:t>", z.read(n).decode("utf-8"))) for n in notes_parts]
    return out, note_texts

def docx_text(path):
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    paras = re.findall(r"<w:p[ >].*?</w:p>", xml, flags=re.S)
    return "\n".join("".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p)) for p in paras)

def unescape(s):
    return s.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'")

slides, slide_notes = pptx_text(deck)
slides = [unescape(s) for s in slides]
deck_all = "\n".join(slides)
doc_all = unescape(docx_text(workbook))
norm = lambda s: re.sub(r"\s+", "", s)

# ---------- 1. 공식 스킬팩 대조 ----------
catalog = open(os.path.join(skill, "reference", "tools-index.md"), encoding="utf-8").read()
official_tools = set(re.findall(r"`([a-z_]+)`", catalog))
check(len(official_tools) == 32, f"스킬팩 tools-index의 도구 수 = {len(official_tools)}")
our_tools = [t[0] for g in content["toolGroups"] for t in g["tools"]]
check(len(our_tools) == 32 and set(our_tools) == official_tools, f"교육자료 도구 목록 {len(our_tools)}개가 공식 목록과 일치")

ident = re.compile(r"\b[a-z]+(?:_[a-z]+)+\b")
not_tools = {"is_validation_run", "test_steps", "linked_requirements", "test_case_id", "execution_list_id",
             # 도구가 아니라 파일·문서 이름
             "claude_desktop_config", "patch_notes_testsuite", "use_mcp_server"}
mentioned = set()
for text in (deck_all, doc_all):
    mentioned |= set(ident.findall(text))
unknown = sorted(m for m in mentioned if m not in official_tools and m not in not_tools)
check(not unknown, "본문에 나온 snake_case 이름이 모두 실제 도구 또는 실제 파라미터" + (f" (확인 필요: {unknown})" if unknown else ""))

wf_dir = os.path.join(skill, "reference", "workflows")
wf_files = {f[:-3] for f in os.listdir(wf_dir) if f.endswith(".md")}
missing_wf = [t["workflow"] for t in content["templates"] if t["workflow"] not in wf_files]
check(not missing_wf and len(wf_files) == 18, f"템플릿 18종의 공식 워크플로 파일이 모두 존재 (스킬팩 워크플로 {len(wf_files)}개)")
index_order = re.findall(r"\(workflows/([a-z0-9-]+)\.md\)", open(os.path.join(skill, "reference", "workflows-index.md"), encoding="utf-8").read())
check(index_order == [t["workflow"] for t in content["templates"]], "T01~T18 순서가 workflows-index.md 순서와 같음")

# ---------- 2. 개수 ----------
check(len(content["templates"]) == 18, "템플릿 18종")
groups = {}
for t in content["templates"]:
    groups[t["group"]] = groups.get(t["group"], 0) + 1
check(groups == {"조회": 2, "작성": 4, "편집": 1, "실행·분석": 7, "관리": 2, "확장": 2}, f"템플릿 묶음별 개수 {groups}")
check(len(content["labs"]) == 6, "Lab 0~5 여섯 개")
check("도구 32개" in deck_all and "도구 32종" in doc_all, "슬라이드와 워크북 모두 도구 수를 32로 표기")

# ---------- 3. 프롬프트 일관성 ----------
for lab in content["labs"]:
    for p in lab["prompts"]:
        key = norm(p["text"].replace("(멀티 유저일 때) ", ""))
        in_doc = key in norm(doc_all)
        in_deck = key in norm(deck_all)
        check(in_doc, f"워크북에 프롬프트 {p['label']} 원문이 있음")
        check(in_deck, f"슬라이드에 프롬프트 {p['label']} 원문이 있음")
for t in [content["preamble"]] + content["templates"]:
    check(norm(t["prompt"]) in norm(doc_all), f"워크북에 {t['id']} 템플릿 본문이 있음")
t06 = next(t for t in content["templates"] if t["id"] == "T06")
check(norm(t06["prompt"]) in norm(deck_all), "슬라이드 31의 T06 본문이 워크북과 같음")

# ---------- 4. 발표자 노트 ----------
check(len(slides) == len(slide_notes) and all(len(n) > 40 for n in slide_notes), f"슬라이드 {len(slides)}장 모두 발표자 노트 있음")

# ---------- 5. 표현 점검 ----------
emoji = re.compile("[\U0001F300-\U0001FAFF☀-⛿✀-➿⭐✅]")
src_all = deck_all + doc_all + "".join(slide_notes)
found_emoji = sorted(set(c for c in emoji.findall(src_all) if c != "☐"))
check(not found_emoji, "이모지 없음 (체크박스 ☐ 제외)" + (f" {found_emoji}" if found_emoji else ""))
check("—" not in src_all and "–" not in src_all, "줄표(—, –) 없음")
check("!" not in src_all.replace("!=", ""), "느낌표 없음")
banned = ["혁신", "획기적", "강력한", "손쉽게", "한눈에", "극대화", "게임 체인저", "놀라운", "완벽한", "마법", "~를 통해", "최첨단", "원활한", "시너지"]
hits = [w for w in banned if w in src_all]
check(not hits, "과장된 상투 표현 없음" + (f" {hits}" if hits else ""))

print("\n".join(notes + fails))
print(f"\n{len(notes)}개 통과, {len(fails)}개 실패")
sys.exit(1 if fails else 0)
