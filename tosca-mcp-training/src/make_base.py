"""디자인 템플릿에서 슬라이드를 모두 빼고 마스터·레이아웃·테마·이미지만 남긴다.

python3 make_base.py <원본 템플릿.pptx> template/tricentis_base.pptx

원본 슬라이드 내용, 변경 이력, 썸네일, 이전 문서의 제목 목록(app.xml)은 남기지 않는다.
"""
import sys, zipfile, re, io
from pptx import Presentation
src, out = sys.argv[1], sys.argv[2]
APP_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
           'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
           '<Application>Microsoft Office PowerPoint</Application><PresentationFormat>Widescreen</PresentationFormat>'
           '<Slides>0</Slides><Notes>0</Notes><Company></Company><AppVersion>16.0000</AppVersion></Properties>')
prs = Presentation(src)
lst = prs.slides._sldIdLst
for sld in list(lst):
    prs.part.drop_rel(sld.rId)
    lst.remove(sld)
# 변경 이력(changesInfo) 관계 제거
for rId, rel in list(prs.part.rels.items()):
    if "changesInfo" in rel.reltype or "revisionInfo" in rel.reltype:
        prs.part.drop_rel(rId)
# 섹션 목록 제거 (원본 슬라이드 id를 가리킴)
for e in prs.part._element.iter("{http://schemas.openxmlformats.org/presentationml/2006/main}ext"):
    if any("sectionLst" in c.tag for c in e):
        e.getparent().remove(e)
cp = prs.core_properties
cp.title = ""; cp.subject = ""; cp.keywords = ""; cp.author = ""; cp.last_modified_by = ""; cp.revision = 1
buf = io.BytesIO(); prs.save(buf)
# 썸네일 제거
zin = zipfile.ZipFile(io.BytesIO(buf.getvalue()))
zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == "docProps/thumbnail.jpeg":
        continue
    if item.filename == "_rels/.rels":
        data = re.sub(rb'<Relationship [^>]*thumbnail[^>]*/>', b'', data)
    if item.filename == "docProps/app.xml":
        data = APP_XML.encode("utf-8")
    zout.writestr(item, data)
zout.close()
print("ok")
