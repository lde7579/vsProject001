"""공개용 pptx 금지어·장 수 검사기.

사내 고유명, SKU 코드, 제품명 등 금지어가 pptx에 남았는지와 장 수 초과 여부를 검사한다.
금지어는 data.json에서 뽑은 값과 고정 목록(FIXED_TERMS)을 합쳐 쓴다.

검사 범위는 두 단계다.
  1단계: 슬라이드·노트의 <a:t> 글자(문단 안 run은 이어 붙임)와 docProps.
  2단계: [Content_Types].xml을 뺀 모든 XML 파트(.xml, .rels)의 태그를 걷어 낸 글자와 속성값.
         차트, 레이아웃·마스터, 대체 텍스트(descr), 개체 이름(name), 메모 등을 덮는다.
         속성값 중 이름공간 선언(xmlns)과 관계 id(r:id, r:embed 등)만 뺀다.

실행: uv run python presentation/leak_check.py <pptx> [--data presentation/data.json] [--max-slides 15]
      --data를 주지 않으면 이 파일과 같은 폴더의 data.json을 쓴다(실행 위치와 무관).
종료 코드: 문제 없으면 0, 누출 또는 장 수 초과가 있으면 1.
"""
import argparse
import json
import re
import sys
import zipfile
from html import unescape
from pathlib import Path

FIXED_TERMS = ['MP', 'ITSM', 'SCM', 'WMS', '구매마스터', 'Logic', '사진', '가상', '엑셀', 'xlsx',
               '8/24', '8/28', '0824', '20260828', '1.645', '1.282', '1.036', '보정계수',
               '서론', '본론', '결론']
EXCLUDED_GROUPS = {'전체', '공통', '미분류'}

_PARA_RE = re.compile(r'<a:p/>|<a:p[ >].*?</a:p>', re.S)
_TEXT_RE = re.compile(r'<a:t(?:\s[^>]*)?>(.*?)</a:t>', re.S)
_SLIDE_RE = re.compile(r'^ppt/slides/slide\d+\.xml$')
_NOTE_RE = re.compile(r'^ppt/notesSlides/[^/]+\.xml$')
_PROPS = ('docProps/core.xml', 'docProps/app.xml')
_XML_PART_RE = re.compile(r'\.(xml|rels)$')
_SKIP_PARTS = {'[Content_Types].xml'}
_TAG_RE = re.compile(r'<[^>]+>')
_ATTR_RE = re.compile(r"""([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')""")
# 속성값 검사에서 빼는 것: 이름공간 URI 선언과 관계 id(파트끼리 잇는 내부 번호)
_SKIP_ATTRS = {'Id', 'r:id', 'r:embed', 'r:link', 'r:pict', 'r:dm', 'r:lo', 'r:qs', 'r:cs'}


def _paragraph_text(xml: str) -> str:
    """슬라이드·노트 XML에서 문단별로 <a:t>를 이어 붙이고 문단 사이는 줄바꿈으로 나눈다."""
    paras = [''.join(unescape(t) for t in _TEXT_RE.findall(p)) for p in _PARA_RE.findall(xml)]
    return '\n'.join(paras)


def _attr_values(xml: str) -> str:
    """태그 안 속성값을 줄바꿈으로 이어 붙인다. xmlns 선언과 관계 id는 뺀다."""
    vals = []
    for tag in _TAG_RE.findall(xml):
        for key, v1, v2 in _ATTR_RE.findall(tag):
            if key == 'xmlns' or key.startswith('xmlns:') or key in _SKIP_ATTRS:
                continue
            vals.append(unescape(v1 or v2))
    return '\n'.join(v for v in vals if v)


def collect_text(pptx_path: Path) -> dict[str, str]:
    """검사 키 -> 글자.

    1단계 키는 zip 안 경로 그대로다(슬라이드, 노트, docProps).
    2단계 키는 '<경로>#text'(태그를 걷어 낸 글자)와 '<경로>#attr'(속성값)이다.
    """
    out: dict[str, str] = {}
    with zipfile.ZipFile(pptx_path) as z:
        for name in z.namelist():
            if _SLIDE_RE.match(name) or _NOTE_RE.match(name):
                out[name] = _paragraph_text(z.read(name).decode('utf-8'))
            elif name in _PROPS:
                out[name] = unescape(_TAG_RE.sub(' ', z.read(name).decode('utf-8')))
        for name in z.namelist():
            if name in _SKIP_PARTS or not _XML_PART_RE.search(name):
                continue
            xml = z.read(name).decode('utf-8', errors='replace')
            # 태그 자리는 공백으로 바꿔 서로 다른 요소의 글자가 붙지 않게 한다
            out[f'{name}#text'] = unescape(_TAG_RE.sub(' ', xml))
            out[f'{name}#attr'] = _attr_values(xml)
    return out


def forbidden_terms(data: dict) -> list[str]:
    """data.json에서 뽑은 금지어와 고정 목록을 중복 없이 모은다."""
    terms: list[str] = []

    def add(t):
        if isinstance(t, str) and t and t not in terms:
            terms.append(t)

    for g in data.get('metrics', {}):
        if g not in EXCLUDED_GROUPS:
            add(g)
    for w in data.get('weights', []):
        add(w.get('group'))
    for key in ('case', 'counter_case'):
        c = data.get(key) or {}
        add(c.get('sku'))
        name = c.get('name')
        add(name)
        if isinstance(name, str):
            for piece in re.split(r'[\s()\[\],/·]+', name):
                if len(piece) >= 2:
                    add(piece)
    for r in (data.get('top') or {}).get('rows', []):
        add(r.get('sku'))
    for t in FIXED_TERMS:
        add(t)
    return terms


def find_leaks(texts: dict[str, str], terms: list[str]) -> list[tuple[str, str]]:
    """(zip 경로, 금지어) 목록. 영문·숫자 금지어는 대소문자 무시 단어 경계로, 그 밖은 부분 문자열로 찾는다."""
    found = []
    for path, text in texts.items():
        for term in terms:
            if re.fullmatch(r'[A-Za-z0-9]+', term):
                hit = re.search(rf'(?<![A-Za-z0-9]){re.escape(term)}(?![A-Za-z0-9])', text, re.IGNORECASE)
            else:
                hit = term in text
            if hit:
                found.append((path, term))
    return found


def count_slides(pptx_path: Path) -> int:
    with zipfile.ZipFile(pptx_path) as z:
        return sum(1 for n in z.namelist() if _SLIDE_RE.match(n))


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description='pptx 금지어·장 수 검사')
    ap.add_argument('pptx')
    ap.add_argument('--data', default=str(Path(__file__).resolve().parent / 'data.json'))
    ap.add_argument('--max-slides', type=int, default=15)
    args = ap.parse_args(argv)

    data = json.loads(Path(args.data).read_text(encoding='utf-8'))
    leaks = find_leaks(collect_text(Path(args.pptx)), forbidden_terms(data))
    n = count_slides(Path(args.pptx))
    for path, term in leaks:
        print(f"{path}: '{term}'")
    bad = bool(leaks)
    if n > args.max_slides:
        print(f'장 수 초과: {n}장 (최대 {args.max_slides}장)')
        bad = True
    if bad:
        return 1
    print(f'금지어 0건 · {n}장')
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
