import json
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / 'presentation'))

from leak_check import collect_text, count_slides, find_leaks, forbidden_terms, main

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:dc="http://purl.org/dc/elements/1.1/"')


def make_pptx(tmp_path, slides, core=''):
    p = tmp_path / 'x.pptx'
    with zipfile.ZipFile(p, 'w') as z:
        for i, body in enumerate(slides, 1):
            z.writestr(f'ppt/slides/slide{i}.xml', f'<p:sld {NS}><p:txBody>{body}</p:txBody></p:sld>')
        z.writestr('docProps/core.xml', f'<cp:coreProperties {NS}>{core}</cp:coreProperties>')
    return p


def para(text):
    return f'<a:p><a:r><a:t>{text}</a:t></a:r></a:p>'


def test_runs_joined_within_paragraph(tmp_path):
    p = make_pptx(tmp_path, ['<a:p><a:r><a:t>X</a:t></a:r><a:r><a:t>YZ 그룹</a:t></a:r></a:p>'])
    assert collect_text(p)['ppt/slides/slide1.xml'] == 'XYZ 그룹'


def test_paragraphs_separated(tmp_path):
    p = make_pptx(tmp_path, ['<a:p><a:r><a:t>M</a:t></a:r></a:p><a:p><a:r><a:t>P</a:t></a:r></a:p>'])
    assert find_leaks(collect_text(p), ['MP']) == []


def test_collect_text_includes_doc_props(tmp_path):
    p = make_pptx(tmp_path, ['<a:p><a:r><a:t>안녕</a:t></a:r></a:p>'], core='<dc:title>MP 시스템 &amp; 과제</dc:title>')
    assert 'MP 시스템 & 과제' in collect_text(p)['docProps/core.xml']


def test_ascii_term_word_boundary():
    texts = {'s': 'ABCDEF · TEMPO · 그룹 A'}
    assert find_leaks(texts, ['ABC', 'MP']) == []
    assert find_leaks({'s': 'ABC 그룹'}, ['ABC']) == [('s', 'ABC')]


def test_korean_term_substring():
    assert find_leaks({'s': '테스트용 가상데이터'}, ['가상']) == [('s', '가상')]


def test_forbidden_terms_from_data():
    data = {'metrics': {'전체': {}, '공통': {}, 'ABC': {}}, 'weights': [{'group': 'ABC'}, {'group': 'XYZ'}],
            'case': {'sku': 'ABC123', 'name': '가나다품 라마바(QA,QB)'},
            'counter_case': {'sku': 'XYZ9', 'name': '샘플 품목 Q'},
            'top': {'rows': [{'sku': 'QQQ1'}]}}
    t = forbidden_terms(data)
    for w in ['ABC', 'XYZ', 'ABC123', 'XYZ9', 'QQQ1', '가나다품 라마바(QA,QB)', '가나다품', '라마바', 'QA', 'QB', '샘플', '품목', 'MP', '서론']:
        assert w in t
    assert 'Q' not in t and '전체' not in t
    assert len(t) == len(set(t))


def test_count_slides(tmp_path):
    p = make_pptx(tmp_path, ['<a:p/>', '<a:p/>', '<a:p/>'])
    assert count_slides(p) == 3


def test_cli_exit_codes(tmp_path):
    data = tmp_path / 'data.json'
    data.write_text(json.dumps({'metrics': {'ABC': {}}}), encoding='utf-8')
    clean = make_pptx(tmp_path, [para('깨끗한 내용')] * 3)
    assert main([str(clean), '--data', str(data)]) == 0
    assert main([str(clean), '--data', str(data), '--max-slides', '2']) == 1
    leaky_dir = tmp_path / 'l'
    leaky_dir.mkdir()
    leaky = make_pptx(leaky_dir, [para('ABC 그룹')])
    assert main([str(leaky), '--data', str(data)]) == 1


def test_doc_props_adjacent_elements_not_joined(tmp_path):
    p = make_pptx(tmp_path, [para('x')], core='<dc:title>M</dc:title><dc:creator>P</dc:creator>')
    assert find_leaks(collect_text(p), ['MP']) == []


def test_ascii_term_ignores_case():
    assert find_leaks({'s': 'logic 변경'}, ['Logic']) == [('s', 'Logic')]
    assert find_leaks({'s': 'temp'}, ['MP']) == []


def test_empty_paragraph_does_not_swallow_next(tmp_path):
    p = make_pptx(tmp_path, ['<a:p/>' + para('ABC')])
    assert 'ABC' in collect_text(p)['ppt/slides/slide1.xml']


def test_notes_slide_text_collected(tmp_path):
    p = make_pptx(tmp_path, [para('x')])
    with zipfile.ZipFile(p, 'a') as z:
        z.writestr('ppt/notesSlides/notesSlide1.xml', f'<p:notes {NS}>{para("메모 ABC")}</p:notes>')
    assert collect_text(p)['ppt/notesSlides/notesSlide1.xml'] == '메모 ABC'


def test_all_xml_parts_scanned_chart_text(tmp_path):
    p = make_pptx(tmp_path, [para('x')])
    with zipfile.ZipFile(p, 'a') as z:
        z.writestr('ppt/charts/chart1.xml', '<c:chartSpace><c:tx><c:v>ABC 그룹 추이</c:v></c:tx></c:chartSpace>')
    assert find_leaks(collect_text(p), ['ABC']) == [('ppt/charts/chart1.xml#text', 'ABC')]


def test_attribute_only_leak_found(tmp_path):
    p = make_pptx(tmp_path, ['<p:cNvPr id="2" name="그림 1" descr="XYZ 그룹 표"/>' + para('깨끗')])
    assert find_leaks(collect_text(p), ['XYZ']) == [('ppt/slides/slide1.xml#attr', 'XYZ')]


def test_layout_and_rels_scanned(tmp_path):
    p = make_pptx(tmp_path, [para('x')])
    with zipfile.ZipFile(p, 'a') as z:
        z.writestr('ppt/slideLayouts/slideLayout1.xml', f'<p:sldLayout {NS}><p:cSld name="ABC 레이아웃"/></p:sldLayout>')
        z.writestr('ppt/slides/_rels/slide1.xml.rels',
                   '<Relationships><Relationship Id="rId1" Target="https://example.com/XYZ"/></Relationships>')
    keys = {k for k, _ in find_leaks(collect_text(p), ['ABC', 'XYZ'])}
    assert keys == {'ppt/slideLayouts/slideLayout1.xml#attr', 'ppt/slides/_rels/slide1.xml.rels#attr'}


def test_namespace_uri_and_relationship_id_ignored(tmp_path):
    p = make_pptx(tmp_path, [para('x')])
    with zipfile.ZipFile(p, 'a') as z:
        z.writestr('ppt/slides/slide2.xml', '<p:sld xmlns:p="http://example.com/MP/main"><p:pic r:embed="MP"/></p:sld>')
        z.writestr('[Content_Types].xml', '<Types><Default Extension="MP" ContentType="MP"/></Types>')
    assert find_leaks(collect_text(p), ['MP']) == []


def test_cli_reports_attribute_leak(tmp_path):
    data = tmp_path / 'data.json'
    data.write_text(json.dumps({'metrics': {'ABC': {}}}), encoding='utf-8')
    d = tmp_path / 'a'
    d.mkdir()
    p = make_pptx(d, ['<p:cNvPr id="2" name="개체" descr="ABC"/>' + para('깨끗')])
    assert main([str(p), '--data', str(data)]) == 1
