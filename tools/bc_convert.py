# -*- coding: utf-8 -*-
"""
베스트 초이스 PPTX → JSON 변환기 (오로지교육)
사용법:  python3 bc_convert.py "26년 10월 상품 비교 분석 가이드_ 베스트 초이스화면.pptx" > bc_data.json

입력 PPTX 규칙(웰스에듀 '화면' 버전):
  - 섹션 개요 슬라이드 : 'BEST CHOICE' 텍스트 + '생명보험'/'손해보험' 텍스트, 그룹(카테고리명 / 상품 3줄)
  - 비교표 슬라이드    : 제목 '타사 비교표_ {보험사} – {상품}', 표 1개, 우측 상단 카테고리 라벨, 각주 텍스트박스
  - 선정이유 슬라이드  : 제목 '선정 이유_ ...', 둥근사각형 라벨(상품 특징 / 선정 이유 / 세일즈 포인트) + 내용 박스
카테고리 ↔ 비교표/선정이유 매칭: 같은 섹션 안에서 BEST(첫 번째) 상품의 보험사명으로 매칭
"""
import sys, re, json
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

DASH = r'[–\-—]'
LABELS = {'상품특징': 'feature', '선정이유': 'reason', '세일즈포인트': 'sales'}
ICON_RULES = [  # (키워드, Tossface 이모지) — 앞에 있는 규칙이 우선
    ('비급여', '💊'), ('간병', '🛏️'), ('순환', '🫀'), ('뇌', '🧠'), ('심장', '🫀'), ('암', '🎗️'), ('관절', '🦴'), ('재해', '🦴'), ('골절', '🦴'),
    ('연금', '💰'), ('변액', '📈'), ('수술', '🏥'), ('간병', '🛏️'), ('비급여', '💊'), ('치아', '🦷'), ('운전', '🚗'),
    ('어린이', '🧒'), ('태아', '👶'), ('실손', '🩺'), ('종신', '🛡️'), ('저축', '🏦'),
]

def norm(s):
    return re.sub(r'[\s/／·ㆍ,\.\n]+', '', s or '')

def paras(shape):
    if not shape.has_text_frame:
        return []
    return [p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip()]

def text(shape):
    return '\n'.join(paras(shape))

def split_product(line):
    """'삼성생명 – 인생대표 / 가족대표' → ('삼성생명', '인생대표 / 가족대표')"""
    line = re.sub(r'\s+', ' ', line).strip()
    m = re.match(r'^(.+?)\s*' + DASH + r'\s*(.+)$', line)
    if not m:
        return line, ''
    return m.group(1).strip(), m.group(2).strip()

def pick_icon(name):
    for k, e in ICON_RULES:
        if k in name:
            return e
    return '✅'

def iter_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(sh.shapes)
        else:
            yield sh

def table_to_grid(tbl):
    """표를 그리드 셀 목록으로 변환. 각 셀: t(text), r(row), c(col), h(rowspan), w(colspan)"""
    rows = []
    for ri, row in enumerate(tbl.rows):
        cells = []
        for ci, cell in enumerate(row.cells):
            if cell.is_spanned:
                continue
            d = {'t': cell.text.strip(), 'c': ci}
            if cell.is_merge_origin:
                if cell.span_height > 1: d['h'] = cell.span_height
                if cell.span_width > 1: d['w'] = cell.span_width
            cells.append(d)
        rows.append(cells)
    return {'cols': len(tbl.columns), 'rows': rows}

def convert(path):
    prs = Presentation(path)
    m = re.search(r'(\d{2})년\s*(\d{1,2})월', path)
    month = f'20{m.group(1)}년 {int(m.group(2))}월' if m else ''

    sections = []   # [{name, categories:[...]}]
    cur = None

    for sidx, slide in enumerate(prs.slides, 1):
        shapes = list(slide.shapes)
        flat = list(iter_shapes(shapes))
        texts = [text(s) for s in flat]
        joined = '\n'.join(texts)

        # ── 1) 섹션 개요 슬라이드 ─────────────────────────────
        if 'BESTCHOICE' in norm(joined) and any(s.shape_type == MSO_SHAPE_TYPE.GROUP for s in shapes):
            sec_name = ''
            for t in texts:
                tt = t.strip()
                if tt in ('생명보험', '손해보험', '제3보험'):
                    sec_name = tt
            if not sec_name:
                mm = re.search(r'(생명보험|손해보험)', joined)
                sec_name = mm.group(1) if mm else f'섹션{len(sections)+1}'
            cur = {'name': sec_name, 'categories': []}
            groups = [s for s in shapes if s.shape_type == MSO_SHAPE_TYPE.GROUP]
            groups.sort(key=lambda g: (g.top, g.left))
            for g in groups:
                boxes = [s for s in g.shapes if s.has_text_frame and text(s)]
                boxes.sort(key=lambda s: s.left)
                if len(boxes) < 2:
                    continue
                cat_name = ' '.join(paras(boxes[0])).replace(' / ', '/').strip()
                cat_name = re.sub(r'\s*/\s*', '/', cat_name)
                prods = []
                for line in paras(boxes[1]):
                    ins, pn = split_product(line)
                    prods.append({'insurer': ins, 'name': pn})
                cur['categories'].append({
                    'name': cat_name, 'icon': pick_icon(cat_name), 'label': '',
                    'products': prods, 'table': None, 'notes': [], 'detail': {},
                })
            sections.append(cur)
            continue

        if cur is None:
            continue

        title = next((t for t in texts if re.match(r'^(타사\s*비교표|선정\s*이유)', t.strip())), None)
        if not title:
            continue
        kind = 'table' if title.strip().startswith('타사') else 'reason'
        body = re.sub(r'^(타사\s*비교표|선정\s*이유)\s*_?\s*', '', title.strip())
        ins, pn = split_product(body)

        # 카테고리 매칭: BEST 상품 보험사명
        cat = next((c for c in cur['categories'] if c['products'] and norm(c['products'][0]['insurer']) == norm(ins)), None)
        if cat is None:
            # 보조: 상품명 앞 6글자 일치
            cat = next((c for c in cur['categories'] if c['products'] and norm(c['products'][0]['name'])[:6] == norm(pn)[:6]), None)
        if cat is None:
            print(f'[경고] slide {sidx}: 카테고리 매칭 실패 → {title}', file=sys.stderr)
            continue

        if kind == 'table':
            tbl_shape = next((s for s in flat if getattr(s, 'has_table', False) and s.has_table), None)
            if tbl_shape is not None:
                cat['table'] = table_to_grid(tbl_shape.table)
            # 우측 상단 라벨(카테고리 보조명) + 각주
            for s in flat:
                if not s.has_text_frame or s is tbl_shape:
                    continue
                t = text(s)
                if not t or t == title or re.fullmatch(r'\d+', t):
                    continue
                if s.top < prs.slide_height * 0.06 and s.left > prs.slide_width * 0.6:
                    cat['label'] = ' '.join(paras(s))          # 우측 상단 카테고리 라벨
                else:
                    cat['notes'].append(' '.join(paras(s)))    # 기준·각주 텍스트
        else:
            detail = {}
            key = None
            for s in flat:
                if not s.has_text_frame:
                    continue
                ps = paras(s)
                if not ps:
                    continue
                k = LABELS.get(norm(''.join(ps)))
                if k:
                    key = k
                    continue
                if key and s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                    detail[key] = [re.sub(r'\s+', ' ', p) for p in ps]
                    key = None
            cat['detail'] = detail

    return {'month': month, 'sections': sections}

if __name__ == '__main__':
    data = convert(sys.argv[1])
    json.dump(data, sys.stdout, ensure_ascii=False, indent=1)
