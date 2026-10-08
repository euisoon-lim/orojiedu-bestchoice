# -*- coding: utf-8 -*-
"""
사용법: python3 tools/build_data.py <pptx 파일>
→ data/YYYY-MM.json + data/latest.json 생성 (화면 데이터 + 로고 data URI)
"""
import sys, os, re, json, datetime, unicodedata
HERE=os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import bc_convert, bc_logos

pptx=sys.argv[1]
data=bc_convert.convert(pptx)
if not data['sections']:
    print('::error::PPTX에서 섹션(생명보험/손해보험 BEST CHOICE 슬라이드)을 찾지 못했습니다.'); sys.exit(1)
problems=[]
for s in data['sections']:
    for c in s['categories']:
        if not c.get('table'): problems.append(f"{s['name']} / {c['name']}: 비교표 없음")
        if not c.get('detail'): problems.append(f"{s['name']} / {c['name']}: 선정 이유 없음")
        if len(c['products'])<2: problems.append(f"{s['name']} / {c['name']}: 상품 2개 미만")
if problems:
    for p in problems: print('::error::'+p)
    sys.exit(1)

insurers=[p['insurer'] for s in data['sections'] for c in s['categories'] for p in c['products']]
logos,logos_w=bc_logos.build(only=insurers)
m=json.load(open(os.path.join(HERE,'logos_map.json'),encoding='utf-8'))
missing=sorted({x for x in insurers if bc_logos.resolve(m,x) is None})
if missing: print('::warning::로고 없음(텍스트 표시): '+', '.join(missing))

mm=re.search(r'(\d{4})년\s*(\d{1,2})월', data['month'] or '')
key=f"{mm.group(1)}-{int(mm.group(2)):02d}" if mm else datetime.date.today().strftime('%Y-%m')
out={'month':data['month'],'key':key,'generated_at':datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),'source':unicodedata.normalize('NFC',os.path.basename(pptx)),
     'sections':data['sections'],'logos_w':logos_w,'logos':logos}
os.makedirs(os.path.join(HERE,'..','data'),exist_ok=True)
js=json.dumps(out,ensure_ascii=False,separators=(',',':'))
open(os.path.join(HERE,'..','data',f'{key}.json'),'w',encoding='utf-8').write(js)
# latest.json 은 더 최신(또는 같은) 호일 때만 교체 — 지난 호를 테스트 업로드해도 사이트가 과거로 돌아가지 않게
lp=os.path.join(HERE,'..','data','latest.json')
try: cur_key=json.load(open(lp,encoding='utf-8')).get('key','')
except Exception: cur_key=''
if key>=cur_key:
    open(lp,'w',encoding='utf-8').write(js)
else:
    print(f'::warning::{key} 는 현재 반영본({cur_key})보다 이전 호 → data/{key}.json 만 저장, latest.json 유지')
print(f'OK {key} ({len(js)//1024} KB) sections={len(out["sections"])} categories={sum(len(s["categories"]) for s in out["sections"])} logos={len(logos)}')
