# -*- coding: utf-8 -*-
"""보험사 로고 → 자동 크롭 + 흰 배경 투명화 + 리사이즈 → data URI 맵 생성"""
import json, io, base64, os
from PIL import Image, ImageDraw
from collections import deque
HERE=os.path.dirname(os.path.abspath(__file__))
H=48  # 출력 높이(px) — 화면 표시 20~24px 의 2배

def flood_white_to_alpha(im, thr=235):
    im=im.convert('RGBA'); w,h=im.size; px=im.load()
    seen=bytearray(w*h); q=deque()
    def white(p): return p[0]>=thr and p[1]>=thr and p[2]>=thr
    for x in range(w):
        for y in (0,h-1):
            if white(px[x,y]): q.append((x,y))
    for y in range(h):
        for x in (0,w-1):
            if white(px[x,y]): q.append((x,y))
    while q:
        x,y=q.popleft(); i=y*w+x
        if seen[i]: continue
        seen[i]=1
        p=px[x,y]
        if not white(p): continue
        px[x,y]=(255,255,255,0)
        if x>0: q.append((x-1,y))
        if x<w-1: q.append((x+1,y))
        if y>0: q.append((x,y-1))
        if y<h-1: q.append((x,y+1))
    return im

def process(path):
    im=Image.open(path)
    im=flood_white_to_alpha(im)
    bbox=im.getchannel('A').getbbox()
    if bbox:
        pad=max(2,int((bbox[3]-bbox[1])*0.06))
        im=im.crop((max(0,bbox[0]-pad),max(0,bbox[1]-pad),min(im.width,bbox[2]+pad),min(im.height,bbox[3]+pad)))
    r=H/im.height
    im=im.resize((max(1,round(im.width*r)),H),Image.LANCZOS)
    # v6.1: 시각적 크기 통일 — 가로로 긴 워드마크는 bbox 기하평균(sqrt(w*h)) 기준으로 축소해 투명 캔버스(높이 H)에 세로 중앙 배치
    REF=H*(4.5**0.5)            # 가로:세로 4.5:1 로고를 기준 크기로
    s=(im.width*im.height)**0.5
    f=max(0.70,min(1.0,REF/s))
    if f<0.995:
        sm=im.resize((max(1,round(im.width*f)),max(1,round(H*f))),Image.LANCZOS)
        cv=Image.new('RGBA',(sm.width,H),(0,0,0,0)); cv.alpha_composite(sm,(0,(H-sm.height)//2)); im=cv
    buf=io.BytesIO(); im.save(buf,'WEBP',quality=82,method=6)
    return 'data:image/webp;base64,'+base64.b64encode(buf.getvalue()).decode(), im

def to_white(im):
    """다크 배경용 흰색 단색 버전: 밝기/채도 기반 알파(흰 내부 글자는 뚫림)"""
    im=im.convert('RGBA'); px=im.load(); w,h=im.size
    out=Image.new('RGBA',im.size,(255,255,255,0)); po=out.load()
    for y in range(h):
        for x in range(w):
            r,g,b,a=px[x,y]
            if a==0: continue
            mx,mn=max(r,g,b),min(r,g,b)
            lum=(0.299*r+0.587*g+0.114*b)/255.0
            sat=(mx-mn)/mx if mx else 0
            k=min(1.0,max(1.0-lum, sat)*1.15)
            po[x,y]=(255,255,255,int(a*k))
    return out

def to_uri(im,q=82):
    buf=io.BytesIO(); im.save(buf,'WEBP',quality=q,method=6)
    return 'data:image/webp;base64,'+base64.b64encode(buf.getvalue()).decode()

def resolve(m, name):
    n=(name or '').replace(' ','')
    cands=[n, n.replace('손보','손해보험'), n.replace('생명보험','생명'), n.replace('손해보험','손보')]
    for c in cands:
        if c in m: return c
    for k in m:
        if k.startswith(n) or n.startswith(k): return k
    return None

def build(only=None):
    """only: 포함할 보험사명 목록(None이면 전체)"""
    m=json.load(open(os.path.join(HERE,'logos_map.json'),encoding='utf-8'))
    if only is not None:
        keys=set(filter(None,(resolve(m,x) for x in only)))
        m={k:v for k,v in m.items() if k in keys}
    out={}; outw={}; sheet=[]
    for name,f in m.items():
        uri,im=process(os.path.join(HERE,'..','logos',f)); out[name]=uri; sheet.append((name,im))
        outw[name]=to_uri(to_white(im))
    # 검수용 시트
    cols=5; cw=240; ch=100
    S=Image.new('RGB',(cols*cw,((len(sheet)+cols-1)//cols)*ch),(20,16,40)); d=ImageDraw.Draw(S)
    for i,(n,im) in enumerate(sheet):
        x=(i%cols)*cw; y=(i//cols)*ch
        bg=Image.new('RGBA',(im.width+16,im.height+16),(255,255,255,255)); bg.alpha_composite(im,(8,8))
        if bg.width>cw-10: bg=bg.resize((cw-10,int(bg.height*(cw-10)/bg.width)))
        S.paste(bg,(x+5,y+5)); d.text((x+5,y+ch-16),n,fill=(255,255,255))
        wi=to_white(im); S.alpha_composite(wi.resize((min(wi.width,cw-10),int(wi.height*min(1,(cw-10)/wi.width)))),(x+5,y+5+bg.height+2)) if False else None
    S.save(os.path.join(HERE,'logos_check.png'))
    return out, outw

if __name__=='__main__':
    o,w=build(); print(len(o),'logos', sum(len(v) for v in o.values())//1024,'KB color', sum(len(v) for v in w.values())//1024,'KB white')
