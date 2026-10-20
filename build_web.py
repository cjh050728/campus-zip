# 사용법: python build_web.py 장소엑셀.xlsx 꿀팁엑셀.xlsx
# 엑셀 두 개를 읽어서 template.html 안에 데이터를 넣고 index.html을 만든다.
import sys, json, datetime, os, re
import pandas as pd

here = os.path.dirname(os.path.abspath(__file__))
place_xlsx = sys.argv[1] if len(sys.argv) > 1 else '최주희_장소정리.xlsx'
tip_xlsx = sys.argv[2] if len(sys.argv) > 2 else '최주희_꿀팁정리.xlsx'

def clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return '-'
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    if s in ('', 'nan', '확인 안 됨', '모르겠음'):
        return '-'
    return s[:10] if s[:4].isdigit() and len(s) >= 10 and s[4] == '-' and ' 00:00' in s else s

def num(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None

p = pd.read_excel(place_xlsx, sheet_name='데이터')
places = [dict(n=clean(r['이름']), a=clean(r['위치']), m=clean(r['지도']), c=clean(r['분류']), d=clean(r['확인한 날']),
               h=clean(r['혼밥']), p=num(r['1인 가격']), z=clean(r['구역']), r=num(r['추천 수']) or 1, b=clean(r['비고']),
               k=clean(r.get('음식 종류')), s=clean(r.get('한 줄 소개')), g=clean(r.get('태그')))
          for _, r in p.iterrows()]
t = pd.read_excel(tip_xlsx, sheet_name='데이터')
tips = [dict(t=clean(r['이름']), k=clean(r['종류']), s=clean(r['한 줄 요약']), b=clean(r['자세한 내용']), loc=clean(r['위치']),
             where=clean(r['자세한 위치']), m=clean(r['지도']), ref=clean(r['참고 링크']), d=clean(r['확인한 날']),
             r=num(r['추천 수']) or 1)
        for _, r in t.iterrows()]
for x in places:  # 정리용 메모(원본 번호가 들어간 부분)는 웹에서 숨김
    x['b'] = ' / '.join(seg for seg in x['b'].split(' / ') if not re.search(r'[TF]\d{3}', seg)) or '-'
for x in tips:
    for k in ('b', 's'):
        x[k] = re.sub(r'\s*\((?:[TF]\d{3}(?:-\d+)?)\)', '', x[k])  # 원본 번호 표시는 웹에서 숨김
data = dict(places=places, tips=tips, updated=datetime.date.today().isoformat())
tpl = open(os.path.join(here, 'template.html'), encoding='utf-8').read()
body = tpl.replace('__DATA__', json.dumps(data, ensure_ascii=False).replace('</', '<\\/'))
open(os.path.join(here, 'campus-zip.html'), 'w', encoding='utf-8').write(body)
# GitHub Pages용: 완전한 HTML 문서로 감싼 index.html
head = '<!doctype html>\n<html lang="ko">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
title_end = body.index('<style>')
doc = head + body[:title_end] + body[title_end:body.index('</style>') + 8] + '\n</head>\n<body>\n' + body[body.index('</style>') + 8:] + '\n</body>\n</html>\n'
doc = doc.replace('<style>', '<style>\n[hidden]{display:none!important}body{margin:0}img{max-width:100%}', 1)
open(os.path.join(here, 'index.html'), 'w', encoding='utf-8').write(doc)
print(f'장소 {len(places)}곳, 꿀팁 {len(tips)}개 → index.html')
