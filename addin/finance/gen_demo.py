# 使い方: python3 gen_demo.py <事業計画ブック.xlsx>
# デモ用スナップショット demo-data.js を再生成します。
import openpyxl, json, re, sys
SRC = sys.argv[1] if len(sys.argv) > 1 else 'JBA事業計画_月次PL_BS付き.xlsx'
wb=openpyxl.load_workbook(SRC,data_only=True)

def grid(ws):
    return [[c for c in row] for row in ws.iter_rows(values_only=True)]

def num(v):
    if v is None or v=='': return None
    if isinstance(v,(int,float)): return v
    s=str(v).replace(',','').replace(' ','').replace('円','')
    try: return float(s)
    except: return None

def is_month(s): return bool(re.match(r'^\d{4}/\d{1,2}$', str(s or '').strip()))
def find_hdr(vals):
    for r,row in enumerate(vals):
        cols=[i for i,c in enumerate(row) if is_month(c)]
        if len(cols)>=6: return r,cols
    return None,None
def norm(s): return re.sub(r'[（(]旧[:：][^）)]*[）)]','',str(s or '')).replace(' ','').strip()

plv=grid(wb['月次PL(実績入力)']); bsv=grid(wb['月次BS(実績入力)'])
f36=grid(wb['36期実績(決算報告書)']); asv=grid(wb['前提条件'])

hr,cols=find_hdr(plv)
months=[str(plv[hr][c]).strip() for c in cols]
annual=1
for i,c in enumerate(plv[hr]):
    if c and '年間計画' in str(c): annual=i
pl=[]; sga_start=-1; sga_end=-1
for r in range(hr+1,len(plv)):
    nm=str(plv[r][0] or '').strip()
    if not nm: continue
    if nm.startswith('【販売費'): sga_start=len(pl); continue
    if nm.startswith('【') or nm.startswith('※'): continue
    pl.append({'name':nm,'annualPlan':num(plv[r][annual]),'v':[num(plv[r][c]) for c in cols]})
    if '販管費' in nm and '合計' in nm: sga_end=len(pl)-1
    if norm(nm)=='経常利益': break
if sga_start>=0 and sga_end>sga_start:
    for i in range(sga_start,sga_end): pl[i]['isSga']=True

bonus=[]; mnashi=None
for row in plv:
    nm=str(row[0] or '').strip()
    m=re.match(r'^決算賞与[\s\u3000]*(.+)$', nm)
    if m and m.group(1).strip()!='計':
        a=num(row[1])
        if a is not None: bonus.append({'name':m.group(1).strip(),'amt':round(a)})
    if 'みなし仕入率' in nm:
        v=num(row[1])
        if v is not None: mnashi = v/100 if v>1 else v

hr2,cols2=find_hdr(bsv)
bs=[]
for r in range(hr2+1,len(bsv)):
    nm=str(bsv[r][0] or '').strip()
    if not nm: continue
    if nm.startswith('■'): break
    bs.append({'name':nm,'v':[num(bsv[r][c]) for c in cols2]})

fy=[]
for row in f36:
    nm=str(row[0] or '').strip(); a=num(row[1] if len(row)>1 else None)
    if nm and a is not None and not nm.startswith('【'): fy.append({'name':nm,'amount':a})

assum={'salesPlan':None,'sgaPlan':None,'cogsRate':None}; term=''
for row in asv:
    nm=str(row[0] or '').strip(); v=num(row[1] if len(row)>1 else None)
    if '想定年間売上高' in nm: assum['salesPlan']=v
    elif '想定年間販管費' in nm: assum['sgaPlan']=v
    elif '売上原価率' in nm: assum['cogsRate']=v
    t=re.search(r'第(\d+)期', nm)
    if t and not term: term='第%s期'%t.group(1)

data={'months':months,'pl':pl,'bs':bs,'fy36':fy,'assum':assum,'bonusPeople':bonus,'mnashi':mnashi,'termLabel':term}
js='/* デモ用スナップショット（ブラウザで直接開いたとき用。Excel接続時は使いません） */\nconst DEMO_DATA = '+json.dumps(data,ensure_ascii=False,separators=(',',':'))+';\n'
open('demo-data.js','w',encoding='utf-8').write(js)
print('months',months)
print('pl items',len(pl),'sga',sum(1 for p in pl if p.get('isSga')))
print('bs items',len(bs),[b['name'] for b in bs][:6],'...last',bs[-1]['name'])
print('fy36',len(fy),'bonus',bonus,'mnashi',mnashi,'term',term)
print('assum',assum)
