from pathlib import Path
from collections import defaultdict
import datetime, json, re, openpyxl

ROOT=Path('.')
target=ROOT/'baocaokpi_thang8_hoanthien.xlsx'
invoice=next((ROOT/'thang8').glob('DanhSachChiTietHoaDon*.xlsx'))
returns=next((ROOT/'thang8').glob('DanhSachChiTietTraHang*.xlsx'))

def num(v):
    try: return float(v or 0)
    except: return 0.0
def cmap(row): return {str(v).strip().lower():i for i,v in enumerate(row) if v is not None}
def find(m,names):
    return next((m[x] for x in names if x in m),None)
def branch(x):
    s=str(x or '').strip(); d={'NT Hàng Bông 247':'Hàng Bông','NT Đường Láng 247':'Đường Láng','NT 24H Minh Châu 1':'Minh Châu','NT 24h Đỗ Quang Đẩu':'Đỗ Quang Đẩu','NT 24H Lê Bình':'Lê Bình','NT 24H Nam Hòa':'Nam Hòa','NT 24H Nguyễn Chí Thanh':'Nguyễn Chí Thanh','NT 24H Nguyễn Thị Thập':'Nguyễn Thị Thập','NT 24H Nguyễn Văn Quá':'Nguyễn Văn Quá','NT 24H Rạch Bùng Binh':'Rạch Bùng Binh','NT 24H Trường Sa':'Trường Sa'}
    if s in d:return d[s]
    s=re.sub(r'^NT\s+(24H\s+|24h\s+)?','',s,flags=re.I);s=re.sub(r'\s+247$','',s);return re.sub(r'\s+1$','',s).strip()

# Output Data totals
wbo=openpyxl.load_workbook(target,read_only=True,data_only=True)
ws=wbo['data']; out=defaultdict(lambda:[0,0,0,0])
for r in ws.iter_rows(min_row=2,max_col=8,values_only=True):
    if r[0] and r[1]:
        a=out[(str(r[0]).strip(),str(r[1]).strip())]
        for i in range(4): a[i]+=num(r[[2,3,4,5][i]])
outtot=[sum(x[i] for x in out.values()) for i in range(4)]

# Source invoice totals with same documented grouping
wbi=openpyxl.load_workbook(invoice,read_only=True,data_only=True); wi=wbi.active; it=wi.iter_rows(values_only=True); m=cmap(next(it))
cols={k:find(m,v) for k,v in {'br':['chi nhánh','chinhanh'],'id':['mã hóa đơn','mahoadon'],'seller':['người bán','nguoiban'],'ch':['kênh bán','kenhban'],'pay':['khách cần trả','khachcantra'],'other':['thu khác','thukhac'],'item':['tên hàng','tenhang'],'amt':['thành tiền','thanhtien'],'total':['tổng tiền hàng','tongtienhang'],'disc':['giảm giá hóa đơn','giamgiahoadon']}.items()}
invs={}
for r in it:
    if not r or any(v is None for v in (cols['br'],cols['id'],cols['seller'],cols['pay'])): continue
    if len(r)<=max(v for v in cols.values() if v is not None):continue
    id=r[cols['id']]; s=str(r[cols['seller']] or '').strip()
    if not id or not s:continue
    if id not in invs: invs[id]={'b':branch(r[cols['br']]),'s':s,'ch':str(r[cols['ch']] or '').strip(),'pay':num(r[cols['pay']]),'other':num(r[cols['other']]) if cols['other'] is not None else 0,'items':[]}
    invs[id]['items'].append((str(r[cols['item']] or '').strip(),num(r[cols['amt']]),num(r[cols['total']]),num(r[cols['disc']])))
src=defaultdict(lambda:[0,0,0,0])
for x in invs.values():
    a=src[(x['b'],x['s'])]; i=2 if x['ch']=='Medigo' else 0; a[i]+=1;a[i+1]+=x['pay']-x['other']
srctot=[sum(x[i] for x in src.values()) for i in range(4)]
mis=[]
for k in sorted(set(src)|set(out)):
    if any(abs(src[k][i]-out[k][i])>.5 for i in range(4)):mis.append((k,src[k],out[k]))

# Project reconciliation: replicate current tool's seller-only classification, including returns
proj=defaultdict(lambda:[0.0,0.0,0.0])
for x in invs.values():
    for name,amt,total,disc in x['items']:
        net=amt-(disc*amt/total if total>0 else 0)
        prefix=name.split()[0].upper() if name else ''
        if prefix.startswith('NY3'): proj[x['s']][0]+=net
        elif prefix.startswith('CK'): proj[x['s']][1]+=net
        elif prefix.startswith('COMBO'): proj[x['s']][2]+=net
wbr=openpyxl.load_workbook(returns,read_only=True,data_only=True); wr=wbr.active; rit=wr.iter_rows(values_only=True); rm=cmap(next(rit))
rs=find(rm,['người bán','nguoiban']); rn=find(rm,['tên hàng','tenhang']); rv=find(rm,['tổng sau giảm giá','tongsaugiamgia','tổng sau gg'])
ret_count=0
if None not in (rs,rn,rv):
    for r in rit:
        if r and len(r)>max(rs,rn,rv) and r[rs] and r[rn]:
            seller=str(r[rs]).strip(); prefix=str(r[rn]).strip().split()[0].upper(); val=num(r[rv]); ret_count+=1
            if prefix.startswith('NY3'): proj[seller][0]-=val
            elif prefix.startswith('CK'): proj[seller][1]-=val
            elif prefix.startswith('COMBO'): proj[seller][2]-=val
wp=wbo['Dự án T8']; output_proj=[]
for r in wp.iter_rows(min_row=3,max_col=7,values_only=True):
    if r[1]: output_proj.append((str(r[0]).strip(),str(r[1]).strip(),[num(r[4]),num(r[5]),num(r[6])]))
dup_names=[n for n in set(x[1] for x in output_proj) if sum(1 for x in output_proj if x[1]==n)>1]
pmis=[x for x in output_proj if any(abs(x[2][i]-proj[x[1]][i])>.5 for i in range(3))]

# cache errors
errs=[]
for s in wbo.worksheets:
    for row in s.iter_rows():
        for c in row:
            if isinstance(c.value,str) and c.value.startswith('#'):errs.append((s.title,c.coordinate,c.value))
print(json.dumps({'target':str(target),'source_invoices':len(invs),'output_data_rows':len(out),'source_data_rows':len(src),'output_totals_off_tx_off_rev_on_tx_on_rev':outtot,'source_totals_off_tx_off_rev_on_tx_on_rev':srctot,'data_mismatch_count':len(mis),'data_mismatch_samples':mis[:10],'returns_applied_to_project':ret_count,'project_rows':len(output_proj),'project_mismatch_count':len(pmis),'project_mismatch_samples':pmis[:10],'duplicate_project_names':dup_names,'cached_formula_errors':errs[:50],'cached_formula_error_count':len(errs)},ensure_ascii=False,default=str,indent=2))
