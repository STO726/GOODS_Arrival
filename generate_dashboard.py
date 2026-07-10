import pandas as pd
import json, re
from datetime import date

today    = pd.Timestamp(date.today())
week_end = today + pd.Timedelta(days=6)

df_data = pd.read_excel('Article search.xlsx', sheet_name='Data')
df_roi  = pd.read_excel('Article search.xlsx', sheet_name='ROI')
df_data['rcv_date'] = pd.to_datetime(
    df_data['Planned Logistics Receiving(Date)'], errors='coerce')

roi_json = []
for _, row in df_roi.iterrows():
    item_no   = str(int(row['ItemNo']))      if pd.notna(row['ItemNo'])   else ''
    item_name = str(row['ItemName']).strip() if pd.notna(row['ItemName']) else ''
    pa        = str(int(row['PA']))          if pd.notna(row['PA'])       else ''
    rcv_date  = str(row['LatestRcvDate'])[:10] if pd.notna(row['LatestRcvDate']) else ''
    qty       = int(row['LatestQty'])        if pd.notna(row['LatestQty']) else 0
    csm_id    = str(row['CSM id']).strip()   if pd.notna(row['CSM id'])   else ''
    ssd       = str(row['SSD'])[:10]         if pd.notna(row['SSD'])      else ''
    eds       = str(row['EDS'])[:10]         if pd.notna(row['EDS'])      else ''
    typ       = str(row['TYPE']).strip()     if pd.notna(row['TYPE'])     else ''
    roi_json.append({'itemNo':item_no,'itemName':item_name,'pa':pa,
                     'latestRcvDate':rcv_date,'qty':qty,'csmId':csm_id,
                     'ssd':ssd,'eds':eds,'type':typ})

csm_map = {}
for _, row in df_data.iterrows():
    cid = str(row['Csm Id']).strip() if pd.notna(row['Csm Id']) else ''
    if cid:
        csm_map[cid] = {
            'planned': row['rcv_date'].strftime('%Y-%m-%d') if pd.notna(row['rcv_date']) else '',
            'bl':   str(row['BL'])            if pd.notna(row['BL'])            else '',
            'cont': str(row['Container NO.']) if pd.notna(row['Container NO.']) else ''
        }
for item in roi_json:
    m = csm_map.get(item['csmId'], {})
    item['plannedRcvDate'] = m.get('planned','')
    item['bl']             = m.get('bl','')
    item['container']      = m.get('cont','')

week_rows = df_data[
    (df_data['rcv_date'] >= today) & (df_data['rcv_date'] <= week_end)].copy()
seen, week_containers = set(), []
for _, row in week_rows.iterrows():
    cont = str(row['Container NO.']).strip() if pd.notna(row['Container NO.']) else ''
    if cont in seen: continue
    seen.add(cont)
    cid = str(row['Csm Id']).strip() if pd.notna(row['Csm Id']) else ''
    articles = [{'itemNo':it['itemNo'],'itemName':it['itemName'],
                 'qty':it['qty'],'type':it['type']}
                for it in roi_json if it['csmId']==cid]
    week_containers.append({
        'type':     str(row['TYPE']).strip()         if pd.notna(row['TYPE'])         else '',
        'bl':       str(row['BL'])                   if pd.notna(row['BL'])           else '',
        'container':cont, 'csmId':cid,
        'rcvDate':  row['rcv_date'].strftime('%Y-%m-%d'),
        'carrier':  str(row['Carrier Name']).strip() if pd.notna(row['Carrier Name']) else '',
        'csmStat':  str(int(row['Csm Stat']))        if pd.notna(row['Csm Stat'])     else '',
        'cbm':      round(float(row['Csm VolGroTot']),2) if pd.notna(row['Csm VolGroTot']) else 0,
        'articles': articles
    })

with open('index.html', 'r', encoding='utf-8') as f:
    html = f.read()

today_str = str(date.today())
html = re.sub(r"new Date\('[0-9\-]+'\)", f"new Date('{today_str}')", html)
html = re.sub(r'const roiData = \[.*?\];',
              f"const roiData = {json.dumps(roi_json, ensure_ascii=False)};",
              html, flags=re.DOTALL)
html = re.sub(r'const weekContainers = \[.*?\];',
              f"const weekContainers = {json.dumps(week_containers, ensure_ascii=False)};",
              html, flags=re.DOTALL)

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print(f"[{today_str}] 완료 — 컨테이너 {len(week_containers)}개, 아티클 {len(roi_json)}개")
