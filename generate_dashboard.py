import pandas as pd
import json, re
from datetime import date

today    = pd.Timestamp(date.today())
week_end = today + pd.Timedelta(days=6)

df_data = pd.read_excel('Article search.xlsx', sheet_name='Data')
df_roi  = pd.read_excel('Article search.xlsx', sheet_name='ROI')
df_shp  = pd.read_excel('Article search.xlsx', sheet_name='SHP', usecols=['CSM id', 'CW'])  # ✅ 추가

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

# ✅ SHP → GD/GY 매핑 딕셔너리
shp_map = {}
for _, row in df_shp.iterrows():
    cid = str(row['CSM id']).strip() if pd.notna(row['CSM id']) else ''
    cw  = str(row['CW']).strip()     if pd.notna(row['CW'])     else ''
    if cid:
        shp_map[cid] = cw

for item in roi_json:
    m = csm_map.get(item['csmId'], {})
    item['plannedRcvDate'] = m.get('planned', '')
    item['bl']             = m.get('bl', '')
    item['container']      = m.get('cont', '')
    item['gdgy']           = shp_map.get(item['csmId'], '')  # ✅ 추가

# 이하 기존 코드 동일 (week_containers, HTML 치환 등 변경 없음)
