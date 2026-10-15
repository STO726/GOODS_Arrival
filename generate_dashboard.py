import json
import re
from datetime import date

import pandas as pd


# =========================================================
# 기본 설정
# =========================================================

EXCEL_FILE = "Article search.xlsx"
HTML_FILE = "index.html"

# 엑셀 업로드 날짜와 실행 날짜가 동일하다는 전제
today = pd.Timestamp(date.today()).normalize()
week_end = today + pd.Timedelta(days=6)

today_str = today.strftime("%Y-%m-%d")
week_end_str = week_end.strftime("%Y-%m-%d")


# =========================================================
# 공통 변환 함수
# =========================================================

def safe_text(value, default=""):
    """
    NaN을 빈 문자열로 바꾸고, 나머지는 문자열로 변환한다.
    """
    if pd.isna(value):
        return default

    return str(value).strip()


def safe_int(value, default=0):
    """
    숫자 값을 안전하게 정수로 변환한다.
    """
    if pd.isna(value):
        return default

    try:
        return int(float(value))
    except (ValueError, TypeError):
        return default


def safe_float(value, default=0):
    """
    숫자 값을 안전하게 실수로 변환한다.
    """
    if pd.isna(value):
        return default

    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def safe_date_text(value):
    """
    날짜를 YYYY-MM-DD 형태로 변환한다.
    """
    if pd.isna(value):
        return ""

    parsed = pd.to_datetime(value, errors="coerce")

    if pd.isna(parsed):
        return ""

    return parsed.strftime("%Y-%m-%d")


# =========================================================
# 엑셀 시트 읽기
# =========================================================

print(f"[INFO] Excel file: {EXCEL_FILE}")
print(f"[INFO] Dashboard date: {today_str}")
print(f"[INFO] Dashboard range: {today_str} ~ {week_end_str}")

df_data = pd.read_excel(
    EXCEL_FILE,
    sheet_name="Data"
)

df_roi = pd.read_excel(
    EXCEL_FILE,
    sheet_name="ROI"
)

df_shp = pd.read_excel(
    EXCEL_FILE,
    sheet_name="SHP"
)


# =========================================================
# Data 시트 날짜 변환
# =========================================================

DATA_DATE_COLUMN = "Planned Logistics Receiving(Date)"

if DATA_DATE_COLUMN not in df_data.columns:
    raise KeyError(
        f"Data 시트에 '{DATA_DATE_COLUMN}' 열이 없습니다.\n"
        f"현재 열 목록: {list(df_data.columns)}"
    )

df_data["rcv_date"] = pd.to_datetime(
    df_data[DATA_DATE_COLUMN],
    errors="coerce"
).dt.normalize()


# =========================================================
# ROI → roi_json
# =========================================================

roi_json = []

for _, row in df_roi.iterrows():
    item_no = (
        str(safe_int(row.get("ItemNo")))
        if not pd.isna(row.get("ItemNo"))
        else ""
    )

    item_name = safe_text(row.get("ItemName"))
    pa = (
        str(safe_int(row.get("PA")))
        if not pd.isna(row.get("PA"))
        else ""
    )

    latest_rcv_date = safe_date_text(
        row.get("LatestRcvDate")
    )

    qty = safe_int(
        row.get("LatestQty")
    )

    csm_id = safe_text(
        row.get("CSM id")
    )

    ssd = safe_date_text(
        row.get("SSD")
    )

    eds = safe_date_text(
        row.get("EDS")
    )

    item_type = safe_text(
        row.get("TYPE")
    )

    roi_json.append({
        "itemNo": item_no,
        "itemName": item_name,
        "pa": pa,
        "latestRcvDate": latest_rcv_date,
        "qty": qty,
        "csmId": csm_id,
        "ssd": ssd,
        "eds": eds,
        "type": item_type,
    })


# =========================================================
# Data → CSM 매핑
# =========================================================

csm_map = {}

for _, row in df_data.iterrows():
    csm_id = safe_text(
        row.get("Csm Id")
    )

    if not csm_id:
        continue

    csm_map[csm_id] = {
        "planned": (
            row["rcv_date"].strftime("%Y-%m-%d")
            if pd.notna(row["rcv_date"])
            else ""
        ),
        "bl": safe_text(
            row.get("BL")
        ),
        "cont": safe_text(
            row.get("Container NO.")
        ),
        "carrier": safe_text(
            row.get("Carrier Name")
        ),
        "type": safe_text(
            row.get("TYPE")
        ),
        "csmStat": (
            str(safe_int(row.get("Csm Stat")))
            if not pd.isna(row.get("Csm Stat"))
            else ""
        ),
        "cbm": round(
            safe_float(row.get("Csm VolGroTot")),
            2
        ),
    }


# =========================================================
# SHP → SHP 매핑
# =========================================================

SHP_CSM_COLUMN = "Csm Id"
SHP_CW_COLUMN = "Column101"

if SHP_CSM_COLUMN not in df_shp.columns:
    raise KeyError(
        f"SHP 시트에 '{SHP_CSM_COLUMN}' 열이 없습니다.\n"
        f"현재 열 목록: {list(df_shp.columns)}"
    )

if SHP_CW_COLUMN not in df_shp.columns:
    raise KeyError(
        f"SHP 시트에 '{SHP_CW_COLUMN}' 열이 없습니다.\n"
        f"현재 열 목록: {list(df_shp.columns)}"
    )

shp_map = {}

for _, row in df_shp.iterrows():
    csm_id = safe_text(
        row.get(SHP_CSM_COLUMN)
    )

    if not csm_id:
        continue

    shp_map[csm_id] = safe_text(
        row.get(SHP_CW_COLUMN)
    )


# =========================================================
# ROI 데이터에 Data / SHP 정보 결합
# =========================================================

for item in roi_json:
    csm_id = item.get("csmId", "")

    data_info = csm_map.get(csm_id, {})

    item["plannedRcvDate"] = data_info.get(
        "planned",
        ""
    )

    item["bl"] = data_info.get(
        "bl",
        ""
    )

    item["container"] = data_info.get(
        "cont",
        ""
    )

    item["gdgy"] = shp_map.get(
        csm_id,
        ""
    )


# =========================================================
# 이번 주 컨테이너 필터링
# =========================================================

week_rows = df_data[
    (df_data["rcv_date"] >= today)
    &
    (df_data["rcv_date"] <= week_end)
].copy()

# 날짜순으로 정렬
week_rows = week_rows.sort_values(
    by=["rcv_date", "Container NO.", "Csm Id"],
    na_position="last"
)

print(f"[INFO] 이번 주 대상 행 수: {len(week_rows)}")


# =========================================================
# CSM ID별 Article 그룹 구성
# =========================================================

articles_by_csm = {}

for item in roi_json:
    csm_id = item.get("csmId", "")

    if not csm_id:
        continue

    articles_by_csm.setdefault(
        csm_id,
        []
    ).append({
        "itemNo": item.get("itemNo", ""),
        "itemName": item.get("itemName", ""),
        "qty": item.get("qty", 0),
        "type": item.get("type", ""),
    })


# =========================================================
# 이번 주 컨테이너 JSON 생성
# =========================================================

seen_containers = set()
week_containers = []

for _, row in week_rows.iterrows():

    container = safe_text(
        row.get("Container NO.")
    )

    csm_id = safe_text(
        row.get("Csm Id")
    )

    rcv_date = row.get("rcv_date")

    # 날짜가 없으면 제외
    if pd.isna(rcv_date):
        continue

    # 컨테이너 번호가 없으면 제외
    if not container:
        continue

    # 컨테이너와 날짜를 함께 기준으로 중복 제거
    container_key = (
        rcv_date.strftime("%Y-%m-%d"),
        container.upper()
    )

    if container_key in seen_containers:
        continue

    seen_containers.add(container_key)

    container_info = csm_map.get(
        csm_id,
        {}
    )

    articles = articles_by_csm.get(
        csm_id,
        []
    )

    week_containers.append({
        "type": safe_text(
            row.get("TYPE")
        ),
        "bl": safe_text(
            row.get("BL")
        ),
        "container": container,
        "csmId": csm_id,
        "rcvDate": rcv_date.strftime("%Y-%m-%d"),
        "carrier": safe_text(
            row.get("Carrier Name")
        ),
        "csmStat": (
            str(safe_int(row.get("Csm Stat")))
            if not pd.isna(row.get("Csm Stat"))
            else container_info.get("csmStat", "")
        ),
        "cbm": round(
            safe_float(row.get("Csm VolGroTot")),
            2
        ),
        "articles": articles,
    })


# =========================================================
# 검증 출력
# =========================================================

print(f"[INFO] ROI 아티클 수: {len(roi_json)}")
print(f"[INFO] 생성된 컨테이너 수: {len(week_containers)}")

container_dates = sorted(
    set(
        item["rcvDate"]
        for item in week_containers
        if item.get("rcvDate")
    )
)

print(f"[INFO] 컨테이너 날짜: {container_dates}")

today_containers = [
    item
    for item in week_containers
    if item.get("rcvDate") == today_str
]

print(
    f"[INFO] {today_str} 컨테이너 수: "
    f"{len(today_containers)}"
)

print(
    f"[INFO] {today_str} 컨테이너 목록: "
    f"{[item.get('container') for item in today_containers]}"
)


# =========================================================
# 기존 index.html 읽기
# =========================================================

with open(
    HTML_FILE,
    "r",
    encoding="utf-8"
) as f:
    html = f.read()


# =========================================================
# HTML 내부 날짜 갱신
# =========================================================

# 일반적인 new Date('YYYY-MM-DD') 갱신
html, date_replacements = re.subn(
    r"new\s+Date\s*\(\s*['\"][0-9]{4}-[0-9]{2}-[0-9]{2}['\"]\s*\)",
    f"new Date('{today_str}')",
    html
)

print(
    f"[INFO] new Date 날짜 갱신 수: "
    f"{date_replacements}"
)


# =========================================================
# DATA_AS_OF 갱신
# =========================================================

# 이전에 추가된 동적 HTML 구조가 있는 경우
# const DATA_AS_OF = maxISODate([...]);
# 를 고정된 최신 기준일로 변경한다.
html, data_as_of_replacements = re.subn(
    r"const\s+DATA_AS_OF\s*=\s*maxISODate\s*\(\s*\[.*?\]\s*\)\s*;",
    f"const DATA_AS_OF = '{today_str}';",
    html,
    flags=re.DOTALL
)

if data_as_of_replacements:
    print(
        f"[INFO] DATA_AS_OF 갱신 수: "
        f"{data_as_of_replacements}"
    )


# =========================================================
# roiData 갱신
# =========================================================

roi_json_text = json.dumps(
    roi_json,
    ensure_ascii=False,
    separators=(",", ":")
)

html, roi_replacements = re.subn(
    r"(?:const|let|var)\s+roiData\s*=\s*\[.*?\]\s*;",
    f"const roiData = {roi_json_text};",
    html,
    flags=re.DOTALL
)

if roi_replacements == 0:
    raise RuntimeError(
        "index.html에서 roiData 배열을 찾지 못했습니다. "
        "roiData 선언 형식을 확인하세요."
    )

print(
    f"[INFO] roiData 갱신 수: "
    f"{roi_replacements}"
)


# =========================================================
# weekContainers 갱신
# =========================================================

week_containers_text = json.dumps(
    week_containers,
    ensure_ascii=False,
    separators=(",", ":")
)

html, container_replacements = re.subn(
    r"(?:const|let|var)\s+weekContainers\s*=\s*\[.*?\]\s*;",
    f"let weekContainers = {week_containers_text};",
    html,
    flags=re.DOTALL
)

if container_replacements == 0:
    raise RuntimeError(
        "index.html에서 weekContainers 배열을 찾지 못했습니다. "
        "현재 HTML의 선언이 const/let/var 형식인지 확인하세요."
    )

print(
    f"[INFO] weekContainers 갱신 수: "
    f"{container_replacements}"
)


# =========================================================
# Today 정적 문구 갱신
# =========================================================

# 예:
# Today: 2026. 10. 13.
# Today: 2026.10.13.
# Today: 2026-10-13
html, today_text_replacements = re.subn(
    r"(Today\s*:\s*)(?:[0-9]{4}[.\-/]\s*[0-9]{1,2}[.\-/]\s*[0-9]{1,2}|-)",
    lambda match: (
        match.group(1)
        + today.strftime("%Y. %-m. %-d.")
    ),
    html,
    flags=re.IGNORECASE
)

print(
    f"[INFO] Today 텍스트 갱신 수: "
    f"{today_text_replacements}"
)


# =========================================================
# KPI 날짜 범위 정적 문구 갱신
# =========================================================

date_range_text = (
    f"{today.month}/{today.day} "
    f"~ "
    f"{week_end.month}/{week_end.day}"
)

# 기존 형태 예:
# 7/10 ~ 7/16
# 10/13 ~ 10/19
html, range_replacements = re.subn(
    r">[ \t]*[0-9]{1,2}/[0-9]{1,2}\s*~\s*[0-9]{1,2}/[0-9]{1,2}[ \t]*<",
    f">{date_range_text}<",
    html
)

print(
    f"[INFO] KPI 날짜 범위 갱신 수: "
    f"{range_replacements}"
)


# =========================================================
# HTML 파일 저장
# =========================================================

with open(
    HTML_FILE,
    "w",
    encoding="utf-8"
) as f:
    f.write(html)


# =========================================================
# 최종 결과 검증
# =========================================================

with open(
    HTML_FILE,
    "r",
    encoding="utf-8"
) as f:
    saved_html = f.read()

if today_str not in saved_html:
    raise RuntimeError(
        f"생성된 index.html에서 기준일 {today_str}을 "
        "찾지 못했습니다."
    )

if "weekContainers" not in saved_html:
    raise RuntimeError(
        "생성된 index.html에 weekContainers가 없습니다."
    )

print()
print("==============================================")
print("대시보드 생성 완료")
print("==============================================")
print(f"기준일              : {today_str}")
print(f"표시 종료일         : {week_end_str}")
print(f"컨테이너 수         : {len(week_containers)}")
print(f"아티클 수           : {len(roi_json)}")
print(f"HTML 날짜 교체 수   : {date_replacements}")
print(f"roiData 교체 수     : {roi_replacements}")
print(f"컨테이너 교체 수    : {container_replacements}")
print("==============================================")
