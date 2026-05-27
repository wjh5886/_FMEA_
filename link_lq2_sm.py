"""
LQ2 Safety Mechanism 연결
detection_action 텍스트에서 [SMxx] 패턴 파싱 →
safety_mechanism_text, safety_mechanism_id 업데이트
"""
import sys, re, requests, urllib3
urllib3.disable_warnings()
sys.stdout.reconfigure(encoding='utf-8')

URL      = "https://itzgdbeiyvodhfhmvrfw.supabase.co"
SUPA_KEY = ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Iml0"
            "emdkYmVpeXZvZGhmaG12cmZ3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3NDE2MzcsImV4cCI"
            "6MjA5MjMxNzYzN30.iqzQr-3Lqf1O4UHFe9euTLyeIyBXreLPoSzUdtEaNP8")
H        = {"apikey": SUPA_KEY, "Authorization": f"Bearer {SUPA_KEY}", "Content-Type": "application/json"}
PROJ_ID  = "89dc5818-2435-4d09-a1a9-36aea664d11d"

# SM ID → DB UUID 맵
sm_rows = requests.get(
    f"{URL}/rest/v1/safety_mechanisms?project_id=eq.{PROJ_ID}&select=id,sm_id,name",
    headers=H, verify=False
).json()
sm_map = {r['sm_id']: r for r in sm_rows}   # "SM01" → {id, sm_id, name}
print(f"SM 목록: {list(sm_map.keys())}")

# detection_action에 SM 텍스트가 있는 항목 로드
items = []
offset = 0
while True:
    r = requests.get(
        f"{URL}/rest/v1/fmea_items"
        f"?project_id=eq.{PROJ_ID}"
        f"&safety_mechanism_id=is.null"
        f"&detection_action=not.is.null"
        f"&select=id,detection_action"
        f"&offset={offset}&limit=1000",
        headers=H, verify=False
    )
    batch = r.json()
    if not batch: break
    items.extend(batch)
    offset += len(batch)
    if len(batch) < 1000: break

print(f"detection_action 있는 항목: {len(items)}개")

# SM 패턴 파싱: "1. MCU Safety Mechanism(Memory) - ..." or "[SM08]" or "SM08"
def find_sm_in_text(text: str):
    """detection_action 텍스트에서 가장 관련성 높은 SM ID 찾기"""
    if not text: return None

    # 패턴 1: [SM01] ~ [SM23]
    m = re.search(r'\[SM(\d+)\]', text)
    if m:
        return f"SM{m.group(1).zfill(2)}"

    # 패턴 2: "MCU Safety Mechanism(Memory)" → SM08
    keyword_map = {
        'watchdog':           'SM01',
        'hw watchdog':        'SM01',
        'power supply.*#1':   'SM02',
        'battery voltage':    'SM03',
        'redundancy can':     'SM04',
        'can monitoring':     'SM05',
        'clock':              'SM07',
        'memory':             'SM08',
        'ign signal.*redund': 'SM09',
        'power supply.*#2':   'SM10',
        'voltage.*monitor':   'SM11',
        r'\bcrc\b':           'SM12',
        'alive counter':      'SM13',
        'stuck fault':        'SM15',
        'vcu crc':            'SM17',
        'flow control':       'SM19',
        'ign power.*redund':  'SM23',
    }
    text_lower = text.lower()
    for pattern, sm_id in keyword_map.items():
        if re.search(pattern, text_lower):
            return sm_id

    return None

ok = skip = no_sm = 0
sm_count = {}

for it in items:
    sm_id = find_sm_in_text(it['detection_action'])
    if not sm_id:
        no_sm += 1
        continue

    sm_row = sm_map.get(sm_id)
    if not sm_row:
        skip += 1
        continue

    payload = {
        "safety_mechanism_id":   sm_row['id'],
        "safety_mechanism_text": sm_row['sm_id'],  # "SM08" 형태
    }
    r2 = requests.patch(
        f"{URL}/rest/v1/fmea_items?id=eq.{it['id']}",
        json=payload, headers=H, verify=False
    )
    if r2.status_code < 300:
        ok += 1
        sm_count[sm_id] = sm_count.get(sm_id, 0) + 1
    else:
        print(f"  오류: {r2.text[:60]}")
        skip += 1

print(f"\nSM 연결 완료: {ok}개 / 매칭없음: {no_sm}개 / 오류: {skip}개")
print("\nSM별 연결 수:")
for sm_id, cnt in sorted(sm_count.items()):
    name = sm_map.get(sm_id, {}).get('name', '')
    print(f"  {sm_id} ({name}): {cnt}개")

# 최종 통계
final = requests.get(
    f"{URL}/rest/v1/fmea_items?project_id=eq.{PROJ_ID}"
    f"&select=safety_mechanism_id&limit=3000",
    headers=H, verify=False
).json()
linked = sum(1 for r in final if r.get('safety_mechanism_id'))
print(f"\n전체 SM 연결: {linked}/{len(final)}개")
