"""
rag_text.py — RAG 임베딩용 텍스트 정규화 (rag_embed.py / rag_search.py 공용)

임베딩 생성과 검색 쿼리가 반드시 같은 텍스트 구성을 써야 하므로 한 곳에 둔다.
변경 시 rag_embed.py로 전체 재임베딩 필요.

정규화 규칙:
  1. 괄호 안 범위 설명 제거:  "BatVolt (0~18V)" → "BatVolt"
  2. 타입 접두어 제거:        "u1_IgnOnStaChk" → "IgnOnStaChk"
  3. camelCase / snake_case 단어 분리: "GearPosSta" → "Gear Pos Sta"
     (MiniLM 토크나이저가 붙은 복합어를 잘 못 쪼개므로 명시적으로 분리)
  4. 노이즈 토큰 제거:        CtAp/CtDcm 컴포넌트 접두어, I/P/B 방향 표시 등
  5. failure_mode를 뒤에 붙임 (없으면 ANY)
"""

import re

# ARXML De/Dg 계열 접두어 (pipeline.normalize_varname과 동일한 세트)
_DE_PREFIXES = {"de", "dg", "di", "dv", "dp", "dm"}
# 의미 없는 토큰 (컴포넌트 접두어, 포트 방향, 타입 표기, 일반어)
_NOISE_TOKENS = {"i", "p", "b", "o", "ct", "ap", "ctap", "ctdcm", "raw",
                 "sig", "msg", "message"}

# 자동차 SW 신호명 약어 → 전체 단어 (관찰된 어휘만, 애매한 것은 제외)
# 같은 신호가 프로젝트마다 다른 약어로 표기되는 문제를 흡수:
#   BDC02MsgTo / BDC05MsgTimeout / BDC_05_Timeout → 모두 "BDC 0x Timeout"
_ABBREV = {
    "to": "timeout", "tout": "timeout",
    "sta": "status", "flt": "fault", "volt": "voltage",
    "ign": "ignition", "idt": "indicator", "lvr": "lever",
    "pos": "position", "snr": "sensor", "chk": "check",
    "trmnl": "terminal", "ctrl": "control", "grp": "group",
    "wrng": "warning", "lmp": "lamp", "bat": "battery",
    "pwr": "power", "cnt": "counter", "spd": "speed",
    "btn": "button", "vhcl": "vehicle", "mot": "motor",
    "cmd": "command", "req": "request", "err": "error",
}


def _camel_split(s: str) -> list[str]:
    """camelCase / PascalCase / snake_case → 단어 리스트"""
    s = re.sub(r"[_\-]+", " ", s)
    s = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", s)          # posSta → pos Sta
    s = re.sub(r"(?<=[A-Z])(?=[A-Z][a-z])", " ", s)         # SBWSig → SBW Sig
    s = re.sub(r"(?<=[A-Za-z])(?=\d)", " ", s)              # BDC02 → BDC 02
    return [w for w in s.split() if w]


def build_embed_text(variable_name: str, failure_mode: str | None) -> str:
    """임베딩할 텍스트 구성: 정규화 단어들 + 소문자 앵커 + 고장 모드

    앵커(canonical) 토큰: 정규화 후 단어를 소문자로 붙인 형태.
    표기만 다른 동일 신호(SactSig/SActSig, BDC02Timeout/BDC_02_Timeout)가
    같은 앵커를 공유하게 되어 exact-match를 최상위로 끌어올린다.
    tail 앵커(마지막 2단어)는 컴포넌트 경로가 붙은 긴 이름과
    핵심 신호명만 있는 이름(CtApSBWSigSet_..._IdtSta ↔ DeIdtSta)을 이어준다.
    """
    base = variable_name.split("(")[0].strip()

    # 타입 접두어는 분리 전에 제거 (u1 → "u 1"로 쪼개지는 것 방지)
    base = re.sub(r"(?:^|_)u\d+_", "_", base).strip("_")

    words = _camel_split(base)

    kept = []
    for w in words:
        lw = w.lower()
        if lw in _NOISE_TOKENS:
            continue
        if re.fullmatch(r"u\d+", lw):
            continue  # 잔여 타입 토큰 (u1, u8 등). 순수 숫자는 메시지 번호이므로 보존
        kept.append(w)

    # De/Dg 접두어가 첫 단어에 붙어있는 경우 (DeIdtSta → De Idt Sta)
    if len(kept) >= 2 and kept[0].lower() in _DE_PREFIXES:
        kept = kept[1:]

    # 낱글자 토큰은 다음 단어에 병합 (SActSig → S Act Sig → SAct Sig)
    # 안 하면 SactSig(→Sact)와 임베딩이 벌어짐
    merged, buf = [], ""
    for w in kept:
        if len(w) == 1 and w.isalpha():
            buf += w
            continue
        merged.append(buf + w if buf else w)
        buf = ""
    if buf:
        if merged:
            merged[-1] += buf
        else:
            merged = [buf]
    kept = merged

    if not kept:
        kept = words or [base]

    # 약어를 전체 단어로 확장한 canonical 형태 —
    # 표기(대소문자·언더스코어)와 약어 수준이 달라도 같은 신호면 같은 canonical
    canon = [_ABBREV.get(w.lower(), w.lower()) for w in kept]

    full_anchor = "".join(canon)
    tail_anchor = "".join(canon[-2:]) if len(canon) >= 2 else full_anchor

    parts = [" ".join(w.capitalize() for w in canon), full_anchor]
    if tail_anchor != full_anchor:
        parts.append(tail_anchor)

    mode = failure_mode or "ANY"
    return f"{' '.join(parts)} {mode}"


if __name__ == "__main__":
    # 정규화 동작 확인
    samples = [
        ("BDC02Timeout", "NO"),
        ("BDC_02_Timeout", "NO"),
        ("CtApSBWSigSet_I_u1_IdtSta", "NO"),
        ("DeIdtSta", "NO"),
        ("u1_IgnOnStaChk", "LATE"),
        ("IgnOnStaChk", "LATE"),
        ("BatVolt (0~18V)", "MORE"),
        ("GearPosSta", "CORRUPT"),
    ]
    for vn, fm in samples:
        print(f"{vn:35s} → {build_embed_text(vn, fm)}")
