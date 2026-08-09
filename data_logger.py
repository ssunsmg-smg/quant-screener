"""
==============================================================
  data_logger.py — 장중 분봉 + 게이트 판정 결과 누적 로거

  ── v2 수정본 (2026-08-09) ─────────────────────────────────
  [수정] 시각 기준을 KST 로 고정.
      GitHub Actions 러너는 UTC 라서 datetime.now() 가 한국시간보다
      9시간 이르다. 그 결과
        - gate_log 의 timestamp 가 UTC (09:30 매수인데 00:30 으로 기록)
        - 분봉 index 는 KST 인데 로그 timestamp 는 UTC → **9시간 어긋남**
      이 상태로 데이터가 쌓이면 나중에 "게이트 통과 후 가격이 어떻게
      움직였는지" 역산할 때 전부 틀린 구간을 보게 된다.
      로그를 쌓는 목적 자체가 무너지므로 지금 고쳐둔다.
      (호출부가 date_str 을 KST 로 넘겨주긴 하지만, 이 모듈 단독으로
       써도 안전하도록 여기서도 KST 를 기준으로 삼는다)

  목적: KIS API는 과거 분봉을 제공하지 않으므로(당일만), 우리가
  직접 매 실행마다 조회한 분봉과 그 시점의 게이트 판정 결과를
  쌓아서 "진짜 데이터"를 만든다. 몇 주~몇 달 쌓이면 이 데이터로
  '게이트 통과 시점 이후 실제로 가격이 어떻게 움직였는지'를
  역산해서 게이트 로직 자체를 검증/튜닝할 수 있다.

  저장 구조:
    data/intraday/bars/{YYYYMMDD}/{code}.csv
      → 그날 그 종목의 분봉 누적 (중복 시각은 덮어쓰지 않고 dedup)
    data/intraday/gate_log_{YYYYMMDD}.jsonl
      → 매 체크 시점의 판정 스냅샷 (1줄 = 1체크 이벤트, append-only)
==============================================================
"""

import json
import os
from datetime import datetime, timedelta, timezone

import pandas as pd

# ── KST 고정 (Actions 러너는 UTC) ──
try:
    from zoneinfo import ZoneInfo
    KST = ZoneInfo("Asia/Seoul")
except Exception:
    KST = timezone(timedelta(hours=9))

DATA_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "intraday")
BARS_DIR  = os.path.join(DATA_ROOT, "bars")
GATE_LOG_TMPL = os.path.join(DATA_ROOT, "gate_log_{date}.jsonl")


def _now_kst() -> datetime:
    return datetime.now(KST)


def _today_str() -> str:
    return _now_kst().strftime("%Y%m%d")


# ══════════════════════════════════════════════════════════
# ① 분봉 누적 저장 (중복 시각 dedup)
# ══════════════════════════════════════════════════════════
def log_minute_bars(code: str, df_min: pd.DataFrame, date_str: str = None) -> None:
    """
    df_min: index=시각, 컬럼 Open/High/Low/Close/Volume
            (kis_intraday v2 부터 index 는 DatetimeIndex — 문자열 인덱스도 그대로 동작)
    같은 날 여러 번 호출돼도 겹치는 시각은 중복 없이 누적된다.
    """
    if df_min is None or df_min.empty:
        return
    date_str = date_str or _today_str()
    day_dir = os.path.join(BARS_DIR, date_str)
    os.makedirs(day_dir, exist_ok=True)
    fpath = os.path.join(day_dir, f"{code}.csv")

    # index 를 항상 문자열로 정규화해서 저장 — 그래야 CSV 로 다시 읽었을 때
    # 기존 데이터와 시각 키가 정확히 일치해 dedup 이 제대로 된다.
    new_df = df_min.copy()
    new_df.index = new_df.index.astype(str)
    new_df.index.name = "time"

    if os.path.exists(fpath):
        try:
            old_df = pd.read_csv(fpath, dtype={"time": str}).set_index("time")
            combined = pd.concat([old_df, new_df])
            combined = combined[~combined.index.duplicated(keep="last")]
            combined = combined.sort_index()
        except Exception as e:
            print(f"  ⚠ [데이터로거] {code} 기존 파일 병합 실패({e}) → 새 데이터로 덮어씀")
            combined = new_df
    else:
        combined = new_df

    combined.to_csv(fpath, index=True, index_label="time")


# ══════════════════════════════════════════════════════════
# ② 게이트 판정 스냅샷 로그 (jsonl, append-only)
# ══════════════════════════════════════════════════════════
def log_gate_check(code: str, name: str, gate_result: dict,
                   current_price: float, date_str: str = None) -> None:
    """
    매 체크 시점의 판정 결과를 한 줄(jsonl)로 누적.
    나중에 이 로그의 timestamp + code로 data/intraday/bars/에서
    그 이후 가격 흐름을 찾아 forward return을 계산하는 식으로 검증한다.

    ★ timestamp 는 KST 다 (분봉 index 와 같은 기준). 이게 어긋나면
      forward return 계산이 통째로 틀어진다.
    """
    date_str = date_str or _today_str()
    os.makedirs(DATA_ROOT, exist_ok=True)
    path = GATE_LOG_TMPL.format(date=date_str)

    detail = gate_result.get("detail") or {}
    record = {
        "timestamp": _now_kst().strftime("%Y-%m-%d %H:%M:%S"),
        "tz": "Asia/Seoul",
        "code": code,
        "name": name,
        "current_price": current_price,
        "gate_pass": gate_result.get("pass"),
        # 봉 부족으로 '판정 불가'였던 건지, 지표가 나빠서 '탈락'한 건지 구분해서 남긴다.
        # 나중에 백테스트할 때 이 둘을 섞으면 통계가 오염된다.
        "insufficient": bool(detail.get("insufficient", False)),
        "checks": gate_result.get("checks"),
        "detail": detail,
    }
    if "multi_detail" in gate_result:
        record["multi_detail"] = gate_result["multi_detail"]

    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False, default=str) + "\n")


# ══════════════════════════════════════════════════════════
# ③ 누적 데이터 현황 확인 (지금까지 며칠치 쌓였는지)
# ══════════════════════════════════════════════════════════
def status_summary() -> dict:
    if not os.path.isdir(BARS_DIR):
        return {"days": 0, "dates": [], "total_stock_day_files": 0}
    dates = sorted(d for d in os.listdir(BARS_DIR) if os.path.isdir(os.path.join(BARS_DIR, d)))
    total_files = sum(len(os.listdir(os.path.join(BARS_DIR, d))) for d in dates)
    return {"days": len(dates), "dates": dates, "total_stock_day_files": total_files}


if __name__ == "__main__":
    s = status_summary()
    print(f"누적 일수: {s['days']}일")
    print(f"날짜 목록: {s.get('dates', [])}")
    print(f"종목×일 파일 수: {s.get('total_stock_day_files', 0)}")
