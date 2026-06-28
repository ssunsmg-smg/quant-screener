"""
==============================================================
  quant_signal_check.py — 장중 L1(가격)+L2(거래량) 시그널 체크
  → 통과한 종목만 KIS 매수 실행

  실행 흐름:
    1) 전날 저녁 스크리닝(quant_daily.yml)이 만든 매매신호_KR_*.json 로드
       (이미 L3 수급 필터까지 반영된 "매수 후보" top_n)
    2) 후보 중 "오늘 아직 안 산" 종목만 골라 분봉 데이터로 L1+L2 게이트 체크
    3) 게이트 통과 → 매수 실행 (기존 KISAutoTrader.place_order 재사용)
       게이트 실패 → pending 상태 유지, 다음 실행(예: 10분 후)에 재시도
    4) 마감 컷오프 시각 이후엔 더 이상 신규 매수 시도 안 함

  실행 주기: VPS cron → cron-job.org 또는 GitHub Actions workflow_dispatch
            (장중 09:05~15:00 사이 10~15분 간격을 권장 — 너무 짧으면 API
             rate limit/비용 부담, 너무 길면 타이밍 의미 희석)

  사용법:
    python quant_signal_check.py --output-dir .
==============================================================
"""

import argparse
import json
import os
import sys
import time
from datetime import datetime

import pandas as pd

# 기존 스크리너 파일과 같은 폴더에 있어야 함
from quant_screener_v41f import (
    BASE_DIR, TRADE_DIR,
    _find_latest_signal_json, _load_signal_json_as_df,
)
from kis_intraday import KISIntraday
from signal_engine import evaluate_entry_gate, evaluate_entry_gate_multi, DEFAULT_MIN_CANDLES
import data_logger

PENDING_PATH_TMPL = os.path.join(TRADE_DIR, "intraday_pending_{date}.json")
MAX_TRIES_PER_STOCK = 30          # 하루 최대 재시도 횟수 (너무 오래 들고 있지 않도록)
CUTOFF_HOUR_MIN = (14, 50)        # 이 시각 이후엔 신규 매수 시도 중단 (장 마감 대비 여유)

# ── 분봉 게이트 설정 ──
# GATE_INTERVALS: 체크할 분봉 목록. 기본은 기존과 동일하게 1분봉 단독.
#   --multi-tf 옵션을 주면 5/10/15/30/60분봉도 같이 평가해서 로그에 남기고,
#   GATE_MODE에 따라 실제 매수 판단에도 반영한다.
# GATE_MODE: "single"      = 1분봉 결과만으로 매수 판단 (기존 동작, 기본값)
#            "all_pass"    = 평가 가능한 모든 분봉이 전부 통과해야 매수 (가장 엄격)
#            "majority"    = 평가 가능한 분봉 중 과반수가 통과하면 매수
GATE_INTERVALS_DEFAULT = (1, 5, 10, 15, 30, 60)

# ── 수동 관리하는 한국 휴장일 목록 (2026년, KRX 공식 일정 기준) ──
# 형식: "YYYYMMDD". 주말은 별도로 자동 체크하니 여기엔 주중 공휴일/임시휴장일만 추가.
# ⚠ 출처는 2차 정리 자료라, 실거래 전에 한국거래소(KRX) 공식 공지로 한 번 더 대조 권장.
#   연도가 바뀌면 이 목록도 매년 갱신해야 함.
KRX_HOLIDAYS = {
    "20260101",  # 신정
    "20260216",  # 설날
    "20260217",  # 설날
    "20260218",  # 설날
    "20260302",  # 삼일절 대체휴일
    "20260501",  # 근로자의날
    "20260505",  # 어린이날
    "20260525",  # 석가탄신일 대체휴일
    "20260603",  # 임시공휴일
    "20260717",  # 제헌절 (※ 실제로는 비거래일 아닌 경우도 있어 KRX 공지로 재확인 권장)
    "20260817",  # 광복절 대체휴일
    "20260924",  # 추석
    "20260925",  # 추석
    "20261005",  # 개천절 대체휴일
    "20261009",  # 한글날
    "20261225",  # 성탄절
    "20261231",  # 연말휴장일
}


def _is_trading_day(now: datetime = None) -> bool:
    """주말(토/일) + 수동 휴장일 목록 기준으로 '오늘이 실제 거래일인지' 판단.
    ⚠ KIS 분봉 API는 휴장일에도 직전 거래일 데이터를 '정상 데이터'처럼 돌려주므로
      (빈 데이터로 휴장 여부를 판단할 수 없음이 실측으로 확인됨),
      반드시 이 함수로 사전에 막아야 한다."""
    now = now or datetime.now()
    if now.weekday() >= 5:   # 5=토, 6=일
        return False
    if now.strftime("%Y%m%d") in KRX_HOLIDAYS:
        return False
    return True


def _today_str() -> str:
    return datetime.now().strftime("%Y%m%d")


def _load_pending(date_str: str) -> dict:
    path = PENDING_PATH_TMPL.format(date=date_str)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"date": date_str, "candidates": {}}


def _save_pending(state: dict, date_str: str):
    path = PENDING_PATH_TMPL.format(date=date_str)
    os.makedirs(TRADE_DIR, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def _past_cutoff() -> bool:
    now = datetime.now()
    return (now.hour, now.minute) >= CUTOFF_HOUR_MIN


def _already_held(trader: KISIntraday, code: str) -> bool:
    """메모리(self.positions)가 아니라 실제 계좌 잔고를 직접 조회해 판단."""
    bal = trader.get_balance()
    for h in bal.get("holdings", []):
        if str(h.get("pdno", "")) == str(code) and int(h.get("hldg_qty", 0) or 0) > 0:
            return True
    return False


def run(args):
    date_str = _today_str()
    print(f"\n  [시그널체크] {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 실행")

    if not _is_trading_day():
        print(f"  🚫 오늘({datetime.now().strftime('%Y-%m-%d (%a)')})은 거래일이 아닙니다 — 매수 시도 안 함")
        print("     (KIS 분봉 API는 휴장일에도 직전 거래일 데이터를 정상처럼 돌려주므로,")
        print("      이 가드 없이는 휴장일 데이터로 잘못 매수할 위험이 있음)")
        return

    if _past_cutoff():
        print(f"  ⏰ 컷오프 시각({CUTOFF_HOUR_MIN[0]:02d}:{CUTOFF_HOUR_MIN[1]:02d}) 이후 — 신규 매수 시도 안 함")
        return

    json_path = _find_latest_signal_json(args.output_dir)
    if not json_path:
        print(f"  ⚠ {args.output_dir} 안에 매매신호_KR_*.json 파일이 없습니다. (전날 스크리닝 먼저 필요)")
        return

    df_top = _load_signal_json_as_df(json_path)
    if df_top.empty:
        print("  ⚠ 신호 파일에 유효 후보 없음")
        return

    sig_series = df_top["매매시그널"].astype(str)
    buy_candidates = df_top[sig_series.str.contains("매수", na=False)]
    if buy_candidates.empty:
        print("  ⚠ 매수 시그널 종목 없음")
        return

    state = _load_pending(date_str)
    cands_state = state["candidates"]

    trader = KISIntraday()
    if not trader._is_configured():
        print("  ⚠ KIS 앱키 미설정 → 중단")
        return

    bal = trader.get_balance()
    total_cap = bal.get("순자산", 10_000_000)
    base_amt = trader.cfg.get("base_invest_amount", 10_000_000)
    buy_top_n = trader.cfg.get("buy_top_n", 20)

    pending_codes = [c for c in buy_candidates.index if str(c) not in cands_state or
                     cands_state.get(str(c), {}).get("status") == "pending"]

    print(f"  [시그널체크] 매수 후보 {len(buy_candidates)}종목 중 미해결 {len(pending_codes)}종목 체크")

    bought_today = [c for c, v in cands_state.items() if v.get("status") == "bought"]
    remaining_budget_targets = max(1, buy_top_n - len(bought_today))
    per_stock_budget = (trader._reinvest_pool or base_amt) / remaining_budget_targets

    results = []
    for code in pending_codes:
        code = str(code)
        rec = cands_state.get(code, {"status": "pending", "tries": 0, "first_seen": date_str})

        if rec.get("tries", 0) >= MAX_TRIES_PER_STOCK:
            rec["status"] = "expired"
            cands_state[code] = rec
            continue

        if _already_held(trader, code):
            rec["status"] = "bought"
            cands_state[code] = rec
            print(f"  ℹ {code} 이미 보유 중 → 스킵")
            continue

        if args.multi_tf:
            charts = trader.get_minute_chart_multi(code, intervals=GATE_INTERVALS_DEFAULT)
            time.sleep(0.4)
            multi = evaluate_entry_gate_multi(charts, direction="BUY")
            if args.gate_mode == "all_pass":
                gate_pass = multi["all_pass"]
            elif args.gate_mode == "majority":
                gate_pass = multi["majority_pass"]
            else:  # "single" — 1분봉 결과만으로 판단 (멀티는 로그만)
                gate_pass = multi["by_interval"].get(1, {}).get("pass", False)
            gate = multi["by_interval"].get(1, {"pass": gate_pass, "checks": {}, "detail": {}})
            gate["pass"] = gate_pass   # 위에서 정한 모드 기준으로 최종 통과여부 덮어씀
            gate["multi_detail"] = {
                "mode": args.gate_mode,
                "passed_intervals": multi["passed_intervals"],
                "failed_intervals": multi["failed_intervals"],
                "pass_count": f"{multi['pass_count']}/{multi['total_count']}",
            }
            df_min = charts.get(1, pd.DataFrame())   # 로깅용 (기존 코드와의 호환)
        else:
            df_min = trader.get_minute_chart(code, interval=args.interval, lookback_calls=2)
            time.sleep(0.4)   # KIS rate limit 여유
            gate = evaluate_entry_gate(df_min, direction="BUY",
                                        min_candles=DEFAULT_MIN_CANDLES.get(args.interval, 10))
        rec["tries"] = rec.get("tries", 0) + 1
        rec["last_check"] = datetime.now().strftime("%H:%M:%S")
        rec["last_detail"] = gate.get("detail", {})

        # ── 분봉 + 게이트 판정 누적 로깅 (백테스트용 데이터 축적, B안) ──
        # 매수 여부와 무관하게 "체크한 모든 시점"을 남겨야 나중에
        # "통과했다면 어떻게 됐을지" / "탈락했는데 사실 올랐는지"를 다 검증할 수 있다.
        row_name = str(buy_candidates.loc[code].get("종목명", "")) if code in buy_candidates.index else ""
        approx_price = float(df_min["Close"].iloc[-1]) if not df_min.empty else 0.0
        data_logger.log_minute_bars(code, df_min, date_str)
        data_logger.log_gate_check(code, row_name, gate, approx_price, date_str)

        if gate["pass"]:
            row = buy_candidates.loc[code]
            cur = trader.get_current_price(code)
            time.sleep(0.4)
            if cur <= 0:
                rec["status"] = "pending"
                cands_state[code] = rec
                continue

            qty = max(1, int(per_stock_budget / cur))
            reason = (f"L1+L2게이트 통과 | VWAP:{gate['detail'].get('vwap'):.0f} "
                      f"CMF:{gate['detail'].get('cmf')} | {row.get('매매시그널','')}")
            r = trader.place_order(code, "BUY", qty, reason=reason)
            if r.get("success"):
                rec["status"] = "bought"
                rec["buy_price"] = r.get("price")
                rec["buy_time"] = datetime.now().strftime("%H:%M:%S")
                results.append({"code": code, "name": row.get("종목명", ""), "action": "BUY",
                                 "price": r.get("price"), "qty": qty})
                print(f"  ✅ {code} 게이트 통과 → 매수 {qty}주 @{r.get('price'):,}원")
            else:
                rec["status"] = "pending"
                print(f"  ⚠ {code} 게이트 통과했지만 주문 실패: {r.get('msg','')}")
        else:
            failed = [k for k, v in gate["checks"].items() if not v]
            extra = ""
            if args.multi_tf:
                md = gate.get("multi_detail", {})
                extra = (f" | 멀티분봉({md.get('mode')}): 통과 {md.get('pass_count')} "
                         f"통과분봉={md.get('passed_intervals')} 실패분봉={md.get('failed_intervals')}")
            print(f"  ⏳ {code} 게이트 미통과 (실패: {', '.join(failed)}) — {rec['tries']}회차, 재시도 대기{extra}")

        cands_state[code] = rec

    state["candidates"] = cands_state
    _save_pending(state, date_str)

    bought_n = sum(1 for v in cands_state.values() if v.get("status") == "bought")
    pending_n = sum(1 for v in cands_state.values() if v.get("status") == "pending")
    expired_n = sum(1 for v in cands_state.values() if v.get("status") == "expired")
    print(f"\n  [시그널체크] 완료 — 매수:{bought_n} 대기:{pending_n} 만료:{expired_n} "
          f"(이번 실행 신규매수: {len(results)}건)")

    data_status = data_logger.status_summary()
    print(f"  [데이터누적] 지금까지 {data_status['days']}거래일치 분봉/게이트 로그 축적됨 "
          f"(종목×일 파일 {data_status.get('total_stock_day_files', 0)}개)")

    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="장중 L1+L2 시그널 게이트 체크 → KIS 매수")
    parser.add_argument("--output-dir", type=str, default=BASE_DIR)
    parser.add_argument("--interval", type=int, choices=[1, 3, 5, 10, 15, 30, 60], default=5,
                         help="--multi-tf를 안 쓸 때 기준으로 삼을 분봉 주기 (기본: 5분)")
    parser.add_argument("--multi-tf", action="store_true",
                         help="기준 분봉(--interval) 단독 대신 1/5/10/15/30/60분봉을 모두 같이 평가 (기본: 끔)")
    parser.add_argument("--gate-mode", choices=["single", "all_pass", "majority"], default="single",
                         help="--multi-tf 켰을 때 최종 매수판단 기준. "
                              "single=1분봉 결과만 사용(멀티는 로그용), "
                              "all_pass=평가된 분봉 전부 통과해야 매수, "
                              "majority=과반수 통과시 매수 (기본: single)")
    args = parser.parse_args()
    run(args)
