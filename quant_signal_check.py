"""
==============================================================
  quant_signal_check.py — 장중 L1(가격)+L2(거래량) 시그널 체크
  → 통과한 종목만 KIS 매수 실행

  ── v2 수정본 (2026-08-09) ─────────────────────────────────
  [수정 1] ★ 신호파일 경로 문제 ★
      quant_daily.yml 은 매매신호_KR_*.json 을 `signals/` 폴더에만
      커밋한다(레포 루트의 원본은 러너와 함께 사라짐).
      그런데 quant_signal.yml 은 `--output-dir .` 로 **루트**를 뒤졌다.
      → 항상 "파일 없음" 으로 조용히 return, 종료코드 0, 초록불.
      이게 장중 매수가 한 번도 실행되지 않은 진짜 원인이다.
      이제 --output-dir 뿐 아니라 signals/ 도 자동으로 탐색하고,
      그래도 못 찾으면 **exit 1 로 빨간불**을 낸다.

  [수정 2] 시각 기준을 KST 로 고정.
      GitHub Actions 러너는 UTC. _past_cutoff() 가 14:50 **UTC**
      (= 23:50 KST) 기준으로 동작해서 장 마감 컷오프 가드가
      사실상 꺼져 있었다.

  [수정 3] 설정 오류를 조용히 넘기지 않는다.
      KIS 앱키 미설정 → 기존엔 print 후 return(종료코드 0).
      워크플로우가 초록불이라 2개월간 아무도 몰랐다. → exit 1.

  [수정 4] 재투자 자금 풀 차감·저장.
      기존엔 풀에서 예산만 읽고 **쓴 만큼 차감하지 않아서**,
      10분 뒤 실행과 10시 quant_trade.yml 이 같은 자금을 중복으로
      배정할 수 있었다(초과매수 위험).

  [수정 5] pending 상태 원자적 저장 + 종목별 예외 격리.
      한 종목에서 예외가 나도 나머지 종목은 계속 처리한다.

  실행 흐름:
    1) 전날 저녁 스크리닝(quant_daily.yml)이 만든 매매신호_KR_*.json 로드
    2) 후보 중 "오늘 아직 안 산" 종목만 골라 분봉 데이터로 L1+L2 게이트 체크
    3) 게이트 통과 → 매수 실행 (기존 KISAutoTrader.place_order 재사용)
       게이트 실패 → pending 상태 유지, 다음 실행(예: 10분 후)에 재시도
    4) 마감 컷오프 시각 이후엔 더 이상 신규 매수 시도 안 함

  사용법:
    python quant_signal_check.py --output-dir signals
==============================================================
"""

import argparse
import json
import os
import sys
import time
import traceback
from datetime import datetime, timedelta, timezone

import pandas as pd

# 기존 스크리너 파일과 같은 폴더에 있어야 함
from quant_screener_v41f import (
    BASE_DIR, TRADE_DIR,
    _find_latest_signal_json, _load_signal_json_as_df,
)
from kis_intraday import KISIntraday
from signal_engine import evaluate_entry_gate
import data_logger

# ── signal_engine 버전 차이 흡수 ─────────────────────────────
# [수정 6] 레포의 signal_engine.py 에는 evaluate_entry_gate_multi 가
#   존재하지 않아 모듈 import 단계에서 ImportError 로 죽었다.
#   (kis_intraday.get_minute_chart_multi 가 없던 것과 같은 패턴 —
#    호출하는 쪽만 있고 실제 함수는 작성되지 않음)
#   기본 동작(single 모드)은 이 함수가 필요 없으므로, 없으면 없는 대로
#   두고 --multi-tf 를 실제로 켰을 때만 크게 실패시킨다.
#   ※ "없으면 조용히 넘어간다"가 아니라 "쓰려고 할 때 확실히 막는다"이다.
try:
    from signal_engine import evaluate_entry_gate_multi
    HAS_MULTI_TF = True
except ImportError:
    evaluate_entry_gate_multi = None
    HAS_MULTI_TF = False

try:
    from signal_engine import DEFAULT_MIN_CANDLES
except ImportError:
    # 지표(VWAP/OBV/CMF) 계산에 필요한 최소 봉 개수 — 주기가 길수록 적게 요구
    DEFAULT_MIN_CANDLES = {1: 20, 3: 15, 5: 12, 10: 10, 15: 10, 30: 8, 60: 6}
    print("  ℹ signal_engine 에 DEFAULT_MIN_CANDLES 없음 → 내장 기본값 사용")

# ── KST 고정 (Actions 러너는 UTC) ──
try:
    from zoneinfo import ZoneInfo
    KST = ZoneInfo("Asia/Seoul")
except Exception:
    KST = timezone(timedelta(hours=9))


def _now() -> datetime:
    """모든 시각 판단의 단일 기준. 절대 datetime.now() 를 직접 쓰지 말 것."""
    return datetime.now(KST)


PENDING_PATH_TMPL = os.path.join(TRADE_DIR, "intraday_pending_{date}.json")
MAX_TRIES_PER_STOCK = 30          # 하루 최대 재시도 횟수 (너무 오래 들고 있지 않도록)
CUTOFF_HOUR_MIN = (14, 50)        # 이 시각(KST) 이후엔 신규 매수 시도 중단
MARKET_OPEN_HOUR_MIN = (9, 0)     # 장 시작 전에는 분봉 자체가 없으므로 실행 의미 없음

# ── 분봉 게이트 설정 ──
# GATE_INTERVALS: 체크할 분봉 목록.
#   --multi-tf 옵션을 주면 5/10/15/30/60분봉도 같이 평가해서 로그에 남기고,
#   GATE_MODE에 따라 실제 매수 판단에도 반영한다.
# GATE_MODE: "single"      = 기준 분봉 결과만으로 매수 판단 (기존 동작, 기본값)
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
    now = now or _now()
    if now.weekday() >= 5:   # 5=토, 6=일
        return False
    if now.strftime("%Y%m%d") in KRX_HOLIDAYS:
        return False
    return True


def _today_str() -> str:
    return _now().strftime("%Y%m%d")


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
    """임시파일 → rename 으로 원자적 저장 (중간에 죽어도 파일이 깨지지 않음)."""
    path = PENDING_PATH_TMPL.format(date=date_str)
    os.makedirs(TRADE_DIR, exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _past_cutoff() -> bool:
    now = _now()
    return (now.hour, now.minute) >= CUTOFF_HOUR_MIN


def _before_open() -> bool:
    now = _now()
    return (now.hour, now.minute) < MARKET_OPEN_HOUR_MIN


def _resolve_signal_json(output_dir: str) -> str:
    """
    매매신호 JSON 을 찾는다.
    ★ quant_daily.yml 이 signals/ 폴더에만 커밋하므로, output_dir 이
      루트여도 signals/ 를 반드시 함께 뒤져야 한다.
    탐색 순서: output_dir → output_dir/signals → BASE_DIR → BASE_DIR/signals
    """
    tried = []
    for d in [output_dir,
              os.path.join(output_dir, "signals"),
              BASE_DIR,
              os.path.join(BASE_DIR, "signals")]:
        if not d or d in tried:
            continue
        tried.append(d)
        if not os.path.isdir(d):
            continue
        path = _find_latest_signal_json(d)
        if path:
            print(f"  📁 신호 파일: {path}")
            return path

    print("  ❌ 매매신호_KR_*.json 을 찾지 못했습니다.")
    print(f"     탐색한 경로: {tried}")
    print("     → quant_daily.yml(저녁 스크리닝)이 먼저 성공했는지,")
    print("       signals/ 폴더가 레포에 커밋되어 있는지 확인하세요.")
    return ""


# ══════════════════════════════════════════════════════════
#  텔레그램 알림 (기존 MonitorEngine 재사용)
# ══════════════════════════════════════════════════════════
# [수정 7] 장중 시그널 워크플로우에는 텔레그램 알림이 아예 없었다.
#   quant_trade.yml 은 telegram_config.json 을 만드는데 quant_signal.yml 은
#   만들지 않았고, 이 스크립트에도 전송 코드가 없어서 장중에 매수가
#   체결돼도 폰에 아무 것도 오지 않았다. → Actions 탭을 열어봐야만 알 수 있었다.
#
# ★ 원칙: 알림 실패가 매매를 막으면 안 된다. 모든 전송은 예외를 삼킨다.
#   (텔레그램 API 장애 때문에 주문이 안 나가는 게 훨씬 나쁜 결과다)
def _make_monitor(no_telegram: bool = False):
    if no_telegram:
        return None
    try:
        from quant_screener_v41f import MonitorEngine
        m = MonitorEngine()
        return m if m.enabled else None
    except Exception as e:
        print(f"  ⚠ 텔레그램 초기화 실패 — 알림 없이 계속 진행: {e}")
        return None


def _tg(monitor, text: str) -> None:
    """절대 예외를 밖으로 던지지 않는 안전 전송."""
    if not monitor:
        return
    try:
        monitor.send(text)
    except Exception as e:
        print(f"  ⚠ 텔레그램 전송 실패(무시): {e}")


def _already_held(trader: KISIntraday, code: str, balance: dict = None) -> bool:
    """메모리(self.positions)가 아니라 실제 계좌 잔고를 직접 조회해 판단.
    balance 를 넘기면 재조회하지 않는다 (종목마다 잔고를 다시 부르면
    API 호출이 종목수만큼 늘어나 rate limit 에 걸린다)."""
    bal = balance if balance is not None else trader.get_balance()
    for h in bal.get("holdings", []):
        if str(h.get("pdno", "")) == str(code) and int(float(h.get("hldg_qty", 0) or 0)) > 0:
            return True
    return False


def run(args) -> list:
    date_str = _today_str()
    print(f"\n  [시그널체크] {_now().strftime('%Y-%m-%d %H:%M:%S')} KST 실행")

    # --multi-tf 를 켰는데 signal_engine 에 해당 함수가 없으면 즉시 중단.
    # 조용히 single 로 떨어뜨리면 "엄격한 게이트로 돌고 있다"고 착각하게 된다.
    if args.multi_tf and not HAS_MULTI_TF:
        print("  ❌ --multi-tf 를 켰지만 signal_engine.py 에 evaluate_entry_gate_multi() 가 없습니다.")
        print("     signal_engine.py 에 해당 함수를 먼저 구현하거나, --multi-tf 없이 실행하세요.")
        raise SystemExit(1)

    if not _is_trading_day():
        print(f"  🚫 오늘({_now().strftime('%Y-%m-%d (%a)')})은 거래일이 아닙니다 — 매수 시도 안 함")
        print("     (KIS 분봉 API는 휴장일에도 직전 거래일 데이터를 정상처럼 돌려주므로,")
        print("      이 가드 없이는 휴장일 데이터로 잘못 매수할 위험이 있음)")
        return []

    if _before_open():
        print(f"  🌙 장 시작({MARKET_OPEN_HOUR_MIN[0]:02d}:{MARKET_OPEN_HOUR_MIN[1]:02d} KST) 전 — 분봉 데이터 없음, 종료")
        return []

    if _past_cutoff():
        print(f"  ⏰ 컷오프 시각({CUTOFF_HOUR_MIN[0]:02d}:{CUTOFF_HOUR_MIN[1]:02d} KST) 이후 — 신규 매수 시도 안 함")
        return []

    # ── 신호 파일 (없으면 설정 오류 → 빨간불) ──
    json_path = _resolve_signal_json(args.output_dir)
    if not json_path:
        raise SystemExit(1)

    df_top = _load_signal_json_as_df(json_path)
    if df_top.empty:
        print("  ⚠ 신호 파일에 유효 후보 없음")
        return []

    sig_series = df_top["매매시그널"].astype(str)
    buy_candidates = df_top[sig_series.str.contains("매수", na=False)]
    if buy_candidates.empty:
        print("  ⚠ 매수 시그널 종목 없음")
        return []

    state = _load_pending(date_str)
    cands_state = state["candidates"]
    is_first_run_today = not cands_state      # 오늘 첫 실행인지 (감시 시작 알림용)
    monitor = _make_monitor(getattr(args, "no_telegram", False))

    trader = KISIntraday()
    if not trader._is_configured():
        # 조용히 return 하면 워크플로우가 초록불이라 아무도 모른다 → 빨간불
        print("  ❌ KIS 앱키 미설정 → 중단 (kis_config.json / GitHub Secrets 확인)")
        raise SystemExit(1)

    bal       = trader.get_balance()          # ← 한 번만 조회해서 재사용
    base_amt  = trader.cfg.get("base_invest_amount", 10_000_000)
    buy_top_n = trader.cfg.get("buy_top_n", 20)

    pending_codes = [c for c in buy_candidates.index
                     if str(c) not in cands_state
                     or cands_state.get(str(c), {}).get("status") == "pending"]

    print(f"  [시그널체크] 매수 후보 {len(buy_candidates)}종목 중 미해결 {len(pending_codes)}종목 체크")

    bought_today = [c for c, v in cands_state.items() if v.get("status") == "bought"]
    remaining_budget_targets = max(1, buy_top_n - len(bought_today))
    invest_pool = trader._reinvest_pool if trader._reinvest_pool else base_amt
    per_stock_budget = invest_pool / remaining_budget_targets
    print(f"  [예산] 가용 풀 {invest_pool:,.0f}원 / 남은 {remaining_budget_targets}종목 "
          f"→ 종목당 {per_stock_budget:,.0f}원")

    if is_first_run_today:
        _tg(monitor,
            f"🔍 <b>[한투 장중] 감시 시작</b>\n"
            f"  {_now().strftime('%m/%d %H:%M')} · 매수 후보 {len(buy_candidates)}종목\n"
            f"  💰 가용 예산 {invest_pool:,.0f}원 (종목당 {per_stock_budget:,.0f}원)\n"
            f"  ⏰ {CUTOFF_HOUR_MIN[0]:02d}:{CUTOFF_HOUR_MIN[1]:02d}까지 15분 간격으로 게이트 확인")

    results = []
    spent_won = 0.0
    errors = []

    for code in pending_codes:
        code = str(code)
        rec = cands_state.get(code, {"status": "pending", "tries": 0, "first_seen": date_str})

        try:
            if rec.get("tries", 0) >= MAX_TRIES_PER_STOCK:
                rec["status"] = "expired"
                cands_state[code] = rec
                continue

            if _already_held(trader, code, balance=bal):
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
                else:  # "single" — 기준 분봉 결과만으로 판단 (멀티는 로그만)
                    gate_pass = multi["by_interval"].get(args.interval, {}).get("pass", False)

                gate = dict(multi["by_interval"].get(args.interval,
                                                     {"pass": gate_pass, "checks": {}, "detail": {}}))
                gate["pass"] = gate_pass   # 위에서 정한 모드 기준으로 최종 통과여부 덮어씀
                gate.setdefault("checks", {})
                gate.setdefault("detail", {})
                gate["multi_detail"] = {
                    "mode": args.gate_mode,
                    "passed_intervals":  multi["passed_intervals"],
                    "failed_intervals":  multi["failed_intervals"],
                    # 봉 부족으로 평가에서 빠진 주기 (실패가 아니라 '판정 불가')
                    "skipped_intervals": multi.get("skipped", []),
                    "pass_count": f"{multi['pass_count']}/{multi['total_count']}",
                }
                df_min = charts.get(args.interval, pd.DataFrame())   # 로깅용
            else:
                df_min = trader.get_minute_chart(code, interval=args.interval)
                time.sleep(0.4)   # KIS rate limit 여유
                gate = evaluate_entry_gate(df_min, direction="BUY",
                                           min_candles=DEFAULT_MIN_CANDLES.get(args.interval, 10))

            rec["tries"] = rec.get("tries", 0) + 1
            rec["last_check"] = _now().strftime("%H:%M:%S")
            rec["last_detail"] = gate.get("detail", {})

            # ── 분봉 + 게이트 판정 누적 로깅 (백테스트용 데이터 축적, B안) ──
            # 매수 여부와 무관하게 "체크한 모든 시점"을 남겨야 나중에
            # "통과했다면 어떻게 됐을지" / "탈락했는데 사실 올랐는지"를 다 검증할 수 있다.
            row_name = str(buy_candidates.loc[code].get("종목명", "")) if code in buy_candidates.index else ""
            approx_price = float(df_min["Close"].iloc[-1]) if not df_min.empty else 0.0
            data_logger.log_minute_bars(code, df_min, date_str)
            data_logger.log_gate_check(code, row_name, gate, approx_price, date_str)

            if gate.get("pass"):
                row = buy_candidates.loc[code]
                cur = trader.get_current_price(code)
                time.sleep(0.4)
                if cur <= 0:
                    rec["status"] = "pending"
                    cands_state[code] = rec
                    continue

                qty = max(1, int(per_stock_budget / cur))
                vwap = gate.get("detail", {}).get("vwap")
                reason = (f"L1+L2게이트 통과 | VWAP:{vwap:.0f} " if isinstance(vwap, (int, float))
                          else "L1+L2게이트 통과 | ")
                reason += (f"CMF:{gate.get('detail', {}).get('cmf')} | "
                           f"{row.get('매매시그널', '')}")

                r = trader.place_order(code, "BUY", qty, reason=reason)
                if r.get("success"):
                    fill = r.get("price") or cur
                    rec["status"] = "bought"
                    rec["buy_price"] = fill
                    rec["buy_time"] = _now().strftime("%H:%M:%S")
                    spent_won += fill * qty
                    results.append({"code": code, "name": row.get("종목명", ""), "action": "BUY",
                                    "price": fill, "qty": qty,
                                    "gate_detail": gate.get("detail", {})})
                    print(f"  ✅ {code} 게이트 통과 → 매수 {qty}주 @{fill:,}원")
                else:
                    rec["status"] = "pending"
                    print(f"  ⚠ {code} 게이트 통과했지만 주문 실패: {r.get('msg','')}")
            else:
                # "봉이 모자라 판정 불가"와 "지표가 나빠서 탈락"은 전혀 다른 상태다.
                # 둘 다 '미통과'로 뭉뚱그리면 장 초반 로그를 보고 원인을 알 수 없다.
                gdetail = gate.get("detail", {})
                if gdetail.get("insufficient"):
                    why = f"판정 불가 — {gdetail.get('error', '데이터 부족')}"
                else:
                    failed = [k for k, v in gate.get("checks", {}).items() if not v]
                    why = f"실패: {', '.join(failed) if failed else '알 수 없음'}"
                extra = ""
                if args.multi_tf:
                    md = gate.get("multi_detail", {})
                    extra = (f" | 멀티분봉({md.get('mode')}): 통과 {md.get('pass_count')} "
                             f"통과분봉={md.get('passed_intervals')} 실패분봉={md.get('failed_intervals')}"
                             f" 제외분봉={md.get('skipped_intervals')}")
                print(f"  ⏳ {code} 게이트 미통과 ({why}) — "
                      f"{rec['tries']}회차, 재시도 대기{extra}")

        except Exception as e:
            # 한 종목의 예외로 나머지 종목까지 날리지 않는다
            print(f"  ⚠ {code} 처리 중 예외 — 이 종목만 건너뜁니다")
            traceback.print_exc()
            rec["status"] = "pending"
            rec["last_error"] = _now().strftime("%H:%M:%S")
            errors.append(f"{code}: {type(e).__name__} {e}")

        cands_state[code] = rec

    state["candidates"] = cands_state
    _save_pending(state, date_str)

    # ── 재투자 풀 차감 (중복 배정 방지) ──
    # 기존엔 풀에서 읽기만 하고 쓴 만큼 빼지 않아서, 10분 뒤 실행과
    # 10시 quant_trade.yml 이 같은 자금을 또 배정할 수 있었다.
    if spent_won > 0:
        new_pool = max(0.0, (trader._reinvest_pool if trader._reinvest_pool else base_amt) - spent_won)
        trader._save_reinvest_pool(new_pool, note=f"장중 시그널 매수 {spent_won:,.0f}원 집행")
        print(f"  💸 매수 집행 {spent_won:,.0f}원 → 재투자 풀 잔액 {new_pool:,.0f}원")

    bought_n = sum(1 for v in cands_state.values() if v.get("status") == "bought")
    pending_n = sum(1 for v in cands_state.values() if v.get("status") == "pending")
    expired_n = sum(1 for v in cands_state.values() if v.get("status") == "expired")
    print(f"\n  [시그널체크] 완료 — 매수:{bought_n} 대기:{pending_n} 만료:{expired_n} "
          f"(이번 실행 신규매수: {len(results)}건)")

    # ── 텔레그램: 체결 알림 ──
    # 게이트 통과 사유(VWAP/CMF)까지 같이 보낸다. "왜 샀는지"가 남아야
    # 나중에 로그를 안 뒤져도 판단을 되짚을 수 있다.
    if results:
        lines = [f"📈 <b>[한투 장중] 매수 체결</b>  {_now().strftime('%m/%d %H:%M')}"]
        for r in results:
            lines.append(f"  • {r.get('name','')} ({r['code']}) "
                         f"{r['qty']:,}주 @{r['price']:,}원 = {r['price']*r['qty']:,}원")
            d = r.get("gate_detail") or {}
            if d.get("vwap") is not None:
                lines.append(f"     VWAP {d['vwap']:,.0f} · CMF {d.get('cmf')}")
        lines.append(f"\n  💰 남은 예산: {(trader._reinvest_pool if trader._reinvest_pool else base_amt):,.0f}원")
        lines.append(f"  📊 오늘 누적 매수 {bought_n}종목 / 대기 {pending_n}종목")
        _tg(monitor, "\n".join(lines))

    # ── 텔레그램: 예외 알림 ──
    # 종목별 예외는 조용히 넘어가면 며칠씩 모르고 지나간다.
    if errors:
        _tg(monitor,
            f"⚠️ <b>[한투 장중] 처리 중 오류 {len(errors)}건</b>  {_now().strftime('%m/%d %H:%M')}\n"
            + "\n".join(f"  • {e[:120]}" for e in errors[:5])
            + ("\n  …" if len(errors) > 5 else ""))

    try:
        data_status = data_logger.status_summary()
        print(f"  [데이터누적] 지금까지 {data_status['days']}거래일치 분봉/게이트 로그 축적됨 "
              f"(종목×일 파일 {data_status.get('total_stock_day_files', 0)}개)")
    except Exception as e:
        print(f"  ⚠ 데이터 누적 상태 조회 실패: {e}")

    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="장중 L1+L2 시그널 게이트 체크 → KIS 매수")
    # ★ 기본값을 signals 로 변경 — quant_daily.yml 이 여기에만 커밋한다
    parser.add_argument("--output-dir", type=str, default="signals",
                        help="매매신호_KR_*.json 이 있는 폴더 (기본: signals). "
                             "여기서 못 찾으면 하위 signals/ 와 레포 루트도 자동 탐색")
    parser.add_argument("--interval", type=int, choices=[1, 3, 5, 10, 15, 30, 60], default=5,
                        help="게이트 판단 기준으로 삼을 분봉 주기 (기본: 5분). "
                             "KIS는 1분봉만 주므로 내부적으로 리샘플해서 만든다")
    parser.add_argument("--multi-tf", action="store_true",
                        help="기준 분봉(--interval) 단독 대신 1/5/10/15/30/60분봉을 모두 같이 평가 (기본: 끔)")
    parser.add_argument("--no-telegram", action="store_true",
                        help="텔레그램 알림 끄기 (기본: telegram_config.json 이 있으면 자동으로 켜짐)")
    parser.add_argument("--gate-mode", choices=["single", "all_pass", "majority"], default="single",
                        help="--multi-tf 켰을 때 최종 매수판단 기준. "
                             "single=기준 분봉 결과만 사용(멀티는 로그용), "
                             "all_pass=평가된 분봉 전부 통과해야 매수, "
                             "majority=과반수 통과시 매수 (기본: single)")
    args = parser.parse_args()
    run(args)
