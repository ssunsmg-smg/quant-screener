"""
==============================================================
  quant_signal_check.py — 장중 매도(손절/트레일링 익절) + 매수 게이트

  ── v3 (2026-08-09) : 매도 일원화 + 트레일링 스톱 ──────────
  이 스크립트가 이제 **매도의 유일한 주체**다.
  quant_trade.yml 쪽 매도는 QUANT_DISABLE_SELL=1 로 꺼서 중복을 없앤다.

  ★ 왜 매도를 장중으로 옮겼나
    기존: 매도 판단은 quant_trade.yml(매일 10:00) 하루 1회.
          → 13:00에 산 종목의 첫 손절 평가가 '다음날 10:00'.
            금요일 오후 매수면 월요일 10시까지 69시간 무방비.
            설정은 -7% 손절이지만 실제 동작은 "다음날 10시 가격에 판다"였다.
    수정: 15분 간격으로 매수와 같은 주기로 평가. 손절 노출 24시간 → 15분.

  ★ 트레일링 익절 (왜 고정 3% 익절이 아닌가)
    기존 조건 `tp_min <= ret <= tp_max` 는 하루 1회 체크라서 밴드처럼
    보였을 뿐, 15분 간격에서는 거의 항상 3%를 먼저 만나 팔린다.
    즉 "3~25% 익절"이 아니라 사실상 "3% 익절"이 되고,
    3% 익절 / 7% 손절은 비용 포함 손익비 0.39 → 본전 승률 72%가 필요하다.
    그래서 3%를 '매도 신호'가 아니라 '익절 감시 시작' 트리거로 쓴다:

      수익률 < 3%        → 손절선만 감시
      수익률 ≥ 3% 도달    → 감시 모드(armed) 진입, 이후 고점가 계속 추적
      고점가 대비 -2%     → 트레일링 익절 매도
      수익률 ≥ 25%       → 무조건 익절 (상한)
      수익률 ≤ -7%       → 손절

    → 원래 의도한 "3~25% 사이에서 익절"이 실제로 살아난다.
      이 방식은 15분 간격이라야 의미가 있다(하루 1회면 고점을 놓친다).

  ★ 매도는 일일 거래한도에서 제외한다
    max_daily_trades 는 신규 매수에만 건다. 한도 때문에 손절 주문이
    막히는 것이 훨씬 나쁜 결과이기 때문.

  ★ 시간 창이 매수/매도가 다르다
    매수 컷오프 14:50 / 매도 컷오프 15:20 — 마감 직전까지 팔 수는 있어야 한다.

  실행 순서:
    ① 매도 판단 (손절/트레일링 익절) → 회수금을 재투자 풀에 반영
    ② 매수 게이트 (L1+L2) → 통과 종목 매수
    (매도가 먼저인 이유: 회수금이 그날 매수 예산에 반영돼야 한다)

  사용법:
    python quant_signal_check.py --output-dir signals --interval 5
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
# 레포의 signal_engine.py 에 evaluate_entry_gate_multi / DEFAULT_MIN_CANDLES 가
# 없던 시기가 있어, import 단계에서 통째로 죽었다. 기본 동작(single)은 multi가
# 필요 없으므로 없으면 없는 대로 두고, --multi-tf 를 실제로 켰을 때만 막는다.
try:
    from signal_engine import evaluate_entry_gate_multi
    HAS_MULTI_TF = True
except ImportError:
    evaluate_entry_gate_multi = None
    HAS_MULTI_TF = False

try:
    from signal_engine import DEFAULT_MIN_CANDLES
except ImportError:
    DEFAULT_MIN_CANDLES = {1: 20, 3: 15, 5: 12, 10: 10, 15: 10, 30: 10, 60: 10}
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
PEAKS_PATH        = os.path.join(TRADE_DIR, "peaks.json")   # 트레일링용 고점 (날짜 무관, 보유하는 동안 유지)

MAX_TRIES_PER_STOCK  = 30      # 하루 최대 재시도 횟수
BUY_CUTOFF_HOUR_MIN  = (14, 50)   # 이후 신규 매수 중단
SELL_CUTOFF_HOUR_MIN = (15, 20)   # 이후 매도도 중단 (동시호가 직전)
MARKET_OPEN_HOUR_MIN = (9, 0)

# ── 매도 기본 파라미터 (kis_config.json 으로 덮어쓸 수 있음) ──
DEFAULT_STOP_LOSS_PCT    = 7.0    # 손절
DEFAULT_TRAIL_TRIGGER_PCT = 3.0   # 이 수익률 도달 시 익절 감시 시작(armed)
DEFAULT_TRAIL_DROP_PCT   = 2.0    # 고점가 대비 이만큼 하락하면 익절 매도
DEFAULT_HARD_TP_PCT      = 25.0   # 무조건 익절 상한

GATE_INTERVALS_DEFAULT = (1, 5, 10, 15, 30, 60)

# ── 수동 관리하는 한국 휴장일 목록 (2026년, KRX 공식 일정 기준) ──
# ⚠ 출처는 2차 정리 자료라, 실거래 전에 한국거래소(KRX) 공식 공지로 재대조 권장.
#   연도가 바뀌면 매년 갱신해야 함.
KRX_HOLIDAYS = {
    "20260101", "20260216", "20260217", "20260218", "20260302",
    "20260501", "20260505", "20260525", "20260603", "20260717",
    "20260817", "20260924", "20260925", "20261005", "20261009",
    "20261225", "20261231",
}


def _is_trading_day(now: datetime = None) -> bool:
    """주말 + 수동 휴장일 기준 거래일 판단.
    ⚠ KIS 분봉 API는 휴장일에도 직전 거래일 데이터를 '정상 데이터'처럼
      돌려주므로(실측 확인됨), 반드시 이 함수로 사전에 막아야 한다."""
    now = now or _now()
    if now.weekday() >= 5:
        return False
    if now.strftime("%Y%m%d") in KRX_HOLIDAYS:
        return False
    return True


def _today_str() -> str:
    return _now().strftime("%Y%m%d")


def _hm() -> tuple:
    n = _now()
    return (n.hour, n.minute)


def _before_open() -> bool:      return _hm() <  MARKET_OPEN_HOUR_MIN
def _past_buy_cutoff() -> bool:  return _hm() >= BUY_CUTOFF_HOUR_MIN
def _past_sell_cutoff() -> bool: return _hm() >= SELL_CUTOFF_HOUR_MIN


# ══════════════════════════════════════════════════════════
#  상태 파일 (pending / peaks)
# ══════════════════════════════════════════════════════════
def _load_json(path: str, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return default


def _save_json(path: str, data) -> None:
    """임시파일 → rename 으로 원자적 저장 (중간에 죽어도 파일이 깨지지 않음)."""
    os.makedirs(TRADE_DIR, exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _load_pending(date_str): return _load_json(PENDING_PATH_TMPL.format(date=date_str),
                                               {"date": date_str, "candidates": {}, "sold_today": []})
def _save_pending(state, date_str): _save_json(PENDING_PATH_TMPL.format(date=date_str), state)
def _load_peaks():  return _load_json(PEAKS_PATH, {})
def _save_peaks(p): _save_json(PEAKS_PATH, p)


# ══════════════════════════════════════════════════════════
#  텔레그램 (기존 MonitorEngine 재사용)
# ══════════════════════════════════════════════════════════
# ★ 원칙: 알림 실패가 매매를 막으면 안 된다. 모든 전송은 예외를 삼킨다.
#   (텔레그램 장애 때문에 손절이 안 나가는 게 훨씬 나쁜 결과다)
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
    if not monitor:
        return
    try:
        monitor.send(text)
    except Exception as e:
        print(f"  ⚠ 텔레그램 전송 실패(무시): {e}")


# ══════════════════════════════════════════════════════════
#  ① 매도 — 손절 / 트레일링 익절
# ══════════════════════════════════════════════════════════
def _decide_sell(ret: float, cur: float, peak_price: float, armed: bool, cfg: dict) -> tuple:
    """
    매도 판단 (순수 함수 — 단위 테스트 가능).
    반환: (sell_reason 또는 None, action, armed_new, peak_price_new)

    우선순위: 손절 > 상한 익절 > 트레일링 익절
    (손절을 먼저 보는 이유: 급락장에서 트레일링 조건과 동시에 성립할 때
     '익절'로 기록되면 나중에 성과 분석이 왜곡된다)
    """
    stop    = abs(cfg.get("stop_loss_pct",        DEFAULT_STOP_LOSS_PCT))     / 100
    trigger = abs(cfg.get("trail_trigger_pct",    DEFAULT_TRAIL_TRIGGER_PCT)) / 100
    drop    = abs(cfg.get("trail_drop_pct",       DEFAULT_TRAIL_DROP_PCT))    / 100
    hard_tp = abs(cfg.get("take_profit_max_pct",  DEFAULT_HARD_TP_PCT))       / 100

    # 고점 갱신 + 감시 모드 진입 (한 번 armed 되면 해제되지 않는다)
    peak_new  = max(peak_price, cur)
    armed_new = armed or (ret >= trigger)

    if ret <= -stop:
        return f"손절 {ret*100:+.2f}%", "STOP_LOSS", armed_new, peak_new
    if ret >= hard_tp:
        return f"익절(상한 {hard_tp*100:.0f}%) {ret*100:+.2f}%", "TAKE_PROFIT", armed_new, peak_new
    if armed_new and peak_new > 0 and cur <= peak_new * (1 - drop):
        fall = (peak_new - cur) / peak_new * 100
        # ★ 고점에서 크게 밀리면 트레일링선에 닿는 시점이 이미 '손실'일 수 있다.
        #   (예: +8%까지 갔다가 급락해 -2%에서 트레일링 발동)
        #   이걸 전부 TAKE_PROFIT 으로 기록하면 나중에 성과 집계가 오염된다
        #   — 실현손익이 마이너스인 거래가 '익절'로 세어지기 때문.
        #   그래서 실제 수익률 부호로 라벨을 나눈다. 매도 동작 자체는 동일.
        if ret > 0:
            return (f"트레일링 익절 {ret*100:+.2f}% (고점 {peak_new:,.0f} 대비 -{fall:.2f}%)",
                    "TAKE_PROFIT", armed_new, peak_new)
        return (f"트레일링 손절 {ret*100:+.2f}% (고점 {peak_new:,.0f} 대비 -{fall:.2f}%)",
                "TRAIL_STOP", armed_new, peak_new)
    return None, None, armed_new, peak_new


def _run_sell_check(trader: KISIntraday, balance: dict, monitor, date_str: str) -> tuple:
    """
    보유 종목 전체에 대해 손절/트레일링 익절 판단 → 매도 실행.
    반환: (executed_list, recovered_won, sold_codes)
    """
    holdings = [h for h in (balance.get("holdings") or [])
                if int(float(h.get("hldg_qty", 0) or 0)) > 0]
    peaks = _load_peaks()
    executed, recovered, sold_codes = [], 0.0, []

    if not holdings:
        print("  [매도점검] 보유 종목 없음")
        # 보유가 사라졌으면 고점 기록도 정리 (다음에 다시 사면 새로 시작해야 한다)
        if peaks:
            _save_peaks({})
        return executed, recovered, sold_codes

    cfg = trader.cfg
    print(f"\n  [매도점검] 보유 {len(holdings)}종목 — "
          f"손절 -{abs(cfg.get('stop_loss_pct', DEFAULT_STOP_LOSS_PCT)):.1f}% | "
          f"익절감시 +{abs(cfg.get('trail_trigger_pct', DEFAULT_TRAIL_TRIGGER_PCT)):.1f}% 도달 후 "
          f"고점 대비 -{abs(cfg.get('trail_drop_pct', DEFAULT_TRAIL_DROP_PCT)):.1f}% | "
          f"상한 +{abs(cfg.get('take_profit_max_pct', DEFAULT_HARD_TP_PCT)):.0f}%")

    # ★ 매도는 일일 거래한도에서 제외한다.
    #   한도(max_daily_trades)를 매수로 다 써서 손절 주문이 막히는 것이
    #   훨씬 나쁜 결과다. 매도 구간에서만 한도를 풀고, 끝난 뒤 되돌린다.
    saved_limit = cfg.get("max_daily_trades", 10)
    cfg["max_daily_trades"] = 10 ** 9

    try:
        for h in holdings:
            code = str(h.get("pdno", "")).strip()
            qty  = int(float(h.get("hldg_qty", 0) or 0))
            avg  = float(h.get("pchs_avg_pric", 0) or 0)
            if not code or qty <= 0 or avg <= 0:
                continue
            try:
                cur = trader.get_current_price(code)
                time.sleep(0.4)
                if cur <= 0:
                    print(f"     ⚠ {code} 현재가 조회 실패 → 이번 회차 판단 보류")
                    continue

                ret = (cur - avg) / avg
                rec = peaks.get(code, {})
                peak_price = float(rec.get("peak_price", max(cur, avg)))
                armed      = bool(rec.get("armed", False))

                reason, action, armed_new, peak_new = _decide_sell(ret, cur, peak_price, armed, cfg)
                peaks[code] = {"peak_price": peak_new, "armed": armed_new,
                               "avg_price": avg, "updated": _now().strftime("%Y-%m-%d %H:%M:%S")}

                if not reason:
                    state = "익절감시중" if armed_new else "손절선만감시"
                    print(f"     · {code} {ret*100:+6.2f}% (고점 {peak_new:,.0f}) — {state}, 보유 유지")
                    continue

                emoji = "🔴" if action == "STOP_LOSS" else "🟢"
                print(f"  {emoji} {code} {reason} → 매도 {qty:,}주")
                r = trader.place_order(code, "SELL", qty, reason=reason)
                time.sleep(0.4)
                if r.get("success"):
                    fill = r.get("price") or cur
                    value = fill * qty
                    recovered += value
                    sold_codes.append(code)
                    peaks.pop(code, None)      # 청산됐으니 고점 기록 삭제
                    executed.append({"action": action, "code": code, "qty": qty,
                                     "price": fill, "ret_pct": round(ret * 100, 2),
                                     "recovered_won": round(value), "reason": reason})
                else:
                    print(f"     ⚠ {code} 매도 주문 실패: {r.get('msg','')}")
            except Exception as e:
                # 한 종목 예외로 나머지 종목의 손절까지 막히면 안 된다
                print(f"     ⚠ {code} 매도 판단 중 예외 — 이 종목만 건너뜁니다: {e}")
                traceback.print_exc()
    finally:
        cfg["max_daily_trades"] = saved_limit
        # 매도로 소진된 카운트를 매수 한도로 넘기지 않는다
        trader.daily_trades = 0
        _save_peaks(peaks)

    return executed, recovered, sold_codes


def _already_held(code: str, balance: dict) -> bool:
    """실제 계좌 잔고 기준 보유 여부 (메모리 positions 를 믿지 않는다)."""
    for h in (balance.get("holdings") or []):
        if str(h.get("pdno", "")) == str(code) and int(float(h.get("hldg_qty", 0) or 0)) > 0:
            return True
    return False


# ══════════════════════════════════════════════════════════
#  메인
# ══════════════════════════════════════════════════════════
def run(args) -> dict:
    date_str = _today_str()
    print(f"\n  [시그널체크] {_now().strftime('%Y-%m-%d %H:%M:%S')} KST 실행")

    if args.multi_tf and not HAS_MULTI_TF:
        print("  ❌ --multi-tf 를 켰지만 signal_engine.py 에 evaluate_entry_gate_multi() 가 없습니다.")
        raise SystemExit(1)

    if not _is_trading_day():
        print(f"  🚫 오늘({_now().strftime('%Y-%m-%d (%a)')})은 거래일이 아닙니다 — 매매 안 함")
        return {"sells": [], "buys": []}
    if _before_open():
        print("  🌙 장 시작 전 — 종료")
        return {"sells": [], "buys": []}
    if _past_sell_cutoff():
        print(f"  ⏰ 매도 컷오프({SELL_CUTOFF_HOUR_MIN[0]:02d}:{SELL_CUTOFF_HOUR_MIN[1]:02d}) 이후 — 종료")
        return {"sells": [], "buys": []}

    trader = KISIntraday()
    if not trader._is_configured():
        print("  ❌ KIS 앱키 미설정 → 중단 (kis_config.json / GitHub Secrets 확인)")
        raise SystemExit(1)

    monitor  = _make_monitor(getattr(args, "no_telegram", False))
    state    = _load_pending(date_str)
    cands_state = state.setdefault("candidates", {})
    sold_today  = set(state.setdefault("sold_today", []))
    is_first_run_today = not cands_state and not sold_today

    base_amt  = trader.cfg.get("base_invest_amount", 10_000_000)
    balance   = trader.get_balance()

    # ═══ ① 매도 먼저 (회수금이 그날 매수 예산에 반영돼야 한다) ═══
    sells, recovered, sold_codes = _run_sell_check(trader, balance, monitor, date_str)
    if recovered > 0:
        new_pool = (trader._reinvest_pool if trader._reinvest_pool else base_amt) + recovered
        trader._save_reinvest_pool(new_pool, note=f"장중 매도 회수 {recovered:,.0f}원")
        print(f"  💰 매도 회수 {recovered:,.0f}원 → 재투자 풀 {new_pool:,.0f}원")
        sold_today |= set(sold_codes)
        state["sold_today"] = sorted(sold_today)
        balance = trader.get_balance()      # 매도 반영된 잔고로 갱신

    if sells:
        lines = [f"💸 <b>[한투 장중] 매도 체결</b>  {_now().strftime('%m/%d %H:%M')}"]
        for s in sells:
            mark = "🔴" if s["action"] in ("STOP_LOSS","TRAIL_STOP") else "🟢"
            lines.append(f"  {mark} {s['code']} {s['qty']:,}주 @{s['price']:,}원 ({s['ret_pct']:+.2f}%)")
            lines.append(f"     {s['reason']}")
        lines.append(f"\n  💰 회수 {recovered:,.0f}원 → 재투자 풀 "
                     f"{(trader._reinvest_pool or 0):,.0f}원")
        _tg(monitor, "\n".join(lines))

    # ═══ ② 매수 (컷오프 전에만) ═══
    buys, spent_won, errors = [], 0.0, []
    if _past_buy_cutoff():
        print(f"  ⏰ 매수 컷오프({BUY_CUTOFF_HOUR_MIN[0]:02d}:{BUY_CUTOFF_HOUR_MIN[1]:02d}) 이후 "
              f"— 신규 매수 생략 (매도 점검은 위에서 완료)")
        _save_pending(state, date_str)
        return {"sells": sells, "buys": buys}

    json_path = _resolve_signal_json(args.output_dir)
    if not json_path:
        _save_pending(state, date_str)
        raise SystemExit(1)

    df_top = _load_signal_json_as_df(json_path)
    if df_top.empty:
        print("  ⚠ 신호 파일에 유효 후보 없음")
        _save_pending(state, date_str)
        return {"sells": sells, "buys": buys}

    sig_series = df_top["매매시그널"].astype(str)
    buy_candidates = df_top[sig_series.str.contains("매수", na=False)]
    if buy_candidates.empty:
        print("  ⚠ 매수 시그널 종목 없음")
        _save_pending(state, date_str)
        return {"sells": sells, "buys": buys}

    buy_top_n = trader.cfg.get("buy_top_n", 20)
    pending_codes = [c for c in buy_candidates.index
                     if str(c) not in cands_state
                     or cands_state.get(str(c), {}).get("status") == "pending"]
    # ★ 오늘 이미 판 종목은 다시 사지 않는다.
    #   안 그러면 손절 → 게이트 통과 → 재매수 → 또 손절 로 하루에 같은 종목을
    #   왕복하며 수수료만 태울 수 있다.
    skipped_resell = [c for c in pending_codes if str(c) in sold_today]
    pending_codes  = [c for c in pending_codes if str(c) not in sold_today]
    if skipped_resell:
        print(f"  ↩ 오늘 매도한 종목은 재매수 제외: {skipped_resell}")

    print(f"\n  [매수게이트] 후보 {len(buy_candidates)}종목 중 미해결 {len(pending_codes)}종목 체크")

    bought_today = [c for c, v in cands_state.items() if v.get("status") == "bought"]
    remaining_targets = max(1, buy_top_n - len(bought_today))
    invest_pool = trader._reinvest_pool if trader._reinvest_pool else base_amt
    per_stock_budget = invest_pool / remaining_targets
    print(f"  [예산] 가용 풀 {invest_pool:,.0f}원 / 남은 {remaining_targets}종목 "
          f"→ 종목당 {per_stock_budget:,.0f}원")

    if is_first_run_today:
        _tg(monitor,
            f"🔍 <b>[한투 장중] 감시 시작</b>\n"
            f"  {_now().strftime('%m/%d %H:%M')} · 매수 후보 {len(buy_candidates)}종목\n"
            f"  💰 가용 예산 {invest_pool:,.0f}원 (종목당 {per_stock_budget:,.0f}원)\n"
            f"  📉 매도: 손절 -{abs(trader.cfg.get('stop_loss_pct', DEFAULT_STOP_LOSS_PCT)):.0f}% / "
            f"익절 +{abs(trader.cfg.get('trail_trigger_pct', DEFAULT_TRAIL_TRIGGER_PCT)):.0f}% 도달 후 "
            f"고점 대비 -{abs(trader.cfg.get('trail_drop_pct', DEFAULT_TRAIL_DROP_PCT)):.0f}%\n"
            f"  ⏰ 매수 {BUY_CUTOFF_HOUR_MIN[0]:02d}:{BUY_CUTOFF_HOUR_MIN[1]:02d} / "
            f"매도 {SELL_CUTOFF_HOUR_MIN[0]:02d}:{SELL_CUTOFF_HOUR_MIN[1]:02d}까지, 15분 간격")

    for code in pending_codes:
        code = str(code)
        rec = cands_state.get(code, {"status": "pending", "tries": 0, "first_seen": date_str})
        try:
            if rec.get("tries", 0) >= MAX_TRIES_PER_STOCK:
                rec["status"] = "expired"
                cands_state[code] = rec
                continue

            if _already_held(code, balance):
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
                else:
                    gate_pass = multi["by_interval"].get(args.interval, {}).get("pass", False)
                gate = dict(multi["by_interval"].get(args.interval,
                                                     {"pass": gate_pass, "checks": {}, "detail": {}}))
                gate["pass"] = gate_pass
                gate.setdefault("checks", {})
                gate.setdefault("detail", {})
                gate["multi_detail"] = {
                    "mode": args.gate_mode,
                    "passed_intervals":  multi["passed_intervals"],
                    "failed_intervals":  multi["failed_intervals"],
                    "skipped_intervals": multi.get("skipped", []),
                    "pass_count": f"{multi['pass_count']}/{multi['total_count']}",
                }
                df_min = charts.get(args.interval, pd.DataFrame())
            else:
                df_min = trader.get_minute_chart(code, interval=args.interval)
                time.sleep(0.4)
                gate = evaluate_entry_gate(df_min, direction="BUY",
                                           min_candles=DEFAULT_MIN_CANDLES.get(args.interval, 10))

            rec["tries"] = rec.get("tries", 0) + 1
            rec["last_check"] = _now().strftime("%H:%M:%S")
            rec["last_detail"] = gate.get("detail", {})

            # 매수 여부와 무관하게 "체크한 모든 시점"을 남겨야 나중에
            # "통과했다면 어떻게 됐을지" / "탈락했는데 사실 올랐는지"를 검증할 수 있다.
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
                reason += f"CMF:{gate.get('detail', {}).get('cmf')} | {row.get('매매시그널', '')}"

                r = trader.place_order(code, "BUY", qty, reason=reason)
                if r.get("success"):
                    fill = r.get("price") or cur
                    rec["status"] = "bought"
                    rec["buy_price"] = fill
                    rec["buy_time"] = _now().strftime("%H:%M:%S")
                    spent_won += fill * qty
                    buys.append({"code": code, "name": row.get("종목명", ""), "action": "BUY",
                                 "price": fill, "qty": qty, "gate_detail": gate.get("detail", {})})
                    print(f"  ✅ {code} 게이트 통과 → 매수 {qty}주 @{fill:,}원")
                else:
                    rec["status"] = "pending"
                    print(f"  ⚠ {code} 게이트 통과했지만 주문 실패: {r.get('msg','')}")
            else:
                # "봉이 모자라 판정 불가"와 "지표가 나빠서 탈락"은 전혀 다른 상태다.
                gdetail = gate.get("detail", {})
                if gdetail.get("insufficient"):
                    why = f"판정 불가 — {gdetail.get('error', '데이터 부족')}"
                else:
                    failed = [k for k, v in gate.get("checks", {}).items() if not v]
                    why = f"실패: {', '.join(failed) if failed else '알 수 없음'}"
                extra = ""
                if args.multi_tf:
                    md = gate.get("multi_detail", {})
                    extra = (f" | 멀티({md.get('mode')}) 통과 {md.get('pass_count')} "
                             f"통과={md.get('passed_intervals')} 실패={md.get('failed_intervals')} "
                             f"제외={md.get('skipped_intervals')}")
                print(f"  ⏳ {code} 게이트 미통과 ({why}) — {rec['tries']}회차, 재시도 대기{extra}")

        except Exception as e:
            print(f"  ⚠ {code} 처리 중 예외 — 이 종목만 건너뜁니다")
            traceback.print_exc()
            rec["status"] = "pending"
            rec["last_error"] = _now().strftime("%H:%M:%S")
            errors.append(f"{code}: {type(e).__name__} {e}")

        cands_state[code] = rec

    state["candidates"] = cands_state
    _save_pending(state, date_str)

    if spent_won > 0:
        new_pool = max(0.0, (trader._reinvest_pool if trader._reinvest_pool else base_amt) - spent_won)
        trader._save_reinvest_pool(new_pool, note=f"장중 시그널 매수 {spent_won:,.0f}원 집행")
        print(f"  💸 매수 집행 {spent_won:,.0f}원 → 재투자 풀 잔액 {new_pool:,.0f}원")

    bought_n = sum(1 for v in cands_state.values() if v.get("status") == "bought")
    pending_n = sum(1 for v in cands_state.values() if v.get("status") == "pending")
    expired_n = sum(1 for v in cands_state.values() if v.get("status") == "expired")
    print(f"\n  [시그널체크] 완료 — 매도:{len(sells)}건 매수:{len(buys)}건 "
          f"(누적 보유대상:{bought_n} 대기:{pending_n} 만료:{expired_n})")

    if buys:
        lines = [f"📈 <b>[한투 장중] 매수 체결</b>  {_now().strftime('%m/%d %H:%M')}"]
        for r in buys:
            lines.append(f"  • {r.get('name','')} ({r['code']}) "
                         f"{r['qty']:,}주 @{r['price']:,}원 = {r['price']*r['qty']:,}원")
            d = r.get("gate_detail") or {}
            if d.get("vwap") is not None:
                lines.append(f"     VWAP {d['vwap']:,.0f} · CMF {d.get('cmf')}")
        lines.append(f"\n  💰 남은 예산: {(trader._reinvest_pool if trader._reinvest_pool else base_amt):,.0f}원")
        _tg(monitor, "\n".join(lines))

    if errors:
        _tg(monitor,
            f"⚠️ <b>[한투 장중] 처리 중 오류 {len(errors)}건</b>  {_now().strftime('%m/%d %H:%M')}\n"
            + "\n".join(f"  • {e[:120]}" for e in errors[:5])
            + ("\n  …" if len(errors) > 5 else ""))

    try:
        ds = data_logger.status_summary()
        print(f"  [데이터누적] {ds['days']}거래일치 분봉/게이트 로그 "
              f"(종목×일 파일 {ds.get('total_stock_day_files', 0)}개)")
    except Exception as e:
        print(f"  ⚠ 데이터 누적 상태 조회 실패: {e}")

    return {"sells": sells, "buys": buys}


def _resolve_signal_json(output_dir: str) -> str:
    """
    매매신호 JSON 탐색. quant_daily.yml 이 signals/ 폴더에만 커밋하므로
    output_dir 이 루트여도 signals/ 를 반드시 함께 뒤진다.
    """
    tried = []
    for d in [output_dir, os.path.join(output_dir, "signals"),
              BASE_DIR, os.path.join(BASE_DIR, "signals")]:
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
    print("     → quant_daily.yml(저녁 스크리닝)이 먼저 성공했는지 확인하세요.")
    return ""


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="장중 매도(손절/트레일링 익절) + L1+L2 매수 게이트")
    parser.add_argument("--output-dir", type=str, default="signals",
                        help="매매신호_KR_*.json 폴더 (기본: signals)")
    parser.add_argument("--interval", type=int, choices=[1, 3, 5, 10, 15, 30, 60], default=5,
                        help="게이트 판단 기준 분봉 주기 (기본: 5분)")
    parser.add_argument("--multi-tf", action="store_true",
                        help="1/5/10/15/30/60분봉을 모두 평가")
    parser.add_argument("--gate-mode", choices=["single", "all_pass", "majority"], default="single",
                        help="--multi-tf 켰을 때 최종 매수판단 기준")
    parser.add_argument("--no-telegram", action="store_true", help="텔레그램 알림 끄기")
    args = parser.parse_args()
    run(args)
