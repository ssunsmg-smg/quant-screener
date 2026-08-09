"""
==============================================================
  kis_intraday.py — KIS API 분봉시세 / 투자자(외인·기관) 데이터
  기존 quant_screener_v41f.py 의 KISAutoTrader 를 그대로 상속해서
  인증(get_access_token)·주문(place_order) 코드는 손대지 않고 재사용.

  ── v2 수정본 (2026-08-09) ─────────────────────────────────
  [수정 1] get_minute_chart() 에 `interval` 인자 추가.
      quant_signal_check.py 가
          trader.get_minute_chart(code, interval=..., lookback_calls=2)
      로 호출하는데 기존 시그니처엔 interval 이 없어서 TypeError 로
      죽고 있었다. (매수 후보 첫 종목에서 즉시 크래시)

  [수정 2] KIS `inquire-time-itemchartprice`(FHKST03010200)는
      **1분봉만** 제공한다. 5/10/15/30/60분봉이라는 파라미터가
      아예 없다. 따라서 1분봉을 받아 pandas resample 로 직접
      합성한다. (기존 코드는 interval 을 아무데도 안 쓰고 있었음)

  [수정 3] get_minute_chart_multi() 신규 구현.
      quant_signal_check.py 의 --multi-tf 경로가 이 메서드를 부르는데
      존재하지 않아 AttributeError 였다.
      주기별로 따로 API 를 때리면 호출수가 6배가 되어 rate limit 에
      걸리므로, **1분봉을 한 번만 받아서** 여러 주기로 리샘플한다.

  [수정 4] 페이지네이션 경계 버그.
      기존: 다음 조회 시각 = 이번 응답의 마지막 행 시각 (그대로)
            → 경계 분봉을 매번 중복 조회
      수정: 가장 이른 시각에서 1분을 빼서 조회 + 중복 시각 집합 추적,
            새 데이터가 안 나오면 즉시 중단 (무한 루프 방지)

  [수정 5] 인덱스를 "HHMMSS" 문자열 → DatetimeIndex 로 변경.
      resample 에 필요하고, 시간 연산도 정확해진다.

  [수정 6] 시각 기준을 KST 로 고정.
      GitHub Actions 러너는 UTC 라서 datetime.now() 가 한국시간보다
      9시간 이르다. 그대로 FID_INPUT_HOUR_1 에 넣으면 장 시작 전
      시각을 조회하게 된다.

  ⚠ 사용 전 확인:
    - quant_screener_v41f.py 와 같은 폴더에 둘 것 (import 때문에)
    - kis_config.json 의 app_key/app_secret 은 실거래용이므로,
      이 모듈은 "조회"만 하지만 같은 토큰을 쓰는 만큼 신중하게 다룰 것
==============================================================
"""

import math
import time
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests

from quant_screener_v41f import KISAutoTrader   # 기존 인증·주문 클래스 재사용

# ── KST 고정 (Actions 러너는 UTC) ──
try:
    from zoneinfo import ZoneInfo
    KST = ZoneInfo("Asia/Seoul")
except Exception:                                  # tzdata 없는 환경 폴백
    KST = timezone(timedelta(hours=9))

# KIS 분봉 API 1회 호출당 반환 건수(분). 공식 스펙상 최대 30건.
MINUTES_PER_CALL = 30
# 정규장 09:00~15:30 = 390분 → 전 구간을 받으려면 13회 호출이면 충분
MAX_CALLS_FULL_SESSION = 13
MARKET_OPEN_HHMMSS = "090000"


def _now_kst() -> datetime:
    return datetime.now(KST)


class KISIntraday(KISAutoTrader):
    """
    KISAutoTrader 상속 → get_access_token(), _headers(), base_url 등
    인증 로직을 그대로 쓰면서, 분봉/투자자 조회 메서드만 추가.
    """

    # 리샘플 집계 규칙 (OHLCV 표준)
    _RESAMPLE_AGG = {
        "Open":   "first",
        "High":   "max",
        "Low":    "min",
        "Close":  "last",
        "Volume": "sum",
    }

    # ──────────────────────────────────────────────────────
    #  내부: 1분봉 원본 수집 (주식당일분봉조회 FHKST03010200)
    # ──────────────────────────────────────────────────────
    def _fetch_1min(self, stock_code: str, max_calls: int = 2) -> pd.DataFrame:
        """
        최근 1분봉 OHLCV 를 시간 오름차순 DataFrame 으로 반환.
        컬럼: Open, High, Low, Close, Volume  (index: tz-naive KST DatetimeIndex)

        max_calls 만큼 과거로 페이지네이션한다 (1회 ≈ 30분).
        """
        all_rows = []
        seen_times = set()

        now_kst   = _now_kst()
        day_str   = now_kst.strftime("%Y%m%d")
        inqr_hour = now_kst.strftime("%H%M%S")

        for _ in range(max(1, int(max_calls))):
            try:
                resp = requests.get(
                    f"{self.base_url}/uapi/domestic-stock/v1/quotations/inquire-time-itemchartprice",
                    headers=self._headers("FHKST03010200"),
                    params={
                        "FID_ETC_CLS_CODE":     "",
                        "FID_COND_MRKT_DIV_CODE": "J",
                        "FID_INPUT_ISCD":       stock_code,
                        "FID_INPUT_HOUR_1":     inqr_hour,
                        "FID_PW_DATA_INCU_YN":  "N",
                    },
                    timeout=10,
                )
                data = resp.json()
                if data.get("rt_cd") != "0":
                    print(f"  ⚠ [분봉조회] {stock_code} 실패: {data.get('msg1', '')}")
                    break

                rows = [r for r in data.get("output2", []) if r.get("stck_cntg_hour")]
                if not rows:
                    break

                # 이미 받은 시각은 버린다 (경계 중복 + 같은 구간 반복 방지)
                fresh = [r for r in rows if r["stck_cntg_hour"] not in seen_times]
                if not fresh:
                    # 더 과거 데이터가 없어 같은 구간이 반복되는 상황 → 중단
                    break
                for r in fresh:
                    seen_times.add(r["stck_cntg_hour"])
                all_rows.extend(fresh)

                # 다음 호출은 이번 응답의 "가장 이른 시각 − 1분" 부터
                oldest = min(r["stck_cntg_hour"] for r in rows)
                try:
                    nxt = datetime.strptime(oldest, "%H%M%S") - timedelta(minutes=1)
                except ValueError:
                    break
                inqr_hour = nxt.strftime("%H%M%S")
                if inqr_hour < MARKET_OPEN_HHMMSS:
                    break                      # 장 시작 이전까지 다 받음

                time.sleep(0.3)                # KIS rate limit 여유
            except Exception as e:
                print(f"  ⚠ [분봉조회] {stock_code} 오류: {e}")
                break

        if not all_rows:
            return pd.DataFrame(columns=["Open", "High", "Low", "Close", "Volume"])

        df = pd.DataFrame(all_rows)
        # KIS 응답 필드명:
        #   stck_cntg_hour(체결시각), stck_oprc/hgpr/lwpr/prpr(시/고/저/현재가), cntg_vol(체결량)
        df = df.rename(columns={
            "stck_cntg_hour": "time", "stck_oprc": "Open", "stck_hgpr": "High",
            "stck_lwpr": "Low", "stck_prpr": "Close", "cntg_vol": "Volume",
        })
        for col in ["Open", "High", "Low", "Close", "Volume"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
            else:
                df[col] = pd.NA

        df = df.dropna(subset=["Close"]).drop_duplicates(subset=["time"])
        if df.empty:
            return pd.DataFrame(columns=["Open", "High", "Low", "Close", "Volume"])

        # "HHMMSS" → 오늘 날짜(KST) 기준 DatetimeIndex
        df["dt"] = pd.to_datetime(
            day_str + df["time"].astype(str).str.zfill(6),
            format="%Y%m%d%H%M%S", errors="coerce",
        )
        df = df.dropna(subset=["dt"]).sort_values("dt").set_index("dt")
        return df[["Open", "High", "Low", "Close", "Volume"]]

    # ──────────────────────────────────────────────────────
    #  내부: 1분봉 → N분봉 리샘플
    # ──────────────────────────────────────────────────────
    @classmethod
    def _resample(cls, df1: pd.DataFrame, interval: int) -> pd.DataFrame:
        """1분봉 DataFrame 을 interval 분봉으로 합성. 09:00 을 기준점으로 정렬."""
        if df1 is None or df1.empty or int(interval) <= 1:
            return df1 if df1 is not None else pd.DataFrame()

        # 09:00 정각을 origin 으로 삼아야 09:00~09:05 / 09:05~09:10 처럼
        # 증권사 HTS 와 같은 경계로 묶인다.
        origin = df1.index[0].normalize() + pd.Timedelta(hours=9)

        out = (df1.resample(f"{int(interval)}min",
                            origin=origin, label="left", closed="left")
                  .agg(cls._RESAMPLE_AGG)
                  .dropna(subset=["Close"]))
        return out

    @staticmethod
    def _calls_since_open(now: datetime = None) -> int:
        """장 시작(09:00 KST)부터 지금까지를 덮으려면 몇 번 호출해야 하는지."""
        now = now or _now_kst()
        open_t = now.replace(hour=9, minute=0, second=0, microsecond=0)
        elapsed = max(0, int((now - open_t).total_seconds() // 60))
        return max(1, min(MAX_CALLS_FULL_SESSION, math.ceil(elapsed / MINUTES_PER_CALL)))

    @classmethod
    def _calls_needed(cls, interval: int, min_bars: int = 12) -> int:
        """
        1분봉을 몇 번 호출해야 하는지.

        ★ 장 시작(09:00)부터 전 구간을 받는다.
          이유: signal_engine.calc_vwap_bands() 는 "넘겨받은 df의 첫 행부터"
          누적해서 VWAP 을 만든다. 최근 60분만 잘라서 넣으면 그건 60분 VWAP
          이지 일중 VWAP 이 아니고, 게이트의 핵심 조건(종가 >= VWAP)이
          영웅문 화면과 전혀 다른 값으로 판정된다.
          → 지표를 맞추려면 반드시 장 시작부터 받아야 한다.

        주기가 길어서 더 많이 필요하면 그쪽을 따른다.
        """
        need_minutes = max(1, int(interval)) * max(1, int(min_bars))
        by_bars = math.ceil(need_minutes / MINUTES_PER_CALL)
        return max(1, min(MAX_CALLS_FULL_SESSION, max(by_bars, cls._calls_since_open())))

    # ──────────────────────────────────────────────────────
    #  공개: 단일 주기 분봉
    # ──────────────────────────────────────────────────────
    def get_minute_chart(self, stock_code: str, interval: int = 1,
                         lookback_calls: int = None, min_bars: int = 12) -> pd.DataFrame:
        """
        interval 분봉 OHLCV 를 시간 오름차순 DataFrame 으로 반환.
        컬럼: Open, High, Low, Close, Volume  (index: DatetimeIndex)

        interval        : 1/3/5/10/15/30/60 (1분봉을 받아 합성)
        lookback_calls  : 기존 호출부 호환용. 지정해도 interval 에 필요한
                          최소 호출수보다 작으면 자동으로 올려서 조회한다.
                          (예: 30분봉을 2회 호출로 받으면 봉이 2개뿐이라
                           게이트 지표가 계산되지 않음)
        """
        need = self._calls_needed(interval, min_bars=min_bars)
        if lookback_calls:
            need = max(need, min(MAX_CALLS_FULL_SESSION, int(lookback_calls)))

        df1 = self._fetch_1min(stock_code, max_calls=need)
        if df1.empty:
            return df1
        return self._resample(df1, interval)

    # ──────────────────────────────────────────────────────
    #  공개: 멀티 주기 분봉 (--multi-tf 용)
    # ──────────────────────────────────────────────────────
    def get_minute_chart_multi(self, stock_code: str,
                               intervals=(1, 5, 10, 15, 30, 60),
                               min_bars: int = 12) -> dict:
        """
        여러 분봉 주기를 한 번에 반환: {interval: DataFrame, ...}

        ★ 핵심: 1분봉을 **단 한 번만** 조회해서 전부 리샘플로 만든다.
          주기마다 API 를 따로 때리면 종목당 호출수가 6배가 되어
          KIS rate limit(초당 건수 제한)에 바로 걸린다.

        ⚠ 60분봉 주의: 정규장이 390분뿐이라 하루에 최대 6~7개 봉밖에
          안 나온다. 장 초반에는 봉이 1~2개라 지표 계산이 불가능하므로
          evaluate_entry_gate_multi 쪽에서 "평가 불가"로 걸러진다.
          이건 API 한계가 아니라 시장 구조상 당연한 결과다.
        """
        intervals = tuple(int(i) for i in intervals if int(i) >= 1)
        if not intervals:
            return {}

        need = self._calls_needed(max(intervals), min_bars=min_bars)
        df1  = self._fetch_1min(stock_code, max_calls=need)

        charts = {}
        for iv in intervals:
            charts[iv] = self._resample(df1, iv) if not df1.empty else pd.DataFrame()
        return charts

    # ── 주식현재가 투자자 (FHKST01010900) ──
    # ⚠ 핵심 제약: 당일 데이터는 장 종료 후에만 채워진다 (KIS 공식 안내).
    #   따라서 장중에는 항상 "전일 기준" 값만 얻을 수 있다.
    #   → 스크리닝(전날 저녁) 단계에서 호출해 L3 필터로 쓰는 용도이지,
    #     장중 시그널 게이트에는 쓰지 않는다.
    def get_investor_flow(self, stock_code: str) -> dict:
        """
        최근 영업일 기준 외국인/기관 순매수(수량) 반환.
        리턴: {"date": "YYYYMMDD", "foreign_net": int, "institution_net": int, "individual_net": int}
        실패 시 모든 값 0.
        """
        try:
            resp = requests.get(
                f"{self.base_url}/uapi/domestic-stock/v1/quotations/inquire-investor",
                headers=self._headers("FHKST01010900"),
                params={"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": stock_code},
                timeout=10,
            )
            data = resp.json()
            if data.get("rt_cd") != "0":
                print(f"  ⚠ [투자자조회] {stock_code} 실패: {data.get('msg1', '')}")
                return {"date": None, "foreign_net": 0, "institution_net": 0, "individual_net": 0}

            rows = data.get("output", [])
            if not rows:
                return {"date": None, "foreign_net": 0, "institution_net": 0, "individual_net": 0}

            latest = rows[0]   # 최신 영업일이 첫 행
            return {
                "date": latest.get("stck_bsop_date"),
                # 필드명은 KIS 응답 기준 — frgn_ntby_qty(외국인순매수수량), orgn_ntby_qty(기관순매수수량)
                "foreign_net":     int(latest.get("frgn_ntby_qty", 0) or 0),
                "institution_net": int(latest.get("orgn_ntby_qty", 0) or 0),
                "individual_net":  int(latest.get("prsn_ntby_qty", 0) or 0),
            }
        except Exception as e:
            print(f"  ⚠ [투자자조회] {stock_code} 오류: {e}")
            return {"date": None, "foreign_net": 0, "institution_net": 0, "individual_net": 0}

    def get_investor_flow_accumulated(self, stock_code: str, days: int = 5) -> dict:
        """
        최근 N영업일 외국인/기관 순매수 합산 (5일 누적 — 체크리스트의 '5일 누적순매수'에 대응).
        inquire-investor는 최근 영업일 여러 건을 한 번에 반환하므로 별도 API 호출 불필요.
        """
        try:
            resp = requests.get(
                f"{self.base_url}/uapi/domestic-stock/v1/quotations/inquire-investor",
                headers=self._headers("FHKST01010900"),
                params={"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": stock_code},
                timeout=10,
            )
            data = resp.json()
            if data.get("rt_cd") != "0":
                print(f"  ⚠ [투자자누적조회] {stock_code} 실패: {data.get('msg1', '')}")
                return {"days": 0, "foreign_net_sum": 0, "institution_net_sum": 0}
            rows = data.get("output", [])[:days]
            foreign_sum = sum(int(r.get("frgn_ntby_qty", 0) or 0) for r in rows)
            inst_sum    = sum(int(r.get("orgn_ntby_qty", 0) or 0) for r in rows)
            return {"days": len(rows), "foreign_net_sum": foreign_sum, "institution_net_sum": inst_sum}
        except Exception as e:
            print(f"  ⚠ [투자자누적조회] {stock_code} 오류: {e}")
            return {"days": 0, "foreign_net_sum": 0, "institution_net_sum": 0}


# ── 단독 실행 시 자가진단 (모의투자 계좌 권장) ──
if __name__ == "__main__":
    import sys
    code = sys.argv[1] if len(sys.argv) > 1 else "005930"
    t = KISIntraday()
    if not t._is_configured():
        print("⚠ kis_config.json 앱키 미설정")
        raise SystemExit(1)

    print(f"\n▶ {code} 1분봉")
    print(t.get_minute_chart(code, interval=1).tail())

    print(f"\n▶ {code} 5분봉 (1분봉 리샘플)")
    print(t.get_minute_chart(code, interval=5).tail())

    print(f"\n▶ {code} 멀티 분봉 봉 개수")
    for iv, df in t.get_minute_chart_multi(code).items():
        print(f"   {iv:>3}분봉: {len(df)}개")

    print(f"\n▶ {code} 투자자 수급")
    print(t.get_investor_flow(code))
