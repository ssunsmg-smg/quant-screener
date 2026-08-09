"""
==============================================================
  signal_engine.py — L1(가격) + L2(거래량) 시그널 엔진
  영웅문 고급 매매 전략 문서의 지표를 그대로 구현한 순수 계산 모듈

  ── v2 수정본 (2026-08-09) ─────────────────────────────────
  [수정 1] evaluate_entry_gate() 에 min_candles 인자 추가.
      quant_signal_check.py 가
          evaluate_entry_gate(df, direction="BUY", min_candles=N)
      으로 부르는데 기존 시그니처에 없어서 TypeError 로 죽는다.
      (kis_intraday.get_minute_chart(interval=...) 과 같은 패턴 —
       호출부만 있고 정의부가 안 따라온 세 번째 사례)
      기존 하드코딩 `len(df) < 10` 을 min_candles 로 바꾸되,
      기본값 10 이라 기존 동작은 그대로다.

  [수정 2] DEFAULT_MIN_CANDLES 상수 신규.
      분봉 주기별로 "지표가 의미를 갖는 최소 봉 개수".
      주기가 길수록 하루에 나올 수 있는 봉이 적으므로 완화한다.

  [수정 3] evaluate_entry_gate_multi() 신규 구현.
      --multi-tf 경로가 이 함수를 부르는데 존재하지 않았다.
      ★ 핵심 설계: 봉이 모자라 "평가 자체가 불가능한" 주기는
        실패로 세지 않고 **평가 대상에서 제외**한다.
        예) 60분봉은 정규장 390분 동안 최대 7봉뿐이라 CMF(20기간)가
            영원히 계산되지 않는다. 이걸 '실패'로 세면 all_pass 모드가
            하루 종일 절대 통과하지 못하는데, 그건 "엄격한 게이트"가
            아니라 그냥 고장난 게이트다.

  ※ L3(외인·기관 수급)는 여기 없음 — KIS investor API가
    "당일 데이터는 장 종료 후 제공"이라서, 장중 실시간 게이트로
    쓸 수 없기 때문. L3는 스크리닝 단계(quant_screener)에서
    전일 마감 기준으로 이미 반영된 것으로 간주한다.

  이 파일은 입력으로 OHLCV DataFrame만 받는 순수 함수 모음이라
  yfinance든 KIS 분봉이든 어디서 가져온 데이터든 그대로 사용 가능.
  → 단위 테스트하기 쉽고, KIS API 응답 포맷 변경에도 영향 안 받음.
==============================================================
"""

import numpy as np
import pandas as pd


# ══════════════════════════════════════════════════════════
# ⓪ 분봉 주기별 최소 캔들 요구치
# ══════════════════════════════════════════════════════════
# 지표(특히 CMF는 기본 20기간, 최소 10기간)가 의미를 가지려면 이 정도는
# 있어야 한다. 주기가 길수록 하루에 만들어지는 봉이 적으므로 완화한다.
#
# ⚠ 정규장은 390분(09:00~15:30)이다. 즉 하루 최대 봉 개수는
#     1분:390  5분:78  10분:39  15분:26  30분:13  60분:7
#   이고, 장 초반에는 그보다 훨씬 적다. 예를 들어 5분봉으로 12개를
#   채우려면 09:00부터 60분이 지나야 하므로 **10:00 이전에는 5분봉
#   게이트가 구조적으로 통과할 수 없다.** 이건 버그가 아니라
#   "판단할 데이터가 아직 없다"는 뜻이다.
DEFAULT_MIN_CANDLES = {
    1:  20,
    3:  15,
    5:  12,
    10: 10,
    15: 10,
    30: 10,
    60: 10,   # 하루 최대 7봉이라 사실상 장중 평가 불가 → multi에서 자동 제외됨
}


# ══════════════════════════════════════════════════════════
# ① VWAP + 표준편차 밴드  (영웅문 수식관리자 코드와 동일 로직)
# ══════════════════════════════════════════════════════════
def calc_vwap_bands(df: pd.DataFrame) -> pd.DataFrame:
    """
    df 컬럼 요구: High, Low, Close, Volume
    리턴: TP, VWAP, DEV, VWAP_UP1, VWAP_DN1, VWAP_UP2, VWAP_DN2 컬럼 추가된 df

    ※ 일중(intraday) 누적 기준. 분봉 데이터를 넣으면 그날 하루 누적으로
      VWAP이 계산된다 (영웅문 N=1 옵션과 동일).

    ⚠ 중요: 이 함수는 "넘겨받은 df의 첫 행"부터 누적한다. 따라서 df가
      장 시작(09:00)부터 담겨 있어야 진짜 일중 VWAP이 된다. 최근 60분만
      잘라서 넣으면 그건 '60분 VWAP'이지 일중 VWAP이 아니다.
      → kis_intraday.get_minute_chart()가 장 시작부터 받아오도록 되어 있다.
    """
    out = df.copy()
    out["TP"] = (out["High"] + out["Low"] + out["Close"]) / 3.0

    cum_vol = out["Volume"].cumsum()
    cum_tpv = (out["TP"] * out["Volume"]).cumsum()
    out["VWAP"] = cum_tpv / cum_vol.replace(0, np.nan)

    # 표준편차 (거래량 가중)
    cum_dev2 = ((out["TP"] - out["VWAP"]) ** 2 * out["Volume"]).cumsum()
    out["DEV"] = np.sqrt(cum_dev2 / cum_vol.replace(0, np.nan))

    out["VWAP_UP1"] = out["VWAP"] + out["DEV"]
    out["VWAP_DN1"] = out["VWAP"] - out["DEV"]
    out["VWAP_UP2"] = out["VWAP"] + 2 * out["DEV"]
    out["VWAP_DN2"] = out["VWAP"] - 2 * out["DEV"]
    return out


# ══════════════════════════════════════════════════════════
# ② 매물대 / Volume Profile / POC
# ══════════════════════════════════════════════════════════
def calc_volume_profile(df: pd.DataFrame, bins: int = 24, value_area_pct: float = 0.70) -> dict:
    """
    가격 구간별 거래량을 집계해 POC(최대 거래량 가격) / VAH·VAL(70% 매물 영역) 산출.

    df 컬럼 요구: High, Low, Close, Volume
    리턴: {"poc": float, "vah": float, "val": float, "profile": pd.Series}
    """
    if df.empty or df["Volume"].sum() == 0:
        return {"poc": np.nan, "vah": np.nan, "val": np.nan, "profile": pd.Series(dtype=float)}

    lo, hi = df["Low"].min(), df["High"].max()
    if hi <= lo:
        mid = float(df["Close"].iloc[-1])
        return {"poc": mid, "vah": mid, "val": mid, "profile": pd.Series(dtype=float)}

    edges = np.linspace(lo, hi, bins + 1)
    # 각 캔들의 거래량을 (고가+저가+종가)/3 가격대 bin에 배분 (간이 방식 — 분봉 단위라 충분히 정확)
    tp = (df["High"] + df["Low"] + df["Close"]) / 3.0
    bin_idx = np.clip(np.digitize(tp, edges) - 1, 0, bins - 1)
    profile = pd.Series(0.0, index=range(bins))
    for idx, vol in zip(bin_idx, df["Volume"]):
        profile.iloc[int(idx)] += float(vol)

    centers = (edges[:-1] + edges[1:]) / 2
    poc_bin = int(np.argmax(profile.to_numpy()))
    poc = float(centers[poc_bin])

    # Value Area: POC에서 시작해 위/아래로 거래량 비중 70% 채울 때까지 확장
    vals = profile.to_numpy()
    total_vol = float(vals.sum())
    target = total_vol * value_area_pct
    included = {poc_bin}
    acc = float(vals[poc_bin])
    lo_i, hi_i = poc_bin, poc_bin
    while acc < target and (lo_i > 0 or hi_i < bins - 1):
        next_lo = vals[lo_i - 1] if lo_i > 0 else -1.0
        next_hi = vals[hi_i + 1] if hi_i < bins - 1 else -1.0
        if next_hi >= next_lo:
            hi_i += 1
            acc += float(vals[hi_i])
            included.add(hi_i)
        else:
            lo_i -= 1
            acc += float(vals[lo_i])
            included.add(lo_i)

    val = float(centers[min(included)])
    vah = float(centers[max(included)])

    profile.index = centers
    return {"poc": poc, "vah": vah, "val": val, "profile": profile}


# ══════════════════════════════════════════════════════════
# ③ OBV (On-Balance Volume)
# ══════════════════════════════════════════════════════════
def calc_obv(df: pd.DataFrame) -> pd.Series:
    """df 컬럼 요구: Close, Volume"""
    direction = np.sign(df["Close"].diff()).fillna(0)
    obv = (direction * df["Volume"]).cumsum()
    obv.name = "OBV"
    return obv


# ══════════════════════════════════════════════════════════
# ④ CMF (Chaikin Money Flow)
# ══════════════════════════════════════════════════════════
def calc_cmf(df: pd.DataFrame, period: int = 20) -> pd.Series:
    """df 컬럼 요구: High, Low, Close, Volume"""
    hl_range = (df["High"] - df["Low"]).replace(0, np.nan)
    mfm = ((df["Close"] - df["Low"]) - (df["High"] - df["Close"])) / hl_range
    mfv = mfm * df["Volume"]
    cmf = mfv.rolling(period, min_periods=max(1, period // 2)).sum() / \
          df["Volume"].rolling(period, min_periods=max(1, period // 2)).sum()
    cmf.name = "CMF"
    return cmf.fillna(0)


# ══════════════════════════════════════════════════════════
# ⑤ L1+L2 진입 게이트 (이중확인 — L3는 스크리닝 단계에서 이미 통과)
# ══════════════════════════════════════════════════════════
def evaluate_entry_gate(df: pd.DataFrame, direction: str = "BUY",
                        cmf_threshold: float = 0.05,
                        vp_bins: int = 24,
                        min_candles: int = 10) -> dict:
    """
    분봉(또는 일봉) OHLCV df를 받아 L1(가격)+L2(거래량) 동시 확인 결과를 반환.

    df 컬럼 요구: Open, High, Low, Close, Volume  (시간 오름차순 정렬)
    direction   : "BUY" | "SELL"
    min_candles : 판정에 필요한 최소 캔들 수. 미달이면 pass=False +
                  detail["insufficient"]=True 로 표시한다.
                  ★ "미달"과 "지표가 나빠서 탈락"은 전혀 다른 상태이므로
                    호출부가 구분할 수 있도록 별도 플래그를 준다.
                  (기본 10 — 기존 하드코딩 값과 동일하므로 동작 변화 없음)

    리턴 예:
    {
      "pass": True,
      "direction": "BUY",
      "checks": {
        "vwap_position": True,   # 가격이 VWAP 위(매수) / 아래(매도)
        "obv_trend": True,       # OBV가 같은 방향으로 움직임
        "cmf_strength": True,    # CMF가 방향과 일치 + 임계값 이상
      },
      "detail": {...}            # 디버깅/로그용 실제 수치
    }
    """
    n = 0 if df is None else len(df)
    if n < min_candles:
        return {"pass": False, "direction": direction, "checks": {},
                "detail": {"error": f"데이터 부족 ({n}/{min_candles} 캔들)",
                           "insufficient": True,
                           "candles": n, "min_candles": min_candles}}

    vw  = calc_vwap_bands(df)
    obv = calc_obv(df)
    cmf = calc_cmf(df)
    vp  = calc_volume_profile(df, bins=vp_bins)

    last_close = float(df["Close"].iloc[-1])
    last_vwap  = float(vw["VWAP"].iloc[-1])
    last_obv   = float(obv.iloc[-1])
    prev_obv   = float(obv.iloc[-min(5, len(obv))])   # 최근 5캔들 전 대비
    last_cmf   = float(cmf.iloc[-1])

    if direction == "BUY":
        vwap_ok = last_close >= last_vwap            # VWAP 위 (또는 ±σ 밴드 하단 터치 후 반등 — 운용 시 조건 조정 가능)
        obv_ok  = last_obv >= prev_obv               # OBV 상승 중
        cmf_ok  = last_cmf >= cmf_threshold          # 매수 압력 강도 충족
    else:  # SELL
        vwap_ok = last_close <= last_vwap
        obv_ok  = last_obv <= prev_obv
        cmf_ok  = last_cmf <= -cmf_threshold

    checks = {"vwap_position": bool(vwap_ok), "obv_trend": bool(obv_ok), "cmf_strength": bool(cmf_ok)}
    passed = all(checks.values())

    return {
        "pass": passed,
        "direction": direction,
        "checks": checks,
        "detail": {
            "close": last_close, "vwap": last_vwap,
            "vwap_up1": float(vw["VWAP_UP1"].iloc[-1]), "vwap_dn1": float(vw["VWAP_DN1"].iloc[-1]),
            "obv": last_obv, "obv_prev5": prev_obv,
            "cmf": round(last_cmf, 4),
            "poc": vp["poc"], "vah": vp["vah"], "val": vp["val"],
            "insufficient": False, "candles": n,
        }
    }


# ══════════════════════════════════════════════════════════
# ⑥ 멀티 타임프레임 게이트 (--multi-tf)
# ══════════════════════════════════════════════════════════
def evaluate_entry_gate_multi(charts: dict, direction: str = "BUY",
                              cmf_threshold: float = 0.05,
                              vp_bins: int = 24,
                              min_candles_map: dict = None) -> dict:
    """
    여러 분봉 주기의 게이트를 한 번에 평가.

    charts : {interval(int): OHLCV DataFrame}  — kis_intraday.get_minute_chart_multi() 결과
    리턴:
    {
      "by_interval":       {interval: gate_dict},   # 평가 불가한 주기도 포함(로그용)
      "evaluable":         [1, 5, 10],              # 실제로 판정된 주기
      "skipped":           [30, 60],                # 봉 부족으로 평가 제외된 주기
      "passed_intervals":  [1, 5],
      "failed_intervals":  [10],
      "pass_count":        2,
      "total_count":       3,                       # = len(evaluable)
      "all_pass":          False,
      "majority_pass":     True,
    }

    ★ 설계 의도: 봉이 모자라 계산 자체가 불가능한 주기는 '실패'가 아니라
      '평가 제외'다. 60분봉은 정규장 390분 동안 최대 7봉뿐이라 CMF가
      영영 계산되지 않는데, 이걸 실패로 세면 all_pass 모드는 하루 종일
      절대 통과할 수 없다. 그건 엄격한 게이트가 아니라 고장난 게이트다.

    ⚠ evaluable 이 0개면 all_pass / majority_pass 는 모두 False 다.
      "판단 근거가 없을 때는 사지 않는다"가 안전한 기본값이다.
    """
    min_map = min_candles_map or DEFAULT_MIN_CANDLES

    by_interval, evaluable, skipped = {}, [], []
    for iv in sorted(charts.keys(), key=int):
        iv = int(iv)
        gate = evaluate_entry_gate(charts[iv], direction=direction,
                                   cmf_threshold=cmf_threshold, vp_bins=vp_bins,
                                   min_candles=min_map.get(iv, 10))
        by_interval[iv] = gate
        (skipped if gate["detail"].get("insufficient") else evaluable).append(iv)

    passed = [iv for iv in evaluable if by_interval[iv]["pass"]]
    failed = [iv for iv in evaluable if not by_interval[iv]["pass"]]
    total  = len(evaluable)

    return {
        "by_interval":      by_interval,
        "evaluable":        evaluable,
        "skipped":          skipped,
        "passed_intervals": passed,
        "failed_intervals": failed,
        "pass_count":       len(passed),
        "total_count":      total,
        "all_pass":         total > 0 and len(passed) == total,
        "majority_pass":    total > 0 and len(passed) * 2 > total,   # 과반수(동수는 불통과)
    }


if __name__ == "__main__":
    # 간단 자체 테스트 (합성 데이터)
    rng = pd.date_range("2026-06-27 09:00", periods=60, freq="1min")
    np.random.seed(0)
    price = 70000 + np.cumsum(np.random.randn(60) * 50)
    test_df = pd.DataFrame({
        "Open": price, "High": price + np.random.rand(60) * 30,
        "Low": price - np.random.rand(60) * 30, "Close": price + np.random.randn(60) * 10,
        "Volume": np.random.randint(1000, 5000, 60),
    }, index=rng)

    result = evaluate_entry_gate(test_df, direction="BUY")
    print("=== 자체 테스트 (합성 데이터) ===")
    print(f"통과 여부: {result['pass']}")
    print(f"세부 체크: {result['checks']}")
    print(f"상세: {result['detail']}")

    # 멀티 타임프레임 테스트
    agg = {"Open": "first", "High": "max", "Low": "min", "Close": "last", "Volume": "sum"}
    charts = {iv: (test_df if iv == 1 else
                   test_df.resample(f"{iv}min").agg(agg).dropna(subset=["Close"]))
              for iv in (1, 5, 10, 15, 30, 60)}
    multi = evaluate_entry_gate_multi(charts, direction="BUY")
    print("\n=== 멀티 타임프레임 ===")
    print(f"평가 대상: {multi['evaluable']}  /  봉 부족 제외: {multi['skipped']}")
    print(f"통과: {multi['passed_intervals']}  실패: {multi['failed_intervals']}")
    print(f"all_pass={multi['all_pass']}  majority_pass={multi['majority_pass']} "
          f"({multi['pass_count']}/{multi['total_count']})")
