"""리샘플된 파생 CSV(일봉/5분봉) 로더."""
from __future__ import annotations
from pathlib import Path
import pandas as pd

DATA = Path(__file__).parent / "data"


def _read_csv_any(path):
    for enc in ("utf-8-sig", "cp949", "euc-kr", "utf-8"):
        try:
            return pd.read_csv(path, encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    with open(path, "rb") as fh:
        if fh.read(2) == bytes([0x9b, 0x20]):
            raise RuntimeError(f"[DRM] {Path(path).name} 암호화됨 → 평문화 필요")
    raise RuntimeError(f"읽기 실패: {path}")


def load_bars(product: str, tf: str) -> pd.DataFrame:
    """product: 'KTB3'|'KTB10', tf: 'daily'|'5min'."""
    path = DATA / f"{product}_{tf}.csv"
    df = _read_csv_any(path)
    df["datetime"] = pd.to_datetime(df["datetime"])
    df = df.dropna(subset=["datetime", "close"]).sort_values("datetime")
    df = df.set_index("datetime")
    df = df[~df.index.duplicated(keep="last")]
    return df
