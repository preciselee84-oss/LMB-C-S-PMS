import pandas as pd


def summarize_points(rows: list[dict]) -> dict[str, int]:
    df = pd.DataFrame(rows)
    if df.empty or "points" not in df.columns:
        return {"total_points": 0}

    return {"total_points": int(df["points"].fillna(0).sum())}

