from io import StringIO

import pandas as pd
import httpx


async def read_google_sheet_csv(url: str, header: int = 0) -> pd.DataFrame:
    async with httpx.AsyncClient(timeout=20) as client:
        response = await client.get(url)
        response.raise_for_status()
    return pd.read_csv(StringIO(response.text), header=header)
