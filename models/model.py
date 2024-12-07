from pydantic import BaseModel
from typing import Optional

#再生成
class RegenerateRequest(BaseModel):
    blob_name: str
    company_name: str
    query_key: str
    custom_query: Optional[str] = None  # ユーザーがカスタムクエリを指定する場合
    include_perplexity: bool = False

#バリュエーション計算
class ValuationInput(BaseModel):
    revenue_current: float  # 売上（直近期）
    revenue_forecast: float  # 売上（当期見込み）
    ebitda_current: float  # EBITDA（直近期）
    ebitda_forecast: float  # EBITDA（当期見込み）
    net_debt_current: float  # Net Debt（直近期）
    net_debt_forecast: Optional[float] = 0.0  # Net Debt（当期見込み）, デフォルト 0.0
    equity_value_current: float  # 想定Equity Value（直近期）
    equity_value_forecast: Optional[float] = 0.0  # 想定Equity Value（当期見込み）, デフォルト 0.0
    industry_median_multiple: float  # マルチプル業界中央値（直近期）
    industry_median_multiple_forecast: Optional[float] = 0.0  # マルチプル業界中央値（当期見込み）
