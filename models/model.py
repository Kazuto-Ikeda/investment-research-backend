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
# インプットモデル
class ValuationInput(BaseModel):
    revenue_current: float  # 売上（直近期）
    revenue_forecast: float  # 売上（進行期見込み）
    ebitda_current: Optional[float] = None  # EBITDA（直近期）
    ebitda_forecast: Optional[float] = None  # EBITDA（進行期見込み）
    net_debt_current: float  # Net Debt（直近期）
    equity_value_current: float  # 想定Equity Value
    category: str
    

# アウトプットモデル
class ValuationOutput(BaseModel):
    ebitda_current: Optional[float]  # EBITDA（直近期）
    ebitda_forecast: Optional[float]  # EBITDA（進行期見込み）
    net_debt_current: float  # Net Debt（直近期）
    net_debt_forecast: Optional[float]  # Net Debt（進行期見込み）
    equity_value_current: float  # 想定Equity Value（直近期）
    equity_value_forecast: Optional[float]  # 想定Equity Value（進行期見込み）
    ev_current: Optional[float]  # EV（直近期）
    ev_forecast: Optional[float]  # EV（進行期見込み）
    entry_multiple_current: Optional[float]  # エントリーマルチプル（直近期）
    entry_multiple_forecast: Optional[float]  # エントリーマルチプル（進行期見込み）
    industry_median_multiple_current: Optional[float]  # マルチプル業界中央値（直近期）
    industry_median_multiple_forecast: Optional[float]  # マルチプル業界中央値（進行期見込み）
    implied_equity_value_current: Optional[float]  # Implied Equity Value（直近期）
    implied_equity_value_forecast: Optional[float]  # Implied Equity Value（進行期見込み）