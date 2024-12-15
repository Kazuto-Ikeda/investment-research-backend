from pydantic import BaseModel
from typing import Optional

#再生成
class RegenerateRequest(BaseModel):
    blob_name: str
    company_name: str
    query_key: str
    custom_query: Optional[str] = None  # ユーザーがカスタムクエリを指定する場合
    include_perplexity: bool = False


# バリュエーション計算
# インプットモデル
class ValuationInput(BaseModel):
    revenue_current: float  # 売上（直近期）
    revenue_forecast: float  # 売上（進行期見込）
    ebitda_current: Optional[float] = None  # EBITDA（直近期）
    ebitda_forecast: Optional[float] = None  # EBITDA（進行期見込）
    net_debt_current: float  # Net Debt（直近期）
    equity_value_current: float  # 想定Equity Value
    category: str  # レポートカテゴリ

# アウトプットモデル
class ValuationOutput(BaseModel):
    # 売上
    revenue_current: Optional[int]  # 売上（直近期）
    revenue_forecast: Optional[int]  # 売上（進行期見込）
    
    # EBITDA
    ebitda_current: Optional[int]  # EBITDA（直近実績）
    ebitda_forecast: Optional[int]  # EBITDA（進行期見込）
    
    # Net Debt
    net_debt_current: Optional[int]  # Net Debt（直近実績）
    net_debt_forecast: Optional[int]  # Net Debt（進行期見込）
    
    # Equity Value
    equity_value_current: Optional[int]  # 想定Equity Value（直近実績）
    equity_value_forecast: Optional[int]  # 想定Equity Value（進行期見込み）
    
    # Enterprise Value (EV)
    ev_current: Optional[int]  # EV（直近実績）
    ev_forecast: Optional[int]  # EV（進行期見込み）
    
    # Entry Multiple
    entry_multiple_current: Optional[str] = None  # エントリーマルチプル（直近実績）
    entry_multiple_forecast: Optional[str] = None  # エントリーマルチプル（進行期見込み）
    
    # Industry Median Multiple
    industry_median_multiple_current: Optional[str] = None  # マルチプル業界中央値（直近実績）
    industry_median_multiple_forecast: Optional[str] = None  # マルチプル業界中央値（進行期見込み）