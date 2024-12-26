from pydantic import BaseModel
from typing import Optional, Dict
import mistune
from mistune import Markdown


#再生成
class RegenerateRequest(BaseModel):
    blob_name: str
    company_name: str
    query_key: str
    custom_query: Optional[str] = None  # ユーザーがカスタムクエリを指定する場合
    include_perplexity: bool = False


## バリュエーション計算
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
    revenue_current: Optional[str]  # 売上（直近期）
    revenue_forecast: Optional[str]  # 売上（進行期見込）
    
    # EBITDA
    ebitda_current: Optional[str]  # EBITDA（直近実績）
    ebitda_forecast: Optional[str]  # EBITDA（進行期見込）
    
    # Net Debt
    net_debt_current: Optional[str]  # Net Debt（直近実績）
    net_debt_forecast: Optional[str]  # Net Debt（進行期見込）
    
    # Equity Value
    equity_value_current: Optional[str]  # 想定Equity Value（直近実績）
    equity_value_forecast: Optional[str]  # 想定Equity Value（進行期見込み）
    
    # Enterprise Value (EV)
    ev_current: Optional[str]  # EV（直近実績）
    ev_forecast: Optional[str]  # EV（進行期見込み）
    
    # Entry Multiple
    entry_multiple_current: Optional[str] = None  # エントリーマルチプル（直近実績）
    entry_multiple_forecast: Optional[str] = None  # エントリーマルチプル（進行期見込み）
    
    # Industry Median Multiple
    industry_median_multiple_current: Optional[str] = None  # マルチプル業界中央値（直近実績）
    industry_median_multiple_forecast: Optional[str] = None  # マルチプル業界中央値（進行期見込み）


class SpeedaInput(BaseModel):
    industry: str
    sector: str
    category: str
    prompt: str
    query_type:str

class PerplexityInput(BaseModel):
    prompt: str
    query_type:str

    

## word出力モデル
class SectionSummaries(BaseModel):
    current_situation: str
    future_outlook: str
    investment_advantages: str
    investment_disadvantages: str
    value_up: str
    use_case: str
    swot_analysis: str

class Summaries(BaseModel):
    Perplexity: SectionSummaries
    ChatGPT: SectionSummaries

class ValuationItem(BaseModel):
    current: str
    forecast: str

class ValuationData(BaseModel):
    売上: ValuationItem
    EBITDA: ValuationItem
    NetDebt: ValuationItem
    想定EquityValue: ValuationItem
    EV: ValuationItem
    エントリーマルチプル: ValuationItem
    マルチプル業界中央値: ValuationItem

class WordExportRequest(BaseModel):
    summaries: Summaries
    valuation_data: Optional[ValuationData] = None
    
    

