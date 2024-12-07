from pydantic import BaseModel
from models.model import ValuationInput
from typing import Optional

# 計算関数
def calculate_valuation(input_data: ValuationInput):
    """
    入力データに基づいて評価指標を計算し、詳細なレスポンスを返す
    """
        # 計算処理
        # EVの計算
    ev_current = input_data["net_debt_current"] + input_data["equity_value_current"]
    ev_forecast = input_data["net_debt_forecast"] + input_data["equity_value_forecast"]

    # Entry Multiple (EV/EBITDA) の計算
    entry_multiple_current = (
        ev_current / input_data["ebitda_current"] if input_data["ebitda_current"] > 0 else None
    )
    entry_multiple_forecast = (
        ev_forecast / input_data["ebitda_forecast"] if input_data["ebitda_forecast"] > 0 else None
    )

    # レスポンスデータの構築
    response = {
        "inputs": input_data,  # ユーザーが送信したデータをそのまま含める
        "calculations": {
            "ev_current": ev_current,
            "ev_forecast": ev_forecast,
            "entry_multiple_current": entry_multiple_current,
            "entry_multiple_forecast": entry_multiple_forecast,
            "industry_median_multiple_current": input_data["industry_median_multiple"],
            "industry_median_multiple_forecast": input_data["industry_median_multiple_forecast"],
            "current_comparison": {
                "entry_multiple_vs_median": (
                    entry_multiple_current - input_data["industry_median_multiple"]
                    if entry_multiple_current is not None else None
                )
            },
            "forecast_comparison": {
                "entry_multiple_vs_median": (
                    entry_multiple_forecast - input_data["industry_median_multiple_forecast"]
                    if entry_multiple_forecast is not None else None
                )
            }
        },
        "details": {
            "calculation_steps": [
                {
                    "name": "EV (Enterprise Value)",
                    "formula": "Net Debt + Equity Value",
                    "current_value": ev_current,
                    "forecast_value": ev_forecast
                },
                {
                    "name": "Entry Multiple",
                    "formula": "EV / EBITDA",
                    "current_value": entry_multiple_current,
                    "forecast_value": entry_multiple_forecast
                },
            ]
        }
    }

    return response
