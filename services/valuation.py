from pydantic import BaseModel
from models.model import ValuationInput, ValuationOutput
from typing import Optional
from fastapi import HTTPException
from azure.storage.blob import BlobServiceClient
from docx import Document
import os
import tempfile
import logging


# Azure Blob Storage設定
BLOB_CONNECTION_STRING = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
BLOB_CONTAINER_NAME = os.getenv("BLOB_CONTAINER_NAME")



# 計算関数

# def calculate_valuation(input_data: ValuationInput):
#     """
#     入力データに基づいて評価指標を計算し、詳細なレスポンスを返す
#     """
#         # 計算処理
#         # EVの計算
#     ev_current = input_data["net_debt_current"] + input_data["equity_value_current"]
#     ev_forecast = input_data["net_debt_forecast"] + input_data["equity_value_forecast"]

#     # Entry Multiple (EV/EBITDA) の計算
#     entry_multiple_current = (
#         ev_current / input_data["ebitda_current"] if input_data["ebitda_current"] > 0 else None
#     )
#     entry_multiple_forecast = (
#         ev_forecast / input_data["ebitda_forecast"] if input_data["ebitda_forecast"] > 0 else None
#     )

#     # レスポンスデータの構築
#     response = {
#         "inputs": input_data,  # ユーザーが送信したデータをそのまま含める
#         "calculations": {
#             "ev_current": ev_current,
#             "ev_forecast": ev_forecast,
#             "entry_multiple_current": entry_multiple_current,
#             "entry_multiple_forecast": entry_multiple_forecast,
#             "industry_median_multiple_current": input_data["industry_median_multiple"],
#             "industry_median_multiple_forecast": input_data["industry_median_multiple_forecast"],
#             "current_comparison": {
#                 "entry_multiple_vs_median": (
#                     entry_multiple_current - input_data["industry_median_multiple"]
#                     if entry_multiple_current is not None else None
#                 )
#             },
#             "forecast_comparison": {
#                 "entry_multiple_vs_median": (
#                     entry_multiple_forecast - input_data["industry_median_multiple_forecast"]
#                     if entry_multiple_forecast is not None else None
#                 )
#             }
#         },
#         "details": {
#             "calculation_steps": [
#                 {
#                     "name": "EV (Enterprise Value)",
#                     "formula": "Net Debt + Equity Value",
#                     "current_value": ev_current,
#                     "forecast_value": ev_forecast
#                 },
#                 {
#                     "name": "Entry Multiple",
#                     "formula": "EV / EBITDA",
#                     "current_value": entry_multiple_current,
#                     "forecast_value": entry_multiple_forecast
#                 },
#             ]
#         }
#     }

#     return response

# def valuation_endpoint(result: dict):
#     try:
#         # 必要な値を辞書から取得（デフォルト値はエラーを回避するために設定）
#         industry_median_multiple_current = result.get("industry_median_multiple_current")
#         industry_median_multiple_forecast = result.get("industry_median_multiple_forecast")
        
#         # 値が不足している場合はエラーをスロー
#         if industry_median_multiple_current is None or industry_median_multiple_forecast is None:
#             raise HTTPException(status_code=400, detail="Industry median multiple values are missing.")

#         # 辞書を ValuationInput モデルに変換
#         input_data = ValuationInput(**result)
        
#         # 計算関数に必要な値を渡す
#         return calculate_valuation(input_data, industry_median_multiple_current, industry_median_multiple_forecast)
#     except Exception as e:
#         raise HTTPException(status_code=400, detail=str(e))
    
    
def calculate_valuation(input_data: ValuationInput) -> ValuationOutput:
    print("Starting valuation calculation with input:", input_data)
    blob_service_client = BlobServiceClient.from_connection_string(BLOB_CONNECTION_STRING)
    blob_name = f"{input_data.category}.docx"
    temp_file_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    text = ""

    try:
        print(f"Accessing Blob: {blob_name}")
        blob_client = blob_service_client.get_blob_client(container=BLOB_CONTAINER_NAME, blob=blob_name)
        logging.info(f"アクセスするBlob名: {blob_name}")

        # Blobストレージからファイルをダウンロード
        with open(temp_file_path, "wb") as file:
            download_stream = blob_client.download_blob()
            file.write(download_stream.readall())
        print(f"Downloaded Blob to {temp_file_path}")

        # Word文書を読み込む
        doc = Document(temp_file_path)
        print("Loaded Word document.")

        # 取得したい列名と行名
        target_column = "企業価値/EBITDA"
        target_row = "平均値"

        industry_median_multiple_current = None

        # 文書内のテーブルを探索
        for table in doc.tables:
            # 列インデックスとヘッダーを特定
            column_index = None
            for header_index, header_cell in enumerate(table.rows[0].cells):
                if header_cell.text.strip() == target_column:
                    column_index = header_index
                    break

            # 指定された列が存在しない場合は次のテーブルへ
            if column_index is None:
                continue

            # 指定された行名を検索
            for row in table.rows:
                if row.cells[0].text.strip() == target_row:  # 行名は通常1列目に存在
                    cell_value = row.cells[column_index].text.strip()
                    try:
                        # "倍" を削除して数値変換
                        industry_median_multiple_current = float(cell_value.replace("倍", "").strip())
                        print(f"Found industry median multiple: {industry_median_multiple_current}")
                    except ValueError:
                        raise HTTPException(status_code=500, detail=f"無効な値: '{cell_value}' を数値に変換できません。")
                    break

            # 値が見つかった場合はループ終了
            if industry_median_multiple_current is not None:
                break

        if industry_median_multiple_current is None:
            raise HTTPException(status_code=404, detail=f"列 '{target_column}' または 行 '{target_row}' が見つかりませんでした。")

        print("Calculating valuation metrics.")

        # 進行期見込みのマルチプルを設定（仮に現状のデータに基づいて予測する例）
        industry_median_multiple_forecast = industry_median_multiple_current * 1.1
        print(f"Industry median multiple forecast: {industry_median_multiple_forecast}")

        # EV（Enterprise Value）の計算
        ev_current = input_data.net_debt_current + input_data.equity_value_current
        ev_forecast = (
            input_data.net_debt_current + input_data.equity_value_current
            if input_data.ebitda_forecast is not None else None
        )
        print(f"EV Current: {ev_current}, EV Forecast: {ev_forecast}")

        # エントリーマルチプルの計算
        entry_multiple_current = (
            ev_current / input_data.ebitda_current
            if input_data.ebitda_current and input_data.ebitda_current > 0 else None
        )
        entry_multiple_forecast = (
            ev_forecast / input_data.ebitda_forecast
            if ev_forecast and input_data.ebitda_forecast and input_data.ebitda_forecast > 0 else None
        )
        print(f"Entry Multiple Current: {entry_multiple_current}, Entry Multiple Forecast: {entry_multiple_forecast}")

        # Implied Equity Valueの計算
        implied_equity_value_current = (
            industry_median_multiple_current * input_data.ebitda_current
            if industry_median_multiple_current and input_data.ebitda_current else None
        )
        implied_equity_value_forecast = (
            industry_median_multiple_forecast * input_data.ebitda_forecast
            if industry_median_multiple_forecast and input_data.ebitda_forecast else None
        )
        print(f"Implied Equity Value Current: {implied_equity_value_current}, Forecast: {implied_equity_value_forecast}")

        return ValuationOutput(
            ebitda_current=input_data.ebitda_current,
            ebitda_forecast=input_data.ebitda_forecast,
            net_debt_current=input_data.net_debt_current,
            net_debt_forecast=None,  # NetDebt（進行期見込み）は不要
            equity_value_current=input_data.equity_value_current,
            equity_value_forecast=input_data.equity_value_current,  # 同じ値を返す
            ev_current=ev_current,
            ev_forecast=ev_forecast,
            entry_multiple_current=entry_multiple_current,
            entry_multiple_forecast=entry_multiple_forecast,
            industry_median_multiple_current=industry_median_multiple_current,
            industry_median_multiple_forecast=industry_median_multiple_forecast,
            implied_equity_value_current=implied_equity_value_current,
            implied_equity_value_forecast=implied_equity_value_forecast,
        )
    except Exception as e:
        logging.error(f"Blobストレージまたは要約処理中のエラー: {e}")
        raise HTTPException(status_code=500, detail="エラーが発生しました。再試行してください。")
    finally:
        # 一時ファイルの削除
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)