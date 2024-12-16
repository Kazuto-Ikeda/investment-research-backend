from pydantic import BaseModel
from models.model import ValuationInput, ValuationOutput
from typing import Optional
from fastapi import HTTPException
from azure.storage.blob.aio import BlobServiceClient as AsyncBlobServiceClient
from docx import Document
import os
import tempfile
import logging


# Azure Blob Storage設定
BLOB_CONNECTION_STRING = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
BLOB_CONTAINER_NAME = os.getenv("BLOB_CONTAINER_NAME")


# ユーティリティ関数
def format_number_with_x(value: Optional[float]) -> Optional[str]:
    if value is None:
        return None
    return f"{round(value, 1)}x"

# 計算関数
async def calculate_valuation(input_data: ValuationInput) -> ValuationOutput:
    print("Starting valuation calculation with input:", input_data)
    
    # 非同期Blobサービスクライアントの初期化
    blob_service_client = AsyncBlobServiceClient.from_connection_string(BLOB_CONNECTION_STRING)
    blob_name = f"{input_data.category}.docx"
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    temp_file_path = temp_file.name
    temp_file.close()
    industry_median_multiple_current = None

    try:
        print(f"Accessing Blob: {blob_name}")
        async with AsyncBlobServiceClient.from_connection_string(BLOB_CONNECTION_STRING) as blob_service_client:
            blob_client = blob_service_client.get_blob_client(container=BLOB_CONTAINER_NAME, blob=blob_name)
            logging.info(f"アクセスするBlob名: {blob_name}")

            # Blobストレージからファイルを非同期にダウンロード
            try:
                download_stream = await blob_client.download_blob()
                data = await download_stream.readall()
                with open(temp_file_path, "wb") as file:
                    file.write(data)
                print(f"Downloaded Blob to {temp_file_path}")
            except Exception as e:
                logging.error(f"Blobダウンロードエラー: {e}")
                raise HTTPException(status_code=500, detail="Blobファイルのダウンロードに失敗しました。")

        # Word文書を読み込む
        try:
            doc = Document(temp_file_path)
            print("Loaded Word document.")
        except Exception as e:
            logging.error(f"Word文書の読み込みエラー: {e}")
            raise HTTPException(status_code=500, detail="Word文書の読み込みに失敗しました。")


        # 取得したい列名と行名
        target_column = "企業価値/EBITDA"
        target_row = "中央値"

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
        ev_forecast = input_data.net_debt_current + (input_data.equity_value_current * 1.1)
        print(f"EV Current: {ev_current}, EV Forecast: {ev_forecast}")

        # エントリーマルチプルの計算
        entry_multiple_current = (
            ev_current / input_data.ebitda_current
            if input_data.ebitda_current and input_data.ebitda_current > 0 else None
        )
        entry_multiple_forecast = (
            ev_forecast / input_data.ebitda_forecast
            if input_data.ebitda_forecast and input_data.ebitda_forecast > 0 else None
        )
        print(f"Entry Multiple Current: {entry_multiple_current}, Entry Multiple Forecast: {entry_multiple_forecast}")

        # フォーマット済みの値を作成
        valuation_output = ValuationOutput(
            # 売上
            revenue_current=int(input_data.revenue_current),
            revenue_forecast=int(input_data.revenue_forecast),
            
            # EBITDA
            ebitda_current=int(input_data.ebitda_current) if input_data.ebitda_current is not None else None,
            ebitda_forecast=int(input_data.ebitda_forecast) if input_data.ebitda_forecast is not None else None,
            
            # Net Debt
            net_debt_current=int(input_data.net_debt_current),
            net_debt_forecast=int(input_data.net_debt_current),  # NetDebt（進行期見込）は同じ値
            
            # Equity Value
            equity_value_current=int(input_data.equity_value_current),
            equity_value_forecast=int(input_data.equity_value_current * 1.1),
            
            # Enterprise Value (EV)
            ev_current=int(ev_current),
            ev_forecast=int(ev_forecast),
            
            # Entry Multiple
            entry_multiple_current=format_number_with_x(entry_multiple_current),
            entry_multiple_forecast=format_number_with_x(entry_multiple_forecast),
            
            # Industry Median Multiple
            industry_median_multiple_current=format_number_with_x(industry_median_multiple_current),
            industry_median_multiple_forecast=format_number_with_x(industry_median_multiple_forecast),
        )
        
        return valuation_output

    except HTTPException as he:
        logging.error(f"HTTPException in calculate_valuation: {he.detail}")
        raise he
    except Exception as e:
        logging.error(f"Blobストレージまたは要約処理中のエラー: {e}")
        raise HTTPException(status_code=500, detail="エラーが発生しました。再試行してください。")
    finally:
        # 一時ファイルの削除
        if os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
                print(f"Temporary file {temp_file_path} deleted.")
            except Exception as e:
                logging.warning(f"一時ファイルの削除に失敗しました: {e}")