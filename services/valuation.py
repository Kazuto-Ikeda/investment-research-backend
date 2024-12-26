from pydantic import BaseModel
from models.model import ValuationInput, ValuationOutput
from typing import Optional
from fastapi import HTTPException
from azure.storage.blob import BlobServiceClient  # 非同期クライアントから同期クライアントに変更
from docx import Document
import os
import tempfile
import logging
import unicodedata


# Azure Blob Storage設定
BLOB_CONNECTION_STRING = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
BLOB_CONTAINER_NAME = os.getenv("BLOB_CONTAINER_NAME")


# ユーティリティ関数
def format_number_with_commas(value: Optional[float]) -> Optional[str]:
    if value is None:
        return None
    try:
        return f"{int(round(value)):,}"
    except ValueError:
        return None

def format_number_with_x(value: Optional[float]) -> Optional[str]:
    if value is None:
        return None
    return f"{round(value, 1)}x"

# Unicode正規化関数（NFD形式で正規化）
def normalize_text(text: str) -> str:
    """文字列をNFD形式で正規化"""
    normalized = unicodedata.normalize('NFD', text)
    logging.debug(f"Original text: '{text}' | Normalized text: '{normalized}'")
    return normalized

# 計算関数
async def calculate_valuation(input_data: ValuationInput) -> ValuationOutput:
    logging.info(f"Starting valuation calculation with input: {input_data}")
    
    # 同期Blobサービスクライアントの初期化
    blob_service_client = BlobServiceClient.from_connection_string(os.getenv("AZURE_STORAGE_CONNECTION_STRING"))
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    temp_file_path = temp_file.name
    temp_file.close()
    industry_median_multiple_current = None

    try:
        # カテゴリー名の正規化（NFD形式）
        normalized_category = normalize_text(input_data.category)
        logging.info(f"Normalized category name: '{normalized_category}'")
        
        # categoryにすでに.docxが含まれているか確認
        if normalized_category.lower().endswith(".docx"):
            blob_name = normalized_category
        else:
            blob_name = f"{normalized_category}.docx"

        logging.info(f"Constructed blob name: '{blob_name}'")
        
        blob_client = blob_service_client.get_blob_client(container=os.getenv("BLOB_CONTAINER_NAME"), blob=blob_name)
        logging.info(f"アクセスするBlob名: '{blob_name}'")

        # Blobの存在確認
        blob_exists = blob_client.exists()
        logging.debug(f"Blob '{blob_name}' の存在: {blob_exists}")
        if not blob_exists:
            logging.error(f"指定されたBlob '{blob_name}' はコンテナ '{os.getenv('BLOB_CONTAINER_NAME')}' に存在しません。")
            raise HTTPException(status_code=404, detail="指定されたBlobファイルが存在しません。")

        # Blobストレージからファイルをダウンロード
        try:
            download_stream = blob_client.download_blob()
            data = download_stream.readall()
            with open(temp_file_path, "wb") as file:
                file.write(data)
            logging.info(f"Blob '{blob_name}' を一時ファイル '{temp_file_path}' にダウンロードしました。")
        except Exception as e:
            logging.error(f"Blobダウンロードエラー: {e}")
            raise HTTPException(status_code=500, detail="Blobファイルのダウンロードに失敗しました。")

        # Word文書を読み込む
        try:
            doc = Document(temp_file_path)
            logging.info(f"Word文書 '{blob_name}' の内容を読み込みました。")
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
                        logging.info(f"Found industry median multiple: {industry_median_multiple_current}")
                    except ValueError:
                        raise HTTPException(status_code=500, detail=f"無効な値: '{cell_value}' を数値に変換できません。")
                    break

            # 値が見つかった場合はループ終了
            if industry_median_multiple_current is not None:
                break

        if industry_median_multiple_current is None:
            raise HTTPException(status_code=404, detail=f"列 '{target_column}' または 行 '{target_row}' が見つかりませんでした。")

        logging.info("Calculating valuation metrics.")

        # 進行期見込みのマルチプルを設定（直近実績と同一）
        industry_median_multiple_forecast = industry_median_multiple_current
        logging.info(f"Industry median multiple forecast: {industry_median_multiple_forecast}")

        # EV（Enterprise Value）の計算
        ev_current = input_data.net_debt_current + input_data.equity_value_current
        ev_forecast = input_data.net_debt_current + input_data.equity_value_current  # 進行期見込みも同じ値
        logging.info(f"EV Current: {ev_current}, EV Forecast: {ev_forecast}")

        # エントリーマルチプルの計算（修正後）
        entry_multiple_current = (
            ev_current / input_data.ebitda_current
            if input_data.ebitda_current and input_data.ebitda_current != 0 else None
        )
        entry_multiple_forecast = (
            ev_forecast / input_data.ebitda_forecast
            if input_data.ebitda_forecast and input_data.ebitda_forecast != 0 else None
        )
        logging.info(f"Entry Multiple Current: {entry_multiple_current}, Entry Multiple Forecast: {entry_multiple_forecast}")

        # フォーマット済みの値を作成
        valuation_output = ValuationOutput(
            # 売上
            revenue_current=format_number_with_commas(input_data.revenue_current),
            revenue_forecast=format_number_with_commas(input_data.revenue_forecast),
            
            # EBITDA
            ebitda_current=format_number_with_commas(input_data.ebitda_current),
            ebitda_forecast=format_number_with_commas(input_data.ebitda_forecast),
            
            # Net Debt
            net_debt_current=format_number_with_commas(input_data.net_debt_current),
            net_debt_forecast=format_number_with_commas(input_data.net_debt_current),  # 同一値
            
            # Equity Value
            equity_value_current=format_number_with_commas(input_data.equity_value_current),
            equity_value_forecast=format_number_with_commas(input_data.equity_value_current),  # 同一値
            
            # Enterprise Value (EV)
            ev_current=format_number_with_commas(ev_current),
            ev_forecast=format_number_with_commas(ev_forecast),
            
            # Entry Multiple
            entry_multiple_current=format_number_with_x(entry_multiple_current),
            entry_multiple_forecast=format_number_with_x(entry_multiple_forecast),
            
            # Industry Median Multiple
            industry_median_multiple_current=format_number_with_x(industry_median_multiple_current),
            industry_median_multiple_forecast=format_number_with_x(industry_median_multiple_forecast),
        )
        
        logging.info(f"Valuation output: {valuation_output}")
        return valuation_output

    except HTTPException as he:
        logging.error(f"HTTPException in calculate_valuation: {he.detail}")
        raise he
    except Exception as e:
        logging.error(f"Blobストレージまたは計算処理中のエラー: {e}")
        raise HTTPException(status_code=500, detail="エラーが発生しました。再試行してください。")
    finally:
        # 一時ファイルの削除
        if os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
                logging.info(f"一時ファイル '{temp_file_path}' を削除しました。")
            except Exception as e:
                logging.warning(f"一時ファイルの削除に失敗しました: {e}")