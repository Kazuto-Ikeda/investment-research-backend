from fastapi import FastAPI
from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from models.model import RegenerateRequest
from services import summarize
from services import valuation
from services.word_export import generate_word_file
from services.valuation import calculate_valuation
import logging
import unicodedata
from services.summarize import (
    download_blob_to_temp_file,
    unison_summary_logic,
    regenerate_summary,
)

app = FastAPI()

logging.basicConfig(level=logging.INFO)

@app.post("/summarize")
async def summary_endpoint(request: dict):
    """
    Blobストレージ -> 要約生成 (任意でPerplexity補足情報と統合要約を実行)
    """
    try:
        def normalize_text(text: str) -> str:
            """文字列をNFCで正規化"""
            return unicodedata.normalize('NFC', text)

        # リクエストデータの取得とバリデーション
        industry = request.get("industry")
        sector = request.get("sector")
        category = request.get("category")
        blob_name = normalize_text(category) + ".docx"  # 小分類に .docx を追加
        company_name = request.get("company_name")
        include_perplexity = request.get("include_perplexity", False)  # デフォルトはFalse

        # 必須フィールドのチェック
        missing_fields = []
        if not industry:
            missing_fields.append("industry")
        if not sector:
            missing_fields.append("sector")
        if not category:
            missing_fields.append("category")
        if not company_name:
            missing_fields.append("company_name")

        if missing_fields:
            logging.error(f"リクエストに不足しているフィールド: {missing_fields}")
            raise HTTPException(
                status_code=400,
                detail=f"必要なフィールドが不足しています: {', '.join(missing_fields)}"
            )

        # Blobストレージからファイルをダウンロードし、要約を生成
        try:
            summaries = download_blob_to_temp_file(
                category=category,
                company_name=company_name,
            )
        except HTTPException as e:
            logging.error(f"Blobストレージまたは要約処理中のエラー: {e.detail}")
            raise e
        except Exception as e:
            logging.error(f"エンドポイント処理中の予期しないエラー: {e}")
            raise HTTPException(
                status_code=500,
                detail="エンドポイント処理中にエラーが発生しました。"
            )

        # 結果を返す
        return {"summaries": summaries}

    except HTTPException as e:
        logging.error(f"HTTPエラー: {e.detail}")
        raise e
    except Exception as e:
        logging.error(f"エンドポイント全体の予期しないエラー: {e}")
        raise HTTPException(
            status_code=500,
            detail="エンドポイント全体の処理中にエラーが発生しました。"
        )        

@app.post("/summarize/perplexity")
async def unison_summary(request: dict):
    """
    2工程目と3工程目: Perplexityと統合要約
    """
    try:
        query_key = request.get("query_key")
        company_name = request.get("company_name")
        industry = request.get("industry")
        chatgpt_summary = request.get("chatgpt_summary")  # 初回要約を入力

        if not query_key or not company_name or not industry or not chatgpt_summary:
            raise HTTPException(status_code=400, detail="必要なフィールドが不足しています。")

        # Perplexityと統合要約を処理
        final_summary = unison_summary_logic(
            query_key=query_key,
            company_name=company_name,
            industry=industry,
            chatgpt_summary=chatgpt_summary,
        )
        return {"status": "success", "final_summary": final_summary}
    except Exception as e:
        logging.error(f"エンドポイント処理中のエラー: {e}")
        raise HTTPException(status_code=500, detail="エンドポイント処理中にエラーが発生しました。")


@app.post("/valuation")
async def valuation_endpoint(request: dict):
    return calculate_valuation(request)

@app.get("/word_export")
async def export_endpoint():
    return word_export()

@app.post("/regenerate-summary")
async def user_regenerate(request: dict):
        try:
            def normalize_text(text: str) -> str:
                return unicodedata.normalize('NFC', text)

            # リクエストデータの取得とバリデーション
            industry = request.get("industry")
            sector = request.get("sector")
            category = request.get("category")
            blob_name = normalize_text(category) + ".docx"  # 小分類に .docx を追加
            company_name = request.get("company_name")
            include_perplexity = request.get("include_perplexity", False)  # デフォルトはFalse
            query_key = request.get("query_key")
            custom_query = request.get("custom_query")
            perplexity_summary = request.get("perplexity_summary")
            
        except HTTPException as e:
            logging.error(f"再要約処理中のエラー: {e.detail}")
            raise e
        except Exception as e:
            logging.error(f"エンドポイント処理中の予期しないエラー: {e}")
            raise HTTPException(
                status_code=500,
                detail="エンドポイント処理中にエラーが発生しました。"
            )
        return regenerate_summary(category, company_name, query_key, perplexity_summary, custom_query, include_perplexity)
    
    
@app.post("/")
async def api_test():
    return print("Hello, world!")


# @app.post("/summarize")
# async def summary_endpoint(request: dict):
#     """
#     Blobストレージ -> 要約生成 (任意でPerplexity補足情報と統合要約を実行)
#     """
#     try:
#         def normalize_text(text) -> str:
#             """文字列をNFCで正規化"""
#             return unicodedata.normalize('NFC', text)

#         # リクエストデータの取得とバリデーション
#         industry = request.get("industry")
#         sector = request.get("sector")
#         category = request.get("category")
#         blob_name = normalize_text(category) + ".docx"  # 小分類に .docx を追加
#         company_name = request.get("company_name")
#         include_perplexity = request.get("include_perplexity", False)  # デフォルトはFalse

#         # 必須フィールドのチェック
#         missing_fields = []
#         if not industry:
#             missing_fields.append("industry")
#         if not sector:
#             missing_fields.append("sector")
#         if not category:
#             missing_fields.append("category")
#         if not company_name:
#             missing_fields.append("company_name")

#         if missing_fields:
#             logging.error(f"リクエストに不足しているフィールド: {missing_fields}")
#             raise HTTPException(
#                 status_code=400,
#                 detail=f"必要なフィールドが不足しています: {', '.join(missing_fields)}"
#             )

#         # Blobストレージからファイルをダウンロードし、要約を生成
#         try:
#             summaries = download_blob_to_temp_file(
#                 category=category,
#                 company_name=company_name,
#                 industry=industry,
#                 include_perplexity=include_perplexity,
#                 query="業界の現状は？"
#             )
#         except HTTPException as e:
#             logging.error(f"Blobストレージまたは要約処理中のエラー: {e.detail}")
#             raise e
#         except Exception as e:
#             logging.error(f"エンドポイント処理中の予期しないエラー: {e}")
#             raise HTTPException(
#                 status_code=500,
#                 detail="エンドポイント処理中にエラーが発生しました。"
#             )

#         # 結果を返す
#         return {"summaries": summaries}

#     except HTTPException as e:
#         logging.error(f"HTTPエラー: {e.detail}")
#         raise e
#     except Exception as e:
#         logging.error(f"エンドポイント全体の予期しないエラー: {e}")
#         raise HTTPException(
#             status_code=500,
#             detail="エンドポイント全体の処理中にエラーが発生しました。"
#         )