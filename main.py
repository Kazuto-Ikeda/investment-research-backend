from fastapi import FastAPI, HTTPException, BackgroundTasks, Query, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, FileResponse
from models.model import RegenerateRequest
from models.model import ValuationInput, ValuationOutput
from services import summarize
from services import valuation
from services.word_export import generate_word_file
from services.valuation import calculate_valuation
from typing import Optional
import logging
import unicodedata
import httpx
from services.summarize import (
    download_blob_to_temp_file,
    unison_summary_logic,
    regenerate_summary,
)

app = FastAPI()


#通信設定
origins = [
    "*"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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
        final_summary = await unison_summary_logic(
            query_key=query_key,
            company_name=company_name,
            industry=industry,
            chatgpt_summary=chatgpt_summary,
        )
        return {"status": "success", "final_summary": final_summary}
    except HTTPException as he:
        raise he
    except Exception as e:
        logging.error(f"エンドポイント処理中のエラー: {e}")
        raise HTTPException(status_code=500, detail="エンドポイント処理中にエラーが発生しました。")


@app.post("/valuation", response_model=ValuationOutput)
async def valuation_endpoint(request: ValuationInput):
    """
    バリュエーション計算エンドポイント
    """
    print("Received valuation request:", request)
    try:
        # 計算を実行
        valuation_result = calculate_valuation(
            input_data=request
        )
        print("Valuation result:", valuation_result)
        return valuation_result
    except Exception as e:
        print("Error in valuation endpoint:", e)
        raise HTTPException(status_code=400, detail=str(e))
    
        
@app.post("/word_export")
async def export_endpoint(
    background_tasks: BackgroundTasks,
    summaries: dict = Body(..., description="要約データを含む辞書形式の入力"),
    valuation_data: Optional[dict] = Body(None, description="バリュエーションデータ"),
    company_name: str = Query(..., description="会社名を指定"),
    file_name: Optional[str] = Query(None, description="生成するWordファイル名 (省略可能)")
):
    """
    Wordファイル生成エンドポイント
    """
    return generate_word_file(
        background_tasks, summaries, valuation_data, company_name, file_name
    )

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
        
        # 必須フィールドのチェック
        if not all([industry, sector, category, company_name, query_key]):
            raise HTTPException(status_code=400, detail="必要なフィールドが不足しています。")
            
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