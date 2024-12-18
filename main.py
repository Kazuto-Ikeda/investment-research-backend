from fastapi import FastAPI, HTTPException, BackgroundTasks, Query, Body, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, FileResponse
from models.model import ValuationInput, ValuationOutput, SpeedaInput,PerplexityInput
from services.summarize import download_blob_to_temp_file
from services.summarize import unison_summary_logic
from services.summarize import get_perplexity_summary
from services.summarize import regenerate_summary
from services.word_export import generate_word_file
from models.model import WordExportRequest

from services.valuation import calculate_valuation
from typing import Optional
import logging
import sys
import unicodedata
import os
import uvicorn
import httpx
from services.summarize import (
    summary_from_speeda,
    perplexity_search,
)



app = FastAPI()

# #port指定
# if __name__ == "__main__":
#     port = int(os.getenv("WEBSITES_PORT", 8000))
#     uvicorn.run(app, host="0.0.0.0", port=port)

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

# ログの設定
logger = logging.getLogger("investment-backend")
logger.setLevel(logging.INFO)

# ストリームハンドラーを追加（標準出力）
handler = logging.StreamHandler(sys.stdout)
handler.setLevel(logging.INFO)

# ログフォーマットを設定
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)

# ハンドラーをロガーに追加
logger.addHandler(handler)

        
# ミドルウェアでリクエストをログに記録
@app.middleware("http")
async def log_requests(request: Request, call_next):
    logger.info(f"Received request: {request.method} {request.url}")
    try:
        response = await call_next(request)
        return response
    except Exception as e:
        logger.error(f"Error processing request: {e}")
        raise e

#使ってない
@app.post("/summarize")
async def summary_endpoint(request: dict):
    """
    Blobストレージ -> 要約生成 (任意でPerplexity補足情報と統合要約を実行
    """
    try:
        def normalize_text(text: str) -> str:
            """文字列をNFCで正規化"""
            return unicodedata.normalize('NFC', text)

        # リクエストデータの取得とバリデーション
        industry = request.get("industry")
        sector = request.get("sector")
        category = request.get("category")
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
            summaries = await download_blob_to_temp_file(
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
        
                
@app.post("/summarize/speeda")
async def summary_speeda(request: SpeedaInput):
    """
    Blobストレージ -> 要約生成
    """
    try:
        def normalize_text(text: str) -> str:
            """文字列をNFCで正規化"""
            return unicodedata.normalize('NFC', text)

        # Blobストレージからファイルをダウンロードし、要約を生成
        try:
            summary = await summary_from_speeda(
                category=request.category,
                prompt=request.prompt)
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
        return JSONResponse(content={request.query_type:summary}, status_code=200)

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
async def unison_summary(request: PerplexityInput):
    """
    2工程目と3工程目: Perplexityと統合要約
    """
    try:
        # Perplexityと統合要約を処理
        perplexity_result = perplexity_search(
            prompt=request.prompt,
        )
        return JSONResponse(content={request.query_type:perplexity_result}, status_code=200)
    except Exception as e:
        logging.error(f"エンドポイント処理中のエラー: {e}")
        raise HTTPException(status_code=500, detail="エンドポイント処理中にエラーが発生しました。")   
    
# エンドポイント
@app.post("/valuation", response_model=ValuationOutput)
async def valuation_endpoint(request: ValuationInput):
    """
    バリュエーション計算エンドポイント
    """
    logging.info(f"Received valuation request: {request}")
    try:
        # 計算を実行（awaitを追加）
        valuation_result = await calculate_valuation(
            input_data=request
        )
        logging.info(f"Valuation result: {valuation_result}")
        return valuation_result
    except HTTPException as he:
        logging.error(f"HTTPException in valuation endpoint: {he.detail}")
        raise he
    except Exception as e:
        logging.error(f"Error in valuation endpoint: {e}")
        raise HTTPException(status_code=400, detail=str(e))
            
        
@app.post("/word_export")
async def export_endpoint(
    background_tasks: BackgroundTasks,
    request: WordExportRequest,
    company_name: str = Query(..., description="会社名を指定"),
    file_name: Optional[str] = Query(None, description="生成するWordファイル名 (省略可能)")
):
    """
    Wordファイル生成エンドポイント
    """
    return generate_word_file(
        background_tasks, request.summaries.dict(), request.valuation_data.dict() if request.valuation_data else None, company_name, file_name
    )    

#使ってない
@app.post("/regenerate-summary")
async def user_regenerate(request: dict):
    """
    /regenerate-summary エンドポイント
    """
    try:
        def normalize_text(text: Optional[str]) -> str:
            """文字列をNFCで正規化"""
            if not text:
                raise ValueError("正規化対象のテキストがNoneです。")
            return unicodedata.normalize('NFC', text)

        # リクエストデータの取得とバリデーション
        industry = request.get("industry")
        sector = request.get("sector")
        category = request.get("category")
        company_name = request.get("company_name")
        include_perplexity = request.get("include_perplexity", False)  # デフォルトはFalse
        query_key = request.get("query_key")
        custom_query = request.get("custom_query")
        perplexity_summary = request.get("perplexity_summary")
        
        # 必須フィールドのバリデーション
        if not all([industry, sector, category, company_name, query_key]):
            missing_fields = [field for field in ["industry", "sector", "category", "company_name", "query_key"]
                              if request.get(field) is None]
            logging.error(f"リクエストに不足しているフィールド: {missing_fields}")
            raise HTTPException(
                status_code=400,
                detail=f"必要なフィールドが不足しています: {', '.join(missing_fields)}"
            )

        # 正規化処理
        # Note: `regenerate_summary` 関数内で Blob 名を処理するため、エンドポイント側で `blob_name` を設定する必要はありません
        normalized_category = normalize_text(category)

    except ValueError as ve:
        logging.error(f"値エラー: {ve}")
        raise HTTPException(status_code=400, detail=str(ve))
    except HTTPException as he:
        logging.error(f"HTTPエラー: {he.detail}")
        raise he
    except Exception as e:
        logging.error(f"エンドポイント処理中の予期しないエラー: {e}")
        raise HTTPException(
            status_code=500,
            detail="エンドポイント処理中に予期しないエラーが発生しました。"
        )
    
    # regenerate_summary を await して呼び出す
    try:
        final_summary_data = await regenerate_summary(
            category_name=category,
            company_name=company_name,
            query_key=query_key,
            perplexity_summary=perplexity_summary,
            custom_query=custom_query,
            include_perplexity=include_perplexity
        )
    except HTTPException as he:
        logging.error(f"regenerate_summary 内でのHTTPエラー: {he.detail}")
        raise he
    except Exception as e:
        logging.error(f"regenerate_summary 内での予期しないエラー: {e}")
        raise HTTPException(
            status_code=500,
            detail="regenerate_summary 処理中にエラーが発生しました。"
        )
    
    return {"status": "success", "final_summary": final_summary_data["final_summary"]}


