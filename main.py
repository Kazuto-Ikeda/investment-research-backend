from fastapi import FastAPI, HTTPException, BackgroundTasks, Query, Body, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, FileResponse
from models.model import ValuationInput, ValuationOutput, SpeedaInput,PerplexityInput
# from services.summarize import download_blob_to_temp_file
# from services.summarize import unison_summary_logic
# from services.summarize import get_perplexity_summary
# from services.summarize import regenerate_summary
from services.summarize import normalize_text
from services.summarize import summary_from_speeda
from services.word_export import generate_word_file
from services.summarize import clean_text
from models.model import WordExportRequest

from services.valuation import calculate_valuation
from typing import Optional
from openai import OpenAI

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

        
                
@app.post("/summarize/speeda")
def summary_speeda(request: SpeedaInput):
    """
    Blobストレージ -> 要約生成
    """
    try:
        # Blobストレージからファイルをダウンロードし、要約を生成
        try:
            summary = summary_from_speeda(
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
        return JSONResponse(content={request.query_type: summary}, status_code=200)

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

