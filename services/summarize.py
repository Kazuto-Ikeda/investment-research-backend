from fastapi import HTTPException, Query
from fastapi.responses import JSONResponse
from docx import Document
from dotenv import load_dotenv
# from azure.storage.blob import BlobServiceClient
# from azure.storage.blob.aio import BlobServiceClient
from azure.storage.blob.aio import BlobServiceClient as AsyncBlobServiceClient
from pydantic import BaseModel
from typing import Optional, Dict
from models.model import RegenerateRequest
import pymysql
import os
import logging
import httpx
import openai
import uvicorn
import requests
import tempfile

# ロギング設定
logging.basicConfig(level=logging.INFO)



# 環境変数の読み込み
load_dotenv()

# APIキーの取得
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
PERPLEXITY_API_KEY = os.getenv("PERPLEXITY_API_KEY")
PERPLEXITY_API_ENDPOINT = os.getenv("PERPLEXITY_API_ENDPOINT")

# Azure Blob Storage設定
BLOB_CONNECTION_STRING = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
BLOB_CONTAINER_NAME = os.getenv("BLOB_CONTAINER_NAME")



# # MySQL設定
# MYSQL_HOST = os.getenv("MYSQL_HOST")
# MYSQL_USER = os.getenv("MYSQL_USER")
# MYSQL_PASSWORD = os.getenv("MYSQL_PASSWORD")
# MYSQL_DB = os.getenv("MYSQL_DB")


# def get_mysql_connection():
#     """MySQL接続"""
#     return pymysql.connect(
#         host=MYSQL_HOST,
#         user=MYSQL_USER,
#         password=MYSQL_PASSWORD,
#         database=MYSQL_DB,
#         ssl={"ca": "/DigiCertGlobalRootCA.crt.pem"} 
#     )

# def get_blob_url_from_mysql(industry: str, sector: str, category: str) -> str:
#     """MySQLから業種・業界・カテゴリに対応するBlob URLを取得"""
#     connection = get_mysql_connection()
#     with connection.cursor() as cursor:
#         query = 'SELECT Blob名前 FROM industry_data WHERE 業界大分類 = "建設" AND 業界中分類 = "インフラ建設" AND 業界小分類 = "土木工事"'
        
#         # クエリ実行前のデバッグログ
#         logging.info(f"実行するクエリ: {query}")
#         logging.info(f"クエリパラメータ: industry={industry}, sector={sector}, category={category}")
        
        
#         # クエリ実行
#         cursor.execute(query)
        
        
        
#         # クエリ結果を取得
#         result = cursor.fetchone()
        
#         # デバッグ: クエリ結果をログに記録
#         logging.info(f"クエリ結果: {result}")
        
#         # Blob名前を取得
#         blob_name = result[0]
    
#         # Blob Storageの完全URLを生成
#         blob_service_client = BlobServiceClient.from_connection_string(BLOB_CONNECTION_STRING)
#         blob_client = blob_service_client.get_blob_client(container=BLOB_CONTAINER_NAME, blob=blob_name)
        
#         # URL取得
#         blob_url = blob_client.url

#         # デバッグ: URLをログに記録
#         logging.info(f"生成されたBlob URL: {blob_url}")
        
#         # Blob をダウンロードして一時ファイルに保存
#         blob_data = blob_client.download_blob().readall()
#         with open('temp.docx', 'wb') as temp_file:
#             temp_file.write(blob_data)

#         # Word ファイルを読み取る
#         doc = Document('temp.docx')
#         for paragraph in doc.paragraphs:
#             print(paragraph.text)  # 各段落のテキストを出力
        
        


async def download_blob_to_temp_file(category: str, company_name: str) -> Dict[str, str]:
    """
    Blobストレージからファイルをダウンロードし、一時ファイルとして保存。
    小分類名に基づいて .docx ファイルを検索します。
    要約を生成して返します。
    """
    temp_file_path = None  # 初期化
    try:
        # 非同期Blobサービスクライアントの初期化
        blob_service_client = BlobServiceClient.from_connection_string(BLOB_CONNECTION_STRING)
        
        # categoryにすでに.docxが含まれているか確認
        if category.lower().endswith(".docx"):
            blob_name = category
        else:
            blob_name = f"{category}.docx"
        
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        temp_file_path = temp_file.name
        temp_file.close()
        text = ""

        # Blobクライアントを取得
        blob_client = blob_service_client.get_blob_client(container=BLOB_CONTAINER_NAME, blob=blob_name)
        logging.info(f"アクセスするBlob名: {blob_name}")

        # Blobストレージからファイルを非同期にダウンロード
        try:
            download_stream = await blob_client.download_blob()
            data = await download_stream.readall()
            with open(temp_file_path, "wb") as file:
                file.write(data)
        except Exception as e:
            logging.error(f"Blobダウンロードエラー: {e}")
            raise HTTPException(status_code=500, detail="Blobファイルのダウンロードに失敗しました。")

        # Word文書を読み込み、段落を結合
        try:
            doc = Document(temp_file_path)
            text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
        except Exception as e:
            logging.error(f"Word文書の読み込みエラー: {e}")
            raise HTTPException(status_code=500, detail="Word文書の読み込みに失敗しました。")

        # 要約対象のクエリ
        queries = {
            "current_situation": "業界の現状を説明してください。",
            "future_outlook": "業界の将来性や抱えている課題を説明してください。",
            "investment_advantages": f"業界の競合情報および{company_name}の差別化要因を教えてください。",
            "investment_disadvantages": f"{company_name}のExit先はどのような相手が有力でしょうか？",
            "swot_analysis": f"{company_name}のSWOT分析をお願いします。",
            "use_case": "業界のM&A事例について過去実績、将来の見込みを教えてください。",
            "value_up": f"{company_name}のバリューアップ施策をDX関連とその他に分けて教えてください。",
        }

        summaries = {}

        # 各クエリについて処理
        for key, query in queries.items():
            try:
                # 初回要約: ChatGPT
                chatgpt_response = await openai.ChatCompletion.acreate(
                    model="gpt-3.5-turbo",
                    messages=[
                        {"role": "user", "content": f"{text}\n\n質問: {query}\n500字以内で要約してください。"}
                    ],
                )
                chatgpt_summary = chatgpt_response.choices[0].message['content'].strip()
                logging.info(f"{key}のChatGPT初回要約結果: {chatgpt_summary}")
            except Exception as e:
                logging.error(f"ChatGPT初回要約エラー: {e}")
                chatgpt_summary = "ChatGPT初回要約エラーが発生しました。"

            # 結果を保存
            summaries[key] = {
                "chatgpt_summary": chatgpt_summary,
            }

        return summaries

    except Exception as e:
        logging.error(f"Blobストレージまたは要約処理中のエラー: {e}")
        raise HTTPException(status_code=500, detail="エラーが発生しました。再試行してください。")
    finally:
        if temp_file_path and os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
            except Exception as e:
                logging.warning(f"一時ファイルの削除に失敗しました: {e}")
                
                
async def unison_summary_logic(query_key: str, company_name: str, industry: str, chatgpt_summary: str) -> str:
    """
    Perplexityと統合要約を処理
    """
    try:
        # Perplexityで補足情報を取得
        perplexity_summary = await get_perplexity_summary(
            query_key=query_key,
            company_name=company_name,
            industry=industry
        )

        # 統合要約を生成
        combined_text = f"ChatGPTによる要約:\n{chatgpt_summary}\n\nPerplexityによる補足情報:\n{perplexity_summary}"
        
        # OpenAI APIを用いて統合要約を生成
        try:
            final_summary_response = await openai.ChatCompletion.acreate(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "user", "content": f"{combined_text}\n\n以上を基に、統合要約を500字以内でお願いします。"}
                ],
            )
            final_summary = final_summary_response.choices[0].message['content'].strip()
            logging.info(f"統合要約結果: {final_summary}")
        except Exception as e:
            logging.error(f"統合要約エラー: {e}")
            final_summary = "統合要約エラーが発生しました。"

        return final_summary
    except Exception as e:
        logging.error(f"統合要約エラー: {e}")
        return "統合要約エラーが発生しました。"
    
    
async def get_perplexity_summary(query_key: str, company_name: str, industry: str) -> str:
    """
    Perplexity APIを呼び出して補足情報を取得
    """
    try:
        headers = {
            "Authorization": f"Bearer {PERPLEXITY_API_KEY}",
            "Content-Type": "application/json"
        }
        payload = {
            "query_key": query_key,
            "company_name": company_name,
            "industry": industry
        }
        async with httpx.AsyncClient() as client:
            response = await client.post(PERPLEXITY_API_ENDPOINT, headers=headers, json=payload)
        
        if response.status_code != 200:
            logging.error(f"Perplexity APIエラー: {response.status_code} - {response.text}")
            return "Perplexityによる補足情報の取得に失敗しました。"
        
        data = response.json()
        # レスポンス形式に応じて適切に要約を取得
        perplexity_summary = data.get("summary", "補足情報が取得できませんでした。")
        return perplexity_summary
    except Exception as e:
        logging.error(f"Perplexity API呼び出し中のエラー: {e}")
        return "Perplexityによる補足情報の取得中にエラーが発生しました。"
        

async def regenerate_summary(
    category_name: str,
    company_name: str,
    query_key: str,
    perplexity_summary: Optional[str],
    custom_query: Optional[str] = None,
    include_perplexity: bool = False,
) -> dict:
    """
    特定の項目だけ再生成する。Perplexityでの補足情報を保持し、それらを含めて結果を返す。
    """
    temp_file_path = None  # 初期化
    try:
        # Blobサービスクライアントの初期化（非同期クライアント）
        blob_service_client = BlobServiceClient.from_connection_string(BLOB_CONNECTION_STRING)
        
        # category_name に .docx が含まれていないことを確認
        if category_name.lower().endswith(".docx"):
            blob_name = category_name
        else:
            blob_name = f"{category_name}.docx"  # 小分類に .docx を付け加えたファイル名
        
        # 一時ファイルの作成
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        temp_file_path = temp_file.name
        temp_file.close()
        text = ""

        # Blobクライアントを取得
        blob_client = blob_service_client.get_blob_client(container=BLOB_CONTAINER_NAME, blob=blob_name)
        logging.info(f"アクセスするBlob名: {blob_name}")

        # Blobストレージからファイルを非同期にダウンロード
        try:
            download_stream = await blob_client.download_blob()
            data = await download_stream.readall()
            with open(temp_file_path, "wb") as file:
                file.write(data)
        except Exception as e:
            logging.error(f"Blobダウンロードエラー: {e}")
            raise HTTPException(status_code=500, detail="Blobファイルのダウンロードに失敗しました。")

        # Word文書を読み込み、段落を結合
        try:
            doc = Document(temp_file_path)
            text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
        except Exception as e:
            logging.error(f"Word文書の読み込みエラー: {e}")
            raise HTTPException(status_code=500, detail="Word文書の読み込みに失敗しました。")

        # デフォルトクエリ
        default_queries = {
            "current_situation": "業界の現状を説明してください。",
            "future_outlook": f"業界の将来性や抱えている課題を説明してください。",
            "investment_advantages": f"業界の競合情報および{company_name}の差別化要因を教えてください。",
            "investment_disadvantages": f"{company_name}のExit先はどのような相手が有力でしょうか？",
            "swot_analysis": f"{company_name}のSWOT分析をお願いします。",
            "use_case": "業界のM&A事例について過去実績、将来の見込みを教えてください。",
            "value_up": f"{company_name}のバリューアップ施策をDX関連とその他に分けて教えてください。",
        }

        # クエリの取得
        query = custom_query if custom_query else default_queries.get(query_key)
        if not query:
            raise HTTPException(status_code=400, detail="指定されたクエリキーが無効です。")

        # 初回要約: ChatGPT
        try:
            chatgpt_response = await openai.ChatCompletion.acreate(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "user", "content": f"{text}\n\n質問: {query}\n500字以内で要約してください。"}
                ],
            )
            chatgpt_summary = chatgpt_response.choices[0].message['content'].strip()
            logging.info(f"ChatGPT初回要約結果: {chatgpt_summary}")
        except Exception as e:
            logging.error(f"ChatGPT初回要約エラー: {e}")
            chatgpt_summary = "ChatGPT初回要約エラーが発生しました。"

        # 統合要約の生成
        if include_perplexity and perplexity_summary:
            combined_text = f"ChatGPTによる要約:\n{chatgpt_summary}\n\nPerplexityによる補足情報:\n{perplexity_summary}"
            try:
                final_summary_response = await openai.ChatCompletion.acreate(
                    model="gpt-3.5-turbo",
                    messages=[
                        {"role": "user", "content": f"{combined_text}\n\n以上を基に、統合要約を500字以内でお願いします。"}
                    ],
                )
                final_summary = final_summary_response.choices[0].message['content'].strip()
                logging.info(f"統合要約結果: {final_summary}")
            except Exception as e:
                logging.error(f"統合要約エラー: {e}")
                final_summary = "統合要約エラーが発生しました。"
        else:
            # Perplexity要約がない場合、ChatGPTの要約をそのまま最終要約とする
            final_summary = chatgpt_summary

    finally:
        # 一時ファイルの削除
        if temp_file_path and os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
            except Exception as e:
                logging.warning(f"一時ファイルの削除に失敗しました: {e}")

    # 結果をまとめて返却
    return {
        "query_key": query_key,
        "chatgpt_summary": chatgpt_summary,
        "perplexity_summary": perplexity_summary if include_perplexity else None,
        "final_summary": final_summary,
    }