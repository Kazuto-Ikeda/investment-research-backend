from fastapi import HTTPException, Query
from fastapi.responses import JSONResponse
from docx import Document
from dotenv import load_dotenv
from azure.storage.blob import BlobServiceClient
from pydantic import BaseModel
from typing import List, Optional
from models.model import RegenerateRequest
import pymysql
import os
import logging
from openai import OpenAI
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
        
        


def download_blob_to_temp_file(category: str, company_name: str,) -> str:
    """
    Blobストレージからファイルをダウンロードし、一時ファイルとして保存。
    小分類名に基づいて .docx ファイルを検索します。
    """
    blob_service_client = BlobServiceClient.from_connection_string(BLOB_CONNECTION_STRING)
    blob_name = f"{category}.docx"  # 小分類に .docx を付け加えたファイル名
    temp_file_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    text = ""
    
    try:
        # Blobクライアントを取得
        blob_client = blob_service_client.get_blob_client(container=BLOB_CONTAINER_NAME, blob=blob_name)
        logging.info(f"アクセスするBlob名: {blob_name}")

        # Blobストレージからファイルをダウンロード
        with open(temp_file_path, "wb") as file:
            download_stream = blob_client.download_blob()
            file.write(download_stream.readall())

        # Word文書を読み込み、段落を結合
        doc = Document(temp_file_path)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"

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
                # 1工程目: 初回ChatGPT要約
                chatgpt_response = client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[
                        {"role": "user", "content": f"{text}\n\n質問: {query}\n300字以内で要約してください。"}
                    ]
                )
                chatgpt_summary = chatgpt_response.choices[0].message.content.strip()
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


def unison_summary_logic(query_key: str, company_name: str, industry: str, chatgpt_summary: str) -> str:
    """
    Perplexityと統合要約を処理
    """
    try:
        # Perplexityで補足情報を取得
        perplexity_summary = f"Perplexityで取得した補足情報: {query_key}, {company_name}, {industry}"

        # 統合要約を生成
        combined_text = f"ChatGPTによる要約:\n{chatgpt_summary}\n\nPerplexityによる補足情報:\n{perplexity_summary}"
        final_summary_response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": f"{combined_text}\n\n以上を基に、統合要約を300字以内でお願いします。"}
            ]
        )
        final_summary = final_summary_response.choices[0].message.content.strip()
        return final_summary
    except Exception as e:
        logging.error(f"統合要約エラー: {e}")
        return "統合要約エラーが発生しました。"


def regenerate_summary(
    category_name: str,
    company_name: str,
    query_key: str,
    perplexity_summary: str,
    custom_query: Optional[str] = None,
    include_perplexity: bool = False,
):
    """
    特定の項目だけ再生成する。Perplexityでの補足情報を保持し、それらを含めて結果を返す。
    """
    blob_service_client = BlobServiceClient.from_connection_string(BLOB_CONNECTION_STRING)
    blob_name = f"{category_name}.docx"  # 小分類に .docx を付け加えたファイル名
    temp_file_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    text = ""

    # Blobクライアントを取得
    blob_client = blob_service_client.get_blob_client(container=BLOB_CONTAINER_NAME, blob=blob_name)
    logging.info(f"アクセスするBlob名: {blob_name}")

    # Blobストレージからファイルをダウンロード
    with open(temp_file_path, "wb") as file:
        download_stream = blob_client.download_blob()
        file.write(download_stream.readall())

    # Word文書を読み込み、段落を結合
    doc = Document(temp_file_path)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text + "\n"


    # デフォルトクエリ
    default_queries = {
        "current_situation": "業界の現状を説明してください。",
        "future_outlook": "業界の将来性や抱えている課題を説明してください。",
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

    try:
        # 初回要約: ChatGPT
        chatgpt_response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": f"{text}\n\n質問: {query}\n300字以内で要約してください。"}
            ],
        )
        chatgpt_summary = chatgpt_response.choices[0].message.content.strip()
        logging.info(f"ChatGPT初回要約結果: {chatgpt_summary}")
    except Exception as e:
        logging.error(f"ChatGPT初回要約エラー: {e}")
        chatgpt_summary = "ChatGPT初回要約エラーが発生しました。"

    # 統合要約
    final_summary = chatgpt_summary
    combined_text = f"ChatGPTによる要約:\n{chatgpt_summary}\n\nPerplexityによる補足情報:\n{perplexity_summary}"
    try:
        final_summary_response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": f"{combined_text}\n\n以上を基に、統合要約を300字以内でお願いします。"}
            ],
        )
        final_summary = final_summary_response.choices[0].message.content.strip()
        logging.info(f"統合要約結果: {final_summary}")
    except Exception as e:
        logging.error(f"統合要約エラー: {e}")
        final_summary = "統合要約エラーが発生しました。"

    # 結果をまとめて返却
    return JSONResponse(
        content={
            "query_key": query_key,
            "chatgpt_summary": chatgpt_summary,
            "perplexity_summary": perplexity_summary,
            "final_summary": final_summary,
        }
    )

