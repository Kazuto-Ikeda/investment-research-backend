from fastapi import HTTPException, Query
from fastapi.responses import JSONResponse
from docx import Document
from dotenv import load_dotenv
from azure.storage.blob import BlobServiceClient  # 非同期クライアントから同期クライアントに変更
from pydantic import BaseModel
from typing import Optional, Dict
from models.model import RegenerateRequest
from openai import OpenAI
import os
import logging
import httpx
import requests
import tempfile
import unicodedata

# ロギング設
logging.basicConfig(level=logging.INFO)



# 環境変数の読み込み
load_dotenv()


# APIキーの取得
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
PERPLEXITY_API_KEY = os.getenv("PERPLEXITY_API_KEY")
PERPLEXITY_API_ENDPOINT = os.getenv("PERPLEXITY_API_ENDPOINT")
GPT_MODEL = os.getenv("GPT_MODEL")

# Azure Blob Storage設定
BLOB_CONNECTION_STRING = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
BLOB_CONTAINER_NAME = os.getenv("BLOB_CONTAINER_NAME")


client = OpenAI(api_key = OPENAI_API_KEY)

# Unicode正規化関数（NFD形式で正規化）
def normalize_text(text: str) -> str:
    """文字列をNFD形式で正規化"""
    normalized = unicodedata.normalize('NFD', text)
    # logging.debug(f"Original text: '{text}' | Normalized text: '{normalized}'")
    return normalized


def summary_from_speeda(category: str, prompt: str) -> str:
    """
    Blobストレージからファイルをダウンロードし、一時ファイルとして保存。
    小分類名に基づいて .docx ファイルを検索します。
    要約を生成して返します。
    """
    temp_file_path = None  # 初期化
    try:
        # カテゴリー名の正規化（NFD形式）
        normalized_category = normalize_text(category)
        logging.info(f"Normalized category name: '{normalized_category}'")
        
        # categoryにすでに.docxが含まれているか確認
        if normalized_category.lower().endswith(".docx"):
            blob_name = normalized_category
        else:
            blob_name = f"{normalized_category}.docx"

        logging.info(f"Constructed blob name: '{blob_name}'")
        
        # Blobサービスクライアントの初期化 (同期)
        blob_service_client = BlobServiceClient.from_connection_string(BLOB_CONNECTION_STRING)
        blob_client = blob_service_client.get_blob_client(container=BLOB_CONTAINER_NAME, blob=blob_name)
        logging.info(f"アクセスするBlob名: '{blob_name}'")

        # Blobの存在確認
        blob_exists = blob_client.exists()
        logging.debug(f"Blob '{blob_name}' の存在: {blob_exists}")
        if not blob_exists:
            logging.error(f"指定されたBlob '{blob_name}' はコンテナ '{BLOB_CONTAINER_NAME}' に存在しません。")
            raise HTTPException(status_code=404, detail="指定されたBlobファイルが存在しません。")

        # 一時ファイルの作成
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_file:
            temp_file_path = temp_file.name

            # Blobストレージからファイルをダウンロード
            try:
                data = blob_client.download_blob().readall()
                with open(temp_file_path, "wb") as file:
                    file.write(data)
                logging.info(f"Blob '{blob_name}' を一時ファイル '{temp_file_path}' にダウンロードしました。")
            except Exception as e:
                logging.error(f"Blobダウンロードエラー: {e}")
                raise HTTPException(status_code=500, detail="Blobファイルのダウンロードに失敗しました。")

        # Word文書を読み込み、段落を結合
        try:
            doc = Document(temp_file_path)
            text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
            logging.info(f"Word文書 '{blob_name}' の内容を読み込みました。")
        except Exception as e:
            logging.error(f"Word文書の読み込みエラー: {e}")
            raise HTTPException(status_code=500, detail="Word文書の読み込みに失敗しました。")

        try:
            # 初回要約: ChatGPT (同期)
            chatgpt_response = client.chat.completions.create(
                model=GPT_MODEL,  # モデル名を最新のものに変更
                messages=[
                    {"role": "system", "content": "あなたは優秀な投資家であり、市場調査の専門家です。"},
                    {"role": "user", "content": f"{text}\n\n質問: {prompt}\n500字以内で要約してください。"}
                ],
            )
            chatgpt_summary = chatgpt_response.choices[0].message.content
            logging.info(f"{prompt}のChatGPT要約結果: {chatgpt_summary}")
        except HTTPException as e:
            logging.error(f"ChatGPT初回要約エラー: {e}")
            chatgpt_summary = "ChatGPT初回要約エラーが発生しました。"

        return chatgpt_summary

    except HTTPException as he:
        # HTTPExceptionをそのまま投げる
        raise he
    except Exception as e:
        logging.error(f"Blobストレージまたは要約処理中のエラー: {e}")
        raise HTTPException(status_code=500, detail="エラーが発生しました。再試行してください。")
    finally:
        if temp_file_path and os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
                logging.info(f"一時ファイル '{temp_file_path}' を削除しました。")
            except Exception as e:
                logging.warning(f"一時ファイルの削除に失敗しました: {e}")
                
                

def perplexity_search(prompt: str) -> str:
    """
    Perplexity APIを呼び出して補足情報を取得
    """
    try:
        headers = {
            "Authorization": f"Bearer {PERPLEXITY_API_KEY}",
            "Content-Type": "application/json"
        }
        payload = {
            "model": "llama-3.1-sonar-small-128k-online",
            "messages": [
                {
                    "role": "system",
                    "content": "あなたは優秀な投資家です。"
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ], 
            "temperature": 0
            }

        response = requests.request("POST", PERPLEXITY_API_ENDPOINT, json=payload, headers=headers)
        print(response)

        if response.status_code != 200:
            logging.error(f"Perplexity APIエラー: {response.status_code} - {response.text}")
            return "Perplexityによる補足情報の取得に失敗しました。"
        
        # レスポンス形式に応じて適切に要約を取得
        data = response.json()
        perplexity_summary = data["choices"][0]["message"]["content"]  # 修正箇所
        return perplexity_summary
    except Exception as e:
        logging.error(f"Perplexity API呼び出し中のエラー: {e}")
        return "Perplexityによる補足情報の取得中にエラーが発生しました。"