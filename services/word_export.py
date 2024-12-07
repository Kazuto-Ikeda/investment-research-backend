from fastapi import FastAPI, HTTPException, Query, Body
from fastapi.responses import FileResponse
from docx import Document
from fastapi.background import BackgroundTasks
import os


def delete_file(path: str):
    """指定されたファイルを削除"""
    try:
        os.remove(path)
        print(f"ファイル {path} を削除しました。")
    except Exception as e:
        print(f"ファイル削除エラー: {e}")



def generate_word_file(
    background_tasks: BackgroundTasks,
    summaries: dict = Body(..., description="要約データを含む辞書形式の入力"),
    company_name: str = Query(..., description="会社名を指定"),
    file_name: str = Query(None, description="生成するWordファイル名 (省略可能)")
):
    """
    受け取った要約データをWordドキュメントに保存し、一時的にダウンロード可能にする
    """
    # 動的ファイル名の設定
    file_name = file_name or f"{company_name}_summary_report.docx"

    # Wordドキュメントを作成
    document = Document()

    # タイトルを追加
    document.add_heading(f"{company_name} - 要約レポート", level=1)

    # 要約内容をセクションごとに記載
    for section, content in summaries.items():
        # セクションヘッダー
        document.add_heading(section.replace("_", " ").capitalize(), level=2)
        # セクション内容
        document.add_paragraph(content)

    # ファイル保存ディレクトリの設定
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)  # 出力ディレクトリを作成
    output_path = os.path.join(output_dir, file_name)
    document.save(output_path)

    # ダウンロード後にファイルを削除
    background_tasks.add_task(delete_file, output_path)

    # 生成されたWordファイルを返却
    return FileResponse(
        output_path,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

