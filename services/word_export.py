from fastapi import Body, Query, BackgroundTasks
from fastapi.responses import FileResponse
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import re
import logging

def delete_file(path: str):
    """指定されたファイルを削除"""
    try:
        os.remove(path)
        logging.info(f"ファイル {path} を削除しました。")
    except Exception as e:
        logging.error(f"ファイル削除エラー: {e}")

def clean_text(text: str) -> str:
    """
    テキストから注釈や特定の記号を削除する関数
    """
    # 注釈（例: [1][3]）を削除
    text = re.sub(r'\[\d+\]', '', text)
    # すべての#を削除
    text = text.replace('#', '')
    # 追加で他のMarkdown記号（***, **, *, ___, __, _）を削除したい場合は以下を使用
    text = re.sub(r'\*\*\*|\*\*|\*|___|__|_', '', text)
    return text

def generate_word_file(
    background_tasks: BackgroundTasks,
    summaries: dict = Body(..., description="要約データを含む辞書形式の入力"),
    valuation_data: dict = Body(None, description="バリュエーションデータ（オプション）"),
    company_name: str = Query(..., description="会社名を指定"),
    file_name: str = Query(None, description="生成するWordファイル名 (省略可能)")
):
    """
    受け取った要約データおよびバリュエーションデータをWordドキュメントへ
    """
    # キーマッピングの定義
    reverse_key_mapping = {
        "current_situation": "現状",
        "future_outlook": "将来性と課題",
        "investment_advantages": "競合と差別化",
        "investment_disadvantages": "Exit先検討",
        "value_up": "バリューアップ施策",
        "use_case": "M&A事例",
        "swot_analysis": "SWOT分析",
    }

    # カテゴリのマッピング（必要に応じて）
    category_mapping = {
        "Perplexity": "Perplexity 分析",
        "ChatGPT": "ChatGPT+SPEEDA 分析",
    }

    # 動的ファイル名の設定
    file_name = file_name or f"{company_name}_summary_report.docx"

    # Wordドキュメントを作成
    document = Document()

    # タイトルを追加（level1=18pt）
    title = document.add_paragraph(f"{company_name} - 要約レポート")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.runs[0]
    run.font.size = Pt(18)
    run.bold = True

    # 要約内容をカテゴリごとにセクション化
    for main_category, sections in summaries.items():
        # カテゴリの見出しを日本語に変換（必要に応じて）
        japanese_category = category_mapping.get(main_category, main_category)
        
        # カテゴリの見出しを追加（level2=16pt）
        category_heading = document.add_paragraph(japanese_category)
        category_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = category_heading.runs[0]
        run.font.size = Pt(16)
        run.bold = True

        # 各セクションの内容を追加（番号を付与、level3=14pt）
        for idx, (section, content) in enumerate(sections.items(), start=1):
            # セクションの見出しを番号付きで日本語ラベルに変換
            japanese_section = reverse_key_mapping.get(section, section)
            section_heading = document.add_paragraph(f"{idx}. {japanese_section}")
            section_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = section_heading.runs[0]
            run.font.size = Pt(14)
            run.bold = True

            # セクションの内容をクリーンアップ
            clean_content = clean_text(content or "内容がありません")

            # セクションの内容を段落として追加（バレットポイントなし）
            paragraph = document.add_paragraph(clean_content, style='Normal')
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # バリュエーションデータを表形式で追加
    if valuation_data:
        # バリュエーション見出しを追加（level2=16pt）
        valuation_heading = document.add_paragraph("バリュエーション")
        valuation_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = valuation_heading.runs[0]
        run.font.size = Pt(16)
        run.bold = True

        # テーブルの作成
        table = document.add_table(rows=1, cols=3)
        table.style = 'Table Grid'

        # ヘッダー行の設定
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '項目'
        hdr_cells[1].text = '直近実績'
        hdr_cells[2].text = '進行期見込'

        # ヘッダーのフォーマット
        for cell in hdr_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(11)

        # バリュエーションデータの追加（番号なし）
        for key, value in valuation_data.items():
            row_cells = table.add_row().cells
            row_cells[0].text = key  # 番号付与を削除

            # Set the current and forecast values
            if isinstance(value, dict):
                row_cells[1].text = value.get('current', '不明')
                row_cells[2].text = value.get('forecast', '不明')
            else:
                row_cells[1].text = str(value)
                row_cells[2].text = '不明'

            # セルのフォーマット（バレットポイントなし）
            for cell in row_cells:
                for paragraph in cell.paragraphs:
                    paragraph.style = document.styles['Normal']  # バレットポイントを削除
                    for run in paragraph.runs:
                        run.font.size = Pt(10)

    # ファイル保存ディレクトリの設定
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
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



# from fastapi import Body, Query, BackgroundTasks
# from fastapi.responses import FileResponse
# from docx import Document
# from docx.shared import Pt
# from docx.enum.text import WD_ALIGN_PARAGRAPH
# import os
# import re
# import logging

# def delete_file(path: str):
#     """指定されたファイルを削除"""
#     try:
#         os.remove(path)
#         logging.info(f"ファイル {path} を削除しました。")
#     except Exception as e:
#         logging.error(f"ファイル削除エラー: {e}")

# def clean_text(text: str) -> str:
#     """
#     テキストからMarkdown記号や注釈を削除する関数
#     """
#     # 注釈（例: [1][3]）を削除
#     text = re.sub(r'\[\d+\]', '', text)
#     # すべての#を削除
#     text = text.replace('#', '')
#     return text

# def apply_markdown_formatting(paragraph, text: str):
#     """
#     テキスト内のMarkdown記号を解析し、対応するWordスタイルを適用する関数
#     Markdown記号は削除される
#     """
#     # 見出し1
#     if re.match(r'^# (.*)', text):
#         heading_text = re.sub(r'^# ', '', text)
#         paragraph.style = 'Heading 1'
#         paragraph.text = heading_text
#         return

#     # 見出し2
#     if re.match(r'^## (.*)', text):
#         heading_text = re.sub(r'^## ', '', text)
#         paragraph.style = 'Heading 2'
#         paragraph.text = heading_text
#         return

#     # 見出し3
#     if re.match(r'^### (.*)', text):
#         heading_text = re.sub(r'^### ', '', text)
#         paragraph.style = 'Heading 3'
#         paragraph.text = heading_text
#         return

#     # リスト項目
#     list_match = re.match(r'^[-*+] (.*)', text)
#     if list_match:
#         list_text = list_match.group(1)
#         paragraph.style = 'List Bullet'
#         paragraph.text = list_text
#         return

#     # 通常の段落
#     # 太字と斜体の処理
#     bold_italic_patterns = [
#         (r'\*\*\*(.*?)\*\*\*', 'bold_italic'),  # ***text***
#         (r'\*\*(.*?)\*\*', 'bold'),             # **text**
#         (r'\*(.*?)\*', 'italic'),               # *text*
#         (r'___(.*?)___', 'bold_italic'),        # ___text___
#         (r'__(.*?)__', 'bold'),                 # __text__
#         (r'_(.*?)_', 'italic'),                 # _text_
#     ]

#     # 全ての既存のランをクリア
#     for run in paragraph.runs:
#         run.text = ''

#     # 太字と斜体の適用
#     for pattern, style in bold_italic_patterns:
#         matches = re.findall(pattern, text)
#         for match in matches:
#             if style == 'bold_italic':
#                 run = paragraph.add_run(match)
#                 run.bold = True
#                 run.italic = True
#             elif style == 'bold':
#                 run = paragraph.add_run(match)
#                 run.bold = True
#             elif style == 'italic':
#                 run = paragraph.add_run(match)
#                 run.italic = True
#             # テキストからMarkdown記号を削除
#             text = re.sub(rf'\*\*\*{re.escape(match)}\*\*\*', match, text)
#             text = re.sub(rf'\*\*{re.escape(match)}\*\*', match, text)
#             text = re.sub(rf'\*{re.escape(match)}\*', match, text)
#             text = re.sub(rf'___{re.escape(match)}___', match, text)
#             text = re.sub(rf'__{re.escape(match)}__', match, text)
#             text = re.sub(rf'_{re.escape(match)}_', match, text)

#     # 残りのテキストを追加
#     paragraph.add_run(text)

# def generate_word_file(
#     background_tasks: BackgroundTasks,
#     summaries: dict = Body(..., description="要約データを含む辞書形式の入力"),
#     valuation_data: dict = Body(None, description="バリュエーションデータ（オプション）"),
#     company_name: str = Query(..., description="会社名を指定"),
#     file_name: str = Query(None, description="生成するWordファイル名 (省略可能)")
# ):
#     """
#     受け取った要約データおよびバリュエーションデータをWordドキュメントへ
#     """
#     # キーマッピングの定義
#     reverse_key_mapping = {
#         "current_situation": "現状",
#         "future_outlook": "将来性と課題",
#         "investment_advantages": "競合と差別化",
#         "investment_disadvantages": "Exit先検討",
#         "value_up": "バリューアップ施策",
#         "use_case": "M&A事例",
#         "swot_analysis": "SWOT分析",
#     }

#     # カテゴリのマッピング（必要に応じて）
#     category_mapping = {
#         "Perplexity": "Perplexity 分析",
#         "ChatGPT": "ChatGPT+SPEEDA 分析",
#     }

#     # 動的ファイル名の設定
#     file_name = file_name or f"{company_name}_summary_report.docx"

#     # Wordドキュメントを作成
#     document = Document()

#     # タイトルを追加（level1=18pt）
#     title = document.add_paragraph(f"{company_name} - 要約レポート")
#     title.alignment = WD_ALIGN_PARAGRAPH.CENTER
#     run = title.runs[0]
#     run.font.size = Pt(18)
#     run.bold = True

#     # 要約内容をカテゴリごとにセクション化
#     for main_category, sections in summaries.items():
#         # カテゴリの見出しを日本語に変換（必要に応じて）
#         japanese_category = category_mapping.get(main_category, main_category)
        
#         # カテゴリの見出しを追加（level2=16pt）
#         category_heading = document.add_paragraph(japanese_category)
#         category_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
#         run = category_heading.runs[0]
#         run.font.size = Pt(16)
#         run.bold = True

#         # 各セクションの内容を追加（番号を付与、level3=14pt）
#         for idx, (section, content) in enumerate(sections.items(), start=1):
#             # セクションの見出しを番号付きで日本語ラベルに変換
#             japanese_section = reverse_key_mapping.get(section, section)
#             section_heading = document.add_paragraph(f"{idx}. {japanese_section}")
#             section_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
#             run = section_heading.runs[0]
#             run.font.size = Pt(14)
#             run.bold = True

#             # セクションの内容をクリーンアップ
#             clean_content = clean_text(content or "内容がありません")

#             # セクションの内容を段落として追加（バレットポイントなし）
#             paragraph = document.add_paragraph(clean_content, style='Normal')
#             paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
#             apply_markdown_formatting(paragraph, clean_content)

#     # バリュエーションデータを表形式で追加
#     if valuation_data:
#         # バリュエーション見出しを追加（level2=16pt）
#         valuation_heading = document.add_paragraph("バリュエーション")
#         valuation_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
#         run = valuation_heading.runs[0]
#         run.font.size = Pt(16)
#         run.bold = True

#         # テーブルの作成
#         table = document.add_table(rows=1, cols=3)
#         table.style = 'Table Grid'

#         # ヘッダー行の設定
#         hdr_cells = table.rows[0].cells
#         hdr_cells[0].text = '項目'
#         hdr_cells[1].text = '直近実績'
#         hdr_cells[2].text = '進行期見込'

#         # ヘッダーのフォーマット
#         for cell in hdr_cells:
#             for paragraph in cell.paragraphs:
#                 for run in paragraph.runs:
#                     run.font.bold = True
#                     run.font.size = Pt(11)

#         # バリュエーションデータの追加（番号なし）
#         for key, value in valuation_data.items():
#             row_cells = table.add_row().cells
#             row_cells[0].text = key  # 番号付与を削除

#             # Set the current and forecast values
#             if isinstance(value, dict):
#                 row_cells[1].text = value.get('current', '不明')
#                 row_cells[2].text = value.get('forecast', '不明')
#             else:
#                 row_cells[1].text = str(value)
#                 row_cells[2].text = '不明'

#             # セルのフォーマット（バレットポイントなし）
#             for cell in row_cells:
#                 for paragraph in cell.paragraphs:
#                     paragraph.style = document.styles['Normal']  # バレットポイントを削除
#                     for run in paragraph.runs:
#                         run.font.size = Pt(10)

#     # ファイル保存ディレクトリの設定
#     output_dir = "output"
#     os.makedirs(output_dir, exist_ok=True)
#     output_path = os.path.join(output_dir, file_name)
#     document.save(output_path)

#     # ダウンロード後にファイルを削除
#     background_tasks.add_task(delete_file, output_path)

#     # 生成されたWordファイルを返却
#     return FileResponse(
#         output_path,
#         filename=file_name,
#         media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
#     )
    
    

