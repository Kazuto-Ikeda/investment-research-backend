from fastapi import Body, Query, BackgroundTasks
from fastapi.responses import FileResponse
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import os
import logging
import mistune
from mistune.plugins import plugin_table, plugin_strikethrough
from mistune import Markdown
from typing import Optional, Dict

# ログの設定
logging.basicConfig(level=logging.INFO)


# カスタムASTRendererの定義
class ASTRenderer(mistune.HTMLRenderer):
    def __init__(self, document: Document):
        super().__init__()
        self.document = document

    def heading(self, text, level):
        # Wordの見出しを追加
        if level == 1:
            paragraph = self.document.add_heading(text, level=1)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif level == 2:
            paragraph = self.document.add_heading(text, level=2)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            paragraph = self.document.add_heading(text, level=3)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return ''

    def paragraph(self, text):
        paragraph = self.document.add_paragraph()
        self.process_inline(paragraph, text)
        return ''

    def list(self, text, ordered, level, start=None):
        # Wordのリストを追加
        if ordered:
            style = 'List Number'
        else:
            style = 'List Bullet'
        paragraph = self.document.add_paragraph(text, style=style)
        return ''

    def table(self, header, body):
        # Wordのテーブルを追加
        num_cols = len(header)
        table = self.document.add_table(rows=1 + len(body), cols=num_cols)
        table.style = 'Table Grid'

        # ヘッダー行の設定
        hdr_cells = table.rows[0].cells
        for idx, header_cell in enumerate(header):
            hdr_cells[idx].text = header_cell
            for paragraph in hdr_cells[idx].paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(11)
                    run.font.name = 'Noto Sans JP'
                    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Noto Sans JP')

        # データ行の追加
        for row_idx, row in enumerate(body, start=1):
            row_cells = table.rows[row_idx].cells
            for col_idx, cell_text in enumerate(row):
                if col_idx < num_cols:
                    row_cells[col_idx].text = cell_text
                    for paragraph in row_cells[col_idx].paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(10)
                            run.font.name = 'Noto Sans JP'
                            run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Noto Sans JP')

        # セルの背景色設定（ヘッダーと特定行）
        for cell in hdr_cells:
            shading_elm_1 = cell._tc.get_or_add_tcPr().get_or_add_shd()
            shading_elm_1.fill = "EFEFEF"  # グレー系の背景色

        # 例: 偶数行に背景色を設定
        for row_idx, row in enumerate(body, start=1):
            if row_idx % 2 == 0:
                for cell in table.rows[row_idx].cells:
                    shading_elm_1 = cell._tc.get_or_add_tcPr().get_or_add_shd()
                    shading_elm_1.fill = "D0EEFB"  # 任意の色コード

    def strong(self, text):
        # 太字を適用
        run = self.document.add_run(text)
        run.bold = True
        return ''

    def emphasis(self, text):
        # 斜体を適用
        run = self.document.add_run(text)
        run.italic = True
        return ''

    def link(self, link, text=None, title=None):
        # リンクを適用
        run = self.document.add_run(text or link)
        run.font.color.rgb = RGBColor(0, 0, 255)
        run.underline = True
        # Wordではハイパーリンクを設定するのは少し複雑なので、簡易的にテキストに色と下線を追加
        return ''

    def inline_text(self, text):
        # 通常のテキスト
        self.document.add_run(text)
        return ''

    def process_inline(self, paragraph, text):
        # インライン要素の処理
        # mistune 2.x では BlockParser を使用
        tokens = mistune.block_parser.BlockParser().parse_inline(text)
        for token in tokens:
            if token['type'] == 'strong':
                run = paragraph.add_run(token['children'][0]['text'])
                run.bold = True
            elif token['type'] == 'emphasis':
                run = paragraph.add_run(token['children'][0]['text'])
                run.italic = True
            elif token['type'] == 'text':
                run = paragraph.add_run(token['text'])
            elif token['type'] == 'link':
                run = paragraph.add_run(token['children'][0]['text'])
                run.font.color.rgb = RGBColor(0, 0, 255)
                run.underline = True
            # 他のインライン要素も必要に応じて処理
        return ''

def delete_file(path: str):
    """指定されたファイルを削除"""
    try:
        os.remove(path)
        logging.info(f"ファイル {path} を削除しました。")
    except Exception as e:
        logging.error(f"ファイル削除エラー: {e}")


def add_markdown_content(document: Document, markdown_text: str):
    """
    Markdownテキストを解析し、Wordドキュメントにスタイルを適用して追加する関数
    """
    try:
        # カスタムASTRendererのインスタンス作成
        renderer = ASTRenderer(document)
        # Markdownパーサーの設定（mistune v2.x）
        markdown = Markdown(renderer=renderer, plugins=[plugin_table, plugin_strikethrough])
        # Markdownテキストを解析
        markdown(markdown_text)
    except Exception as e:
        logging.error(f"Markdown解析エラー: {e}")
        raise e


def generate_word_file(
    background_tasks: BackgroundTasks,
    summaries: dict = Body(..., description="要約データを含む辞書形式の入力"),
    valuation_data: Optional[dict] = Body(None, description="バリュエーションデータ（オプション）"),
    company_name: str = Query(..., description="会社名を指定"),
    file_name: Optional[str] = Query(None, description="生成するWordファイル名 (省略可能)")
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
        category_heading = document.add_heading(japanese_category, level=2)
        category_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = category_heading.runs[0]
        run.font.size = Pt(16)
        run.bold = True

        # 各セクションの内容を追加（番号を付与、level3=14pt）
        for idx, (section, content) in enumerate(sections.items(), start=1):
            # セクションの見出しを番号付きで日本語ラベルに変換
            japanese_section = reverse_key_mapping.get(section, section)
            section_heading = document.add_heading(f"{idx}. {japanese_section}", level=3)
            section_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = section_heading.runs[0]
            run.font.size = Pt(14)
            run.bold = True

            # セクションの内容をMarkdownとして解析し、スタイルを適用
            content = content or "内容がありません"
            add_markdown_content(document, content)

    # バリュエーションデータを表形式で追加
    if valuation_data:
        # バリュエーション見出しを追加（level2=16pt）
        valuation_heading = document.add_heading("バリュエーション", level=2)
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
                    run.font.name = 'Noto Sans JP'
                    run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Noto Sans JP')

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
                        run.font.name = 'Noto Sans JP'
                        run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Noto Sans JP')

    # フォントファミリーの統一
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Noto Sans JP'
            run.element.rPr.rFonts.set(qn('w:eastAsia'), 'Noto Sans JP')

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



