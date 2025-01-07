from fastapi import Body, Query, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import Dict, List, Optional
import os
import re
import logging
import mistune
from mistune import Markdown
from mistune import create_markdown
from mistune.plugins import plugin_table



########################################
# カスタムレンダラー: DocxRenderer
########################################
class DocxRenderer(mistune.AstRenderer):
    """
    Mistune 2.x の AstRenderer を継承し、MarkdownのASTをpython-docxでWordに変換する。

    【本サンプルで対応する要素】
      - heading (h1～h6)
      - paragraph (段落)
      - list / list_item (箇条書き, 番号リスト, 入れ子)
      - table (表)
      - strong, emphasis (太字, 斜体)
      - blockquote, thematic_break (引用, 区切り線) などの追加実装例

    段落内では複数の Run を生成し、強調(太字/斜体)を適宜反映。
    リストは入れ子にも対応。テーブルは表のヘッダとセルを単純に埋める。
    """

    def __init__(self, document: Document):
        super().__init__()
        self.document = document
        self.current_paragraph = None  # 段落オブジェクト（処理中の段落）

    def render(self, tokens, state):
        """
        AST全体を走査し、トークンの種類に応じた _render_xxx() を呼び出す。
        """
        for token in tokens:
            node_type = token['type']
            if node_type == 'heading':
                self._render_heading(token)
            elif node_type == 'paragraph':
                self._render_paragraph(token)
            elif node_type == 'list':
                self._render_list(token, level=0)
            elif node_type == 'blockquote':
                self._render_blockquote(token)
            elif node_type == 'thematic_break':
                self._render_thematic_break(token)
            elif node_type == 'table':
                self._render_table(token)
            else:
                logging.debug(f"Skipping unknown node type: {node_type}")
        return ''  # Word文書に直接書き込むので文字列は返さない

    ############################
    # heading (見出し)
    ############################
    def _render_heading(self, token):
        level = token['level']
        children = token.get('children', [])

        # 新しい段落を作成
        self.current_paragraph = self.document.add_paragraph()
        # レベルに応じたフォントサイズ, alignment
        if level == 1:
            fsize = Pt(18)
            self.current_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif level == 2:
            fsize = Pt(16)
            self.current_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            fsize = Pt(14)
            self.current_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # heading は基本太字
        for child in children:
            text_ = self._extract_text(child)
            run_ = self.current_paragraph.add_run(text_)
            run_.font.size = fsize
            run_.bold = True

    ############################
    # paragraph (段落)
    ############################
    def _render_paragraph(self, token):
        self.current_paragraph = self.document.add_paragraph()
        children = token.get('children', [])
        self._render_inline_children(children)

    def _render_inline_children(self, children):
        """
        段落内のインライン要素 (text, strong, emphasis, など) をRunに変換
        """
        for node in children:
            ntype = node['type']
            if ntype == 'text':
                run_ = self.current_paragraph.add_run(node['text'])
            elif ntype == 'strong':
                txt = self._extract_text(node)
                run_ = self.current_paragraph.add_run(txt)
                run_.bold = True
            elif ntype == 'emphasis':
                txt = self._extract_text(node)
                run_ = self.current_paragraph.add_run(txt)
                run_.italic = True
            else:
                # link, codespan, image等は必要に応じて追加
                txt = self._extract_text(node)
                run_ = self.current_paragraph.add_run(txt)

    ############################
    # list + list_item (入れ子対応)
    ############################
    def _render_list(self, token, level=0):
        ordered = token.get('ordered', False)
        for child in token.get('children', []):
            if child['type'] == 'list_item':
                self._render_list_item(child, ordered=ordered, level=level)
            elif child['type'] == 'list':
                # 入れ子リスト (list の中に list)
                self._render_list(child, level=level+1)

    def _render_list_item(self, token, ordered, level):
        """
        list_itemには子要素が paragraph であることが多いが、単なるtextの場合もあり得る
        """
        style = 'List Number' if ordered else 'List Bullet'
        self.current_paragraph = self.document.add_paragraph(style=style)
        # レベルに応じてインデント (0.5cm * levelなどは任意調整)
        self.current_paragraph.paragraph_format.left_indent = Cm(0.5 * level)

        children = token.get('children', [])
        # list_item 内に複数 paragraph が入る場合がある。
        # シンプルに最初の paragraphだけ同じ段落に書く例
        for child in children:
            ctype = child['type']
            if ctype == 'paragraph':
                inline_nodes = child.get('children', [])
                self._render_inline_children(inline_nodes)
            elif ctype == 'list':
                # さらに入れ子のリスト
                self._render_list(child, level=level+1)
            else:
                # textなど
                txt = self._extract_text(child)
                self.current_paragraph.add_run(txt)

    ############################
    # blockquote (引用)
    ############################
    def _render_blockquote(self, token):
        """
        引用ブロック。ここではスタイル 'Intense Quote' を適用。
        """
        self.current_paragraph = self.document.add_paragraph(style='Intense Quote')
        children = token.get('children', [])
        # blockquote内は複数paragraphの場合もあるので要注意
        for child in children:
            if child['type'] == 'paragraph':
                inline_nodes = child.get('children', [])
                run_text = ''.join(self._extract_text(n) for n in inline_nodes)
                self.current_paragraph.add_run(run_text)
            else:
                # さらにlistやheadingがある場合は要拡張
                pass

    ############################
    # thematic_break (水平線)
    ############################
    def _render_thematic_break(self, token):
        """
        水平線。簡単に区切り行を挿入する例。
        """
        hr_para = self.document.add_paragraph()
        hr_para.add_run('----------').bold = True

    ############################
    # table (表)
    ############################
    def _render_table(self, token):
        """
        token は { 'type':'table', 'header': [...], 'align': [...], 'cells': [ [...], [...]] }
        Mistuneのtableプラグイン利用
        """
        header = token.get('header', [])
        cells = token.get('cells', [])
        if not header or not cells:
            return

        col_count = len(header)
        table = self.document.add_table(rows=1, cols=col_count)
        table.style = 'Table Grid'

        # ヘッダー
        hdr_cells = table.rows[0].cells
        for i, cell_ast in enumerate(header):
            text_ = self._extract_text(cell_ast)
            hdr_cells[i].text = text_
            # 太字
            for p in hdr_cells[i].paragraphs:
                for r in p.runs:
                    r.bold = True

        # 本文
        for row_data in cells:
            row_cells = table.add_row().cells
            for col_idx, cell_ast in enumerate(row_data):
                text_ = self._extract_text(cell_ast)
                row_cells[col_idx].text = text_

    ############################
    # ユーティリティ: インラインテキスト抽出
    ############################
    def _extract_text(self, node):
        """
        子要素を再帰的に走査してtextを連結。必要に応じてstrong/emphasis/linkの処理拡張
        """
        if 'text' in node:
            return node['text']
        elif 'children' in node:
            return ''.join(self._extract_text(child) for child in node['children'])
        return ''


########################################
# 後処理で一時ファイルを削除する関数
########################################
def delete_file(path: str):
    try:
        os.remove(path)
        logging.info(f"ファイル {path} を削除しました。")
    except Exception as e:
        logging.error(f"ファイル削除エラー: {e}")

########################################
# 2) Markdown → Word + 追加のバリュエーション表
########################################
def generate_word_file(
    background_tasks: BackgroundTasks,
    summaries: dict,
    valuation_data: Optional[dict],
    company_name: str,
    file_name: Optional[str] = None,
) -> FileResponse:
    """
    1. MarkdownをDocxRendererでWordに変換
    2. バリュエーションデータを表出力
    3. .docxファイルを返却
    """
    # 見出しキー変換
    reverse_key_mapping = {
        "current_situation": "現状",
        "future_outlook": "将来性と課題",
        "investment_advantages": "競合と差別化",
        "investment_disadvantages": "Exit先検討",
        "value_up": "バリューアップ施策",
        "use_case": "M&A事例",
        "swot_analysis": "SWOT分析",
    }

    category_mapping = {
        "Perplexity": "Perplexity 分析",
        "ChatGPT": "ChatGPT+SPEEDA 分析",
    }

    # ファイル名
    file_name = file_name or f"{company_name}_summary_report.docx"

    # Word文書
    document = Document()

    # タイトル作成
    title_para = document.add_paragraph(f"{company_name} - 要約レポート")
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.runs[0]
    title_run.font.size = Pt(18)
    title_run.bold = True

    # Mistuneレンダラー
    docx_renderer = DocxRenderer(document)
    md_parser = create_markdown(renderer=docx_renderer, plugins=[plugin_table])

    # 1) summaries (カテゴリ→セクション) を処理
    for main_category, sections in summaries.items():
        # カテゴリ見出し
        cat_jp = category_mapping.get(main_category, main_category)
        cat_heading = document.add_paragraph(cat_jp)
        cat_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        cat_run = cat_heading.runs[0]
        cat_run.font.size = Pt(16)
        cat_run.bold = True

        # セクションを順次
        for idx, (sec_key, sec_value) in enumerate(sections.items(), start=1):
            sec_jp = reverse_key_mapping.get(sec_key, sec_key)
            sec_heading = document.add_paragraph(f"{idx}. {sec_jp}")
            sec_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
            heading_run = sec_heading.runs[0]
            heading_run.font.size = Pt(14)
            heading_run.bold = True

            # Markdown解析 → Wordに書き込み
            markdown_text = sec_value or "内容がありません"
            md_parser(markdown_text)

    # 2) バリュエーション表
    if valuation_data:
        val_heading = document.add_paragraph("バリュエーション")
        val_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        val_run = val_heading.runs[0]
        val_run.font.size = Pt(16)
        val_run.bold = True

        val_table = document.add_table(rows=1, cols=3)
        val_table.style = 'Table Grid'

        hdr_cells = val_table.rows[0].cells
        hdr_cells[0].text = '項目'
        hdr_cells[1].text = '直近実績'
        hdr_cells[2].text = '進行期見込'

        # ヘッダー書式
        for cell_ in hdr_cells:
            for p_ in cell_.paragraphs:
                for r_ in p_.runs:
                    r_.font.bold = True
                    r_.font.size = Pt(11)

        # バリュエーションデータを行追加
        for key, val_obj in valuation_data.items():
            row_ = val_table.add_row().cells
            row_[0].text = key
            if isinstance(val_obj, dict):
                row_[1].text = val_obj.get("current", "不明")
                row_[2].text = val_obj.get("forecast", "不明")
            else:
                row_[1].text = str(val_obj)
                row_[2].text = "不明"

            # セル書式
            for c_ in row_:
                for p_ in c_.paragraphs:
                    p_.style = document.styles['Normal']
                    for r_ in p_.runs:
                        r_.font.size = Pt(10)

    # 3) ファイル保存 → FileResponse
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, file_name)
    document.save(output_path)

    background_tasks.add_task(delete_file, output_path)
    return FileResponse(
        output_path,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# def generate_word_file(
#     background_tasks: BackgroundTasks,
#     summaries: dict = Body(..., description="要約データを含む辞書形式の入力"),
#     valuation_data: dict = Body(None, description="バリュエーションデータ（オプション）"),
#     company_name: str = Query(..., description="会社名を指定"),
#     file_name: str = Query(None, description="生成するWordファイル名 (省略可能)")
# ) -> FileResponse:
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

#     # 1) summaries をカテゴリ (Perplexity, ChatGPT...) ごとに処理
#     for main_category, sections in summaries.items():
#         # カテゴリ見出し
#         jap_category = category_mapping.get(main_category, main_category)
#         cat_heading = document.add_paragraph(jap_category)
#         cat_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
#         cat_run = cat_heading.runs[0]
#         cat_run.font.size = Pt(16)
#         cat_run.bold = True

#         # カテゴリ配下の各セクション
#         for idx, (sec_key, sec_content) in enumerate(sections.items(), start=1):
#             # セクション見出し
#             jap_section = reverse_key_mapping.get(sec_key, sec_key)
#             sec_heading = document.add_paragraph(f"{idx}. {jap_section}")
#             sec_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
#             heading_run = sec_heading.runs[0]
#             heading_run.font.size = Pt(14)
#             heading_run.bold = True

#             # セクション本文（Markdown からテーブルを抽出してみる）
#             raw_text = sec_content or "内容がありません"
#             tables = parse_markdown_table(raw_text)

#             if tables:
#                 # テーブルが見つかった場合は表として追加
#                 for table_data in tables:
#                     add_table_to_document(document, table_data)
#                     document.add_paragraph()  # 表の後に空白段落を追加

#                 # テーブル部分を除いた残りのテキストが欲しい場合は、
#                 # 追加のロジックで "表以外" を抜き出す必要がある。
#                 # シンプルに全体を段落に入れたい時は clean_text して段落にする:
#                 #  paragraph = document.add_paragraph(clean_text(raw_text), style='Normal')
#             else:
#                 # テーブルが無ければ普通の文章として処理
#                 paragraph = document.add_paragraph(clean_text(raw_text), style='Normal')
#                 paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


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