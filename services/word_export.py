from fastapi import Body, Query, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import Dict, List, Optional
import os
import re
import logging
# import mistune
# from mistune import Markdown
# from mistune import create_markdown
# from mistune.plugins import plugin_table



# class DocxRenderer(mistune.AstRenderer):
#     """
#     Mistune 2.x の AstRenderer を継承し、Markdown の AST を Python-docx で Word に変換。
    
#     主な対応要素:
#       - 見出し(heading) -> Word段落+run（太字/フォントサイズ）
#       - 段落(paragraph) -> Word段落
#       - 箇条書き(list, list_item) -> List Bullet/Number スタイル, 入れ子リスト対応
#       - テーブル(table) -> plugin_table が生成したASTをそのまま受け取り、Word表に変換
#       - 強調(strong, emphasis) -> Runの bold/italic
#       - blockquote(引用), thematic_break(水平線) -> サンプル実装
    
#     段落の途中で太字や斜体にしたい場合、Runを分割して追加しています。
#     """

#     def __init__(self, document: Document):
#         super().__init__()
#         self.document = document
#         self.current_paragraph = None  # 処理中の段落オブジェクト

#     def render(self, tokens, state):
#         """
#         AST全体を走査し、トークン種類に応じて _render_xxx() メソッドを呼び出す。
#         """
#         for token in tokens:
#             node_type = token['type']
#             if node_type == 'heading':
#                 self._render_heading(token)
#             elif node_type == 'paragraph':
#                 self._render_paragraph(token)
#             elif node_type == 'list':
#                 self._render_list(token, level=0)
#             elif node_type == 'blockquote':
#                 self._render_blockquote(token)
#             elif node_type == 'thematic_break':
#                 self._render_thematic_break(token)
#             elif node_type == 'table':
#                 self._render_table(token)
#             else:
#                 logging.debug(f"[DocxRenderer] Skip unknown node: {node_type}")
#         return ''  # 文字列は返さず、Word文書に直接書き込む

#     #######################################################
#     # heading (見出し)
#     #######################################################
#     def _render_heading(self, token):
#         """
#         token例:
#           {
#             'type': 'heading',
#             'level': 1..6,
#             'children': [...]
#           }
#         """
#         level = token['level']
#         children = token.get('children', [])

#         self.current_paragraph = self.document.add_paragraph()

#         # heading レベル別にフォントサイズや整列を切り替え
#         if level == 1:
#             fsize = Pt(18)
#             self.current_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
#         elif level == 2:
#             fsize = Pt(16)
#             self.current_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
#         else:
#             fsize = Pt(14)
#             self.current_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

#         # heading は基本的に太字
#         for child in children:
#             txt = self._extract_text(child)
#             run = self.current_paragraph.add_run(txt)
#             run.font.size = fsize
#             run.bold = True

#     #######################################################
#     # paragraph (段落)
#     #######################################################
#     def _render_paragraph(self, token):
#         """
#         token例:
#           { 'type':'paragraph', 'children':[...inline elements...] }
#         """
#         logging.debug(f"_render_paragraph: {token}")
#         self.current_paragraph = self.document.add_paragraph()
#         inline_children = token.get('children', [])
#         self._render_inline_children(inline_children)

#     def _render_inline_children(self, children):
#         """
#         段落内に複数の Run を生成し、強調(strong,emphasis)などを適宜反映
#         """
#         for node in children:
#             ntype = node['type']
#             if ntype == 'text':
#                 run = self.current_paragraph.add_run(node['text'])
#             elif ntype == 'strong':
#                 txt = self._extract_text(node)
#                 run = self.current_paragraph.add_run(txt)
#                 run.bold = True
#             elif ntype == 'emphasis':
#                 txt = self._extract_text(node)
#                 run = self.current_paragraph.add_run(txt)
#                 run.italic = True
#             else:
#                 # link/codespanなどの実装は要件に応じて追加
#                 txt = self._extract_text(node)
#                 run = self.current_paragraph.add_run(txt)

#     #######################################################
#     # list + list_item (入れ子対応)
#     #######################################################
#     def _render_list(self, token, level=0):
#         """
#         token例:
#           {
#             'type': 'list',
#             'ordered': True/False,
#             'children': [...list_item... or nested list...]
#           }
#         """
#         ordered = token.get('ordered', False)
#         for child in token.get('children', []):
#             if child['type'] == 'list_item':
#                 self._render_list_item(child, ordered, level)
#             elif child['type'] == 'list':
#                 # 入れ子リスト
#                 self._render_list(child, level=level+1)

#     def _render_list_item(self, token, ordered, level):
#         """
#         list_item: { 'type':'list_item', 'children':[...] }
#         """
#         style = 'List Number' if ordered else 'List Bullet'
#         self.current_paragraph = self.document.add_paragraph(style=style)
#         # levelに応じてインデントを増やす例 (0.5cm * level)
#         self.current_paragraph.paragraph_format.left_indent = Cm(0.5 * level)

#         # list_item内の要素を処理
#         for child in token.get('children', []):
#             ctype = child['type']
#             if ctype == 'paragraph':
#                 # リスト項目の段落
#                 inline_nodes = child.get('children', [])
#                 self._render_inline_children(inline_nodes)
#             elif ctype == 'list':
#                 # さらに入れ子のリスト
#                 self._render_list(child, level=level+1)
#             else:
#                 txt = self._extract_text(child)
#                 self.current_paragraph.add_run(txt)

#     #######################################################
#     # blockquote (引用)
#     #######################################################
#     def _render_blockquote(self, token):
#         """
#         引用ブロックを例示的に実装。スタイル 'Intense Quote' を適用。
#         """
#         self.current_paragraph = self.document.add_paragraph(style='Intense Quote')
#         for child in token.get('children', []):
#             if child['type'] == 'paragraph':
#                 text = ''.join(self._extract_text(n) for n in child.get('children', []))
#                 self.current_paragraph.add_run(text)

#     #######################################################
#     # thematic_break (水平線)
#     #######################################################
#     def _render_thematic_break(self, token):
#         """
#         簡易的に区切り線として "--------" を追加する例
#         """
#         hr_para = self.document.add_paragraph()
#         hr_para.add_run('--------------').bold = True

#     #######################################################
#     # table (表)
#     #######################################################
#     def _render_table(self, token):
#         """
#         plugin_table で生成されたAST:
#           {
#             'type': 'table',
#             'header': [ [cell1, cell2...], ... ],
#             'cells': [ [ [cell1, cell2...], [cell1, cell2...] ], ... ],
#             'align': [...]
#           }
#         """
#         header = token.get('header', [])
#         cells = token.get('cells', [])
#         if not header or not cells:
#             return

#         col_count = len(header)
#         table = self.document.add_table(rows=1, cols=col_count)
#         table.style = 'Table Grid'

#         # ヘッダー行
#         hdr_cells = table.rows[0].cells
#         for i, cell_ast in enumerate(header):
#             cell_txt = self._extract_text(cell_ast)
#             hdr_cells[i].text = cell_txt
#             # ヘッダーは太字
#             for p in hdr_cells[i].paragraphs:
#                 for r in p.runs:
#                     r.bold = True

#         # データ行
#         for row_data in cells:
#             row_cells = table.add_row().cells
#             for col_idx, cell_ast in enumerate(row_data):
#                 txt = self._extract_text(cell_ast)
#                 row_cells[col_idx].text = txt

#     #######################################################
#     # テキスト抽出用のヘルパー
#     #######################################################
#     def _extract_text(self, node):
#         """
#         子ノードを再帰的に走査してテキストを連結
#         strong/emphasis等の見出しはここでは単なるテキストとして結合
#         """
#         if 'text' in node:
#             return node['text']
#         elif 'children' in node:
#             return ''.join(self._extract_text(child) for child in node['children'])
#         return ''  # 該当なし

# def delete_file(path: str):
#     """
#     FileResponse返却後にバックグラウンドで削除するための後処理
#     """
#     try:
#         os.remove(path)
#         logging.info(f"[delete_file] Deleted file: {path}")
#     except Exception as e:
#         logging.error(f"[delete_file] Error: {e}")


# def generate_word_file(
#     background_tasks: BackgroundTasks,
#     summaries: dict,            # 例: { "Perplexity": {...}, "ChatGPT": {...} }
#     valuation_data: Optional[dict],
#     company_name: str,
#     file_name: Optional[str] = None,
# ) -> FileResponse:
#     """
#     1. MarkdownをDocxRenderer (Mistune) でWordに変換
#     2. valuation_dataを表形式で出力
#     3. Wordファイルを保存し、FileResponseで返却 (後処理でファイル削除)
#     """
    
#     logging.info(f"[generate_word_file] Received summaries: {summaries}")
#     logging.info(f"[generate_word_file] Received valuation_data: {valuation_data}")

#     ########################################
#     # 1) Wordドキュメント作成
#     ########################################
    
#     if not file_name.lower().endswith('.docx'):
#         file_name = file_name + '.docx'
#     # file_name = f"{company_name}_summary_report.docx"
#     document = Document()

#     # タイトル段落
#     title_para = document.add_paragraph(f"{company_name} - 要約レポート")
#     title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
#     run_title = title_para.runs[0]
#     run_title.font.size = Pt(18)
#     run_title.bold = True

#     ########################################
#     # 2) カスタムDocxRendererでMarkdownを解析
#     ########################################

#     docx_renderer = DocxRenderer(document)
#     md_parser = create_markdown(renderer=docx_renderer, plugins=[plugin_table])

#     # カテゴリ見出しのマッピング (任意)
#     category_mapping = {
#         "Perplexity": "Perplexity 分析",
#         "ChatGPT": "ChatGPT+SPEEDA 分析",
#     }

#     # セクション見出しのマッピング (任意)
#     reverse_key_mapping = {
#         "current_situation": "現状",
#         "future_outlook": "将来性と課題",
#         "investment_advantages": "競合と差別化",
#         "investment_disadvantages": "Exit先検討",
#         "value_up": "バリューアップ施策",
#         "use_case": "M&A事例",
#         "swot_analysis": "SWOT分析",
#     }


#     for main_category, sections in summaries.items():
#         # カテゴリ見出し
#         cat_jp = category_mapping.get(main_category, main_category)
#         cat_para = document.add_paragraph(cat_jp)
#         cat_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
#         cat_run = cat_para.runs[0]
#         cat_run.font.size = Pt(16)
#         cat_run.bold = True

#         # 各セクション
#         for idx, (sec_key, sec_text) in enumerate(sections.items(), start=1):
#             sec_jp = reverse_key_mapping.get(sec_key, sec_key)
#             sec_para = document.add_paragraph(f"{idx}. {sec_jp}")
#             sec_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
#             sec_run = sec_para.runs[0]
#             sec_run.font.size = Pt(14)
#             sec_run.bold = True

#             # Markdown → Word
#             md_parser(sec_text or "内容がありません")
            
#             ast_data = md_parser.parse(sec_text)
#             logging.info(f"Parsed AST for section={sec_key}: {ast_data}")

#     ########################################
#     # 3) バリュエーションテーブル追加 (オプション)
#     ########################################
#     if valuation_data:
#         val_para = document.add_paragraph("バリュエーション")
#         val_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
#         val_run = val_para.runs[0]
#         val_run.font.size = Pt(16)
#         val_run.bold = True

#         table = document.add_table(rows=1, cols=3)
#         table.style = "Table Grid"

#         hdr = table.rows[0].cells
#         hdr[0].text = "項目"
#         hdr[1].text = "直近実績"
#         hdr[2].text = "進行期見込"
#         for c_ in hdr:
#             for p_ in c_.paragraphs:
#                 for r_ in p_.runs:
#                     r_.font.bold = True
#                     r_.font.size = Pt(11)

#         # データ行追加
#         for label_, val_obj in valuation_data.items():
#             row_cells = table.add_row().cells
#             row_cells[0].text = label_
#             if isinstance(val_obj, dict):
#                 row_cells[1].text = val_obj.get("current", "不明")
#                 row_cells[2].text = val_obj.get("forecast", "不明")
#             else:
#                 row_cells[1].text = str(val_obj)
#                 row_cells[2].text = "不明"

#             # セル書式
#             for cell_ in row_cells:
#                 for p_ in cell_.paragraphs:
#                     p_.style = document.styles['Normal']
#                     for r_ in p_.runs:
#                         r_.font.size = Pt(10)

#     ########################################
#     # 4) Wordファイル保存 & FileResponse
#     ########################################
#     output_dir = "output"
#     os.makedirs(output_dir, exist_ok=True)
#     output_path = os.path.join(output_dir, file_name)
#     document.save(output_path)

#     background_tasks.add_task(delete_file, output_path)

#     return FileResponse(
#         output_path,
#         filename=file_name,
#         media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
#     )
    

############################################
# 以下fix前

    
def delete_file(path: str):
    try:
        os.remove(path)
        logging.info(f"ファイル {path} を削除しました。")
    except Exception as e:
        logging.error(f"ファイル削除エラー: {e}")


def generate_word_file(
    background_tasks: BackgroundTasks,
    summaries: dict = Body(..., description="要約データを含む辞書形式の入力"),
    valuation_data: dict = Body(None, description="バリュエーションデータ（オプション）"),
    company_name: str = Query(..., description="会社名を指定"),
    file_name: str = Query(None, description="生成するWordファイル名 (省略可能)")
) -> FileResponse:
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

    # 1) summaries をカテゴリ (Perplexity, ChatGPT...) ごとに処理
    for main_category, sections in summaries.items():
        # カテゴリ見出し
        jap_category = category_mapping.get(main_category, main_category)
        cat_heading = document.add_paragraph(jap_category)
        cat_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        cat_run = cat_heading.runs[0]
        cat_run.font.size = Pt(16)
        cat_run.bold = True

        # カテゴリ配下の各セクション
        for idx, (sec_key, sec_content) in enumerate(sections.items(), start=1):
            # セクション見出し
            jap_section = reverse_key_mapping.get(sec_key, sec_key)
            sec_heading = document.add_paragraph(f"{idx}. {jap_section}")
            sec_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
            heading_run = sec_heading.runs[0]
            heading_run.font.size = Pt(14)
            heading_run.bold = True

            # # セクション本文（Markdown からテーブルを抽出してみる）
            # raw_text = sec_content or "内容がありません"
            # tables = parse_markdown_table(raw_text)

            # if tables:
            #     # テーブルが見つかった場合は表として追加
            #     for table_data in tables:
            #         add_table_to_document(document, table_data)
            #         document.add_paragraph()  # 表の後に空白段落を追加

            #     # テーブル部分を除いた残りのテキストが欲しい場合は、
            #     # 追加のロジックで "表以外" を抜き出す必要がある。
            #     # シンプルに全体を段落に入れたい時は clean_text して段落にする:
            #     #  paragraph = document.add_paragraph(clean_text(raw_text), style='Normal')
            # else:
            #     # テーブルが無ければ普通の文章として処理
            #     paragraph = document.add_paragraph(clean_text(raw_text), style='Normal')
            #     paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


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