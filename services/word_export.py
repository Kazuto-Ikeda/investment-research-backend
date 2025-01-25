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
from docx.oxml import OxmlElement
from docx.oxml.ns import qn



class DocxRenderer(mistune.AstRenderer):
    """
    Mistune 2.x の AstRenderer を継承し、Markdown の AST を Python-docx で Word に変換。
    
    主な対応要素:
      - 見出し(heading) -> Word段落+run（太字/フォントサイズ）
      - 段落(paragraph) -> Word段落
      - 箇条書き(list, list_item) -> List Bullet/Number スタイル, 入れ子リスト対応
      - テーブル(table) -> plugin_table が生成したASTをそのまま受け取り、Word表に変換
      - 強調(strong, emphasis) -> Runの bold/italic
      - blockquote(引用), thematic_break(水平線) -> サンプル実装
    
    段落の途中で太字や斜体にしたい場合、Runを分割して追加しています。
    """

    def __init__(self, document: Document):
        super().__init__()
        self.document = document
        self.current_paragraph = None  # 処理中の段落オブジェクト
        self._reset_numbering_for_next_list = False
        


    def render(self, tokens, state):
        """
        AST全体を走査し、トークン種類に応じて _render_xxx() メソッドを呼び出す。
        """
        for token in tokens:
            node_type = token['type']
            if node_type == 'heading':
                self._render_heading(token)
            elif node_type == 'paragraph':
                self.current_paragraph = self.document.add_paragraph()
                inline_children = token.get('children', [])
                self._render_inline_children(inline_children)
            elif node_type == 'list':
                self._render_list(token, level=1)
            elif node_type == 'table':
                self._render_table(token)
            else:
                logging.debug(f"[DocxRenderer] Skip unknown node: {node_type}")
        return ''  # 文字列は返さず、Word文書に直接書き込む


    #######################################################
    # heading (見出し)
    #######################################################
    def _render_heading(self, token):
        """
        見出しを描画し、描画後に「次のリストは番号をリセットする」フラグを立てる。
        """
        level = token['level']
        children = token.get('children', [])

        self.current_paragraph = self.document.add_paragraph()

        if level == 1:
            fsize = Pt(18)
            self.current_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif level == 2:
            fsize = Pt(16)
            self.current_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            fsize = Pt(14)
            self.current_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

        for child in children:
            txt = self._extract_text(child)
            run = self.current_paragraph.add_run(txt)
            run.font.size = fsize
            run.bold = True

        # ★この見出しの直後からリストが始まった場合に 1から振り直したい場合
        self._reset_numbering_for_next_list = True

    #######################################################
    # paragraph (段落)
    #######################################################
    def _render_paragraph(self, token):
        """
        token例:
          { 'type':'paragraph', 'children':[...inline elements...] }
        """
        logging.debug(f"_render_paragraph: {token}")
        self.current_paragraph = self.document.add_paragraph()
        inline_children = token.get('children', [])
        self._render_inline_children(inline_children)

    def _render_inline_children(self, children):
        for node in children:
            ntype = node['type']
            if ntype == 'text':
                run = self.current_paragraph.add_run(node['text'])
            elif ntype == 'strong':
                txt = self._extract_text(node)
                run = self.current_paragraph.add_run(txt)
                run.bold = True
            elif ntype == 'emphasis':
                txt = self._extract_text(node)
                run = self.current_paragraph.add_run(txt)
                run.italic = True
            else:
                txt = self._extract_text(node)
                run = self.current_paragraph.add_run(txt)

    #######################################################
    # list + list_item (入れ子対応)
    #######################################################
    def _restart_numbering(self, paragraph, level=0, num_id=1):
        """
        指定した段落を num_id で与えられるリスト定義に属する段落として設定し、
        レベル(level) も設定する。結果的にその段落の番号付けを再スタートできる。
        """
        p = paragraph._p  # 段落の内部オブジェクト
        pPr = p.get_or_add_pPr()  
        numPr = pPr.get_or_add_numPr()

        # 既存の <w:ilvl> や <w:numId> を削除
        for child in numPr.iterchildren():
            numPr.remove(child)

        # <w:ilvl w:val="0"/>
        ilvl = OxmlElement('w:ilvl')
        ilvl.set(qn('w:val'), str(level))
        numPr.append(ilvl)

        # <w:numId w:val="1"/>
        numId_elm = OxmlElement('w:numId')
        numId_elm.set(qn('w:val'), str(num_id))
        numPr.append(numId_elm)


    def _render_list(self, token, level=1):
        ordered = token.get('ordered', False)
        for child in token.get('children', []):
            if child['type'] == 'list_item':
                self._render_list_item(child, ordered, level)
            elif child['type'] == 'list':
                self._render_list(child, level=level+1)

    def _render_list_item(self, token, ordered, level):
        """
        list_item: { 'type':'list_item', 'children':[...] }
        """
        style = 'List Number' if ordered else 'List Bullet'
        self.current_paragraph = self.document.add_paragraph(style=style)
        self.current_paragraph.paragraph_format.left_indent = Cm(0.5 * level)

        # ★「見出しの直後の最初の orderedリスト（番号リスト）」でリセットしたい場合
        #   ここでは「最初の“番号付き”list_item が来たらリセット」のロジックにしています。
        if ordered and self._reset_numbering_for_next_list:
            # 引数を位置引数だけにするか、キーワードだけにするかで
            # "got multiple values for argument" エラーを回避
            # ここでは位置引数を使って呼ぶ例に統一
            self._restart_numbering(self.current_paragraph, 0, 1)

            # リセットは1回だけ
            self._reset_numbering_for_next_list = False

        # list_item内の要素を処理
        for child in token.get('children', []):
            ctype = child['type']
            if ctype == 'paragraph':
                inline_nodes = child.get('children', [])
                self._render_inline_children(inline_nodes)
            elif ctype == 'list':
                self._render_list(child, level=level+1)
            else:
                txt = self._extract_text(child)
                self.current_paragraph.add_run(txt)


    #######################################################
    # blockquote (引用)
    #######################################################
    def _render_blockquote(self, token):
        """
        引用ブロックを例示的に実装。スタイル 'Intense Quote' を適用。
        """
        self.current_paragraph = self.document.add_paragraph(style='Intense Quote')
        for child in token.get('children', []):
            if child['type'] == 'paragraph':
                text = ''.join(self._extract_text(n) for n in child.get('children', []))
                self.current_paragraph.add_run(text)

    #######################################################
    # thematic_break (水平線)
    #######################################################
    def _render_thematic_break(self, token):
        """
        簡易的に区切り線として "--------" を追加する例
        """
        hr_para = self.document.add_paragraph()
        hr_para.add_run('--------------').bold = True

    #######################################################
    # table (表)
    #######################################################
    def _render_table(self, token):
        """
        Mistune 2.x plugin_table で生成されたトークンに合わせた実装。
        token 例:

        {
        'type': 'table',
        'children': [
            {
            'type': 'table_head',
            'children': [
                { 'type': 'table_cell', 'children': [...], 'align': None, 'is_head': True },
                ...
            ]
            },
            {
            'type': 'table_body',
            'children': [
                {
                'type': 'table_row',
                'children': [
                    { 'type': 'table_cell', 'children': [...], 'align': None, 'is_head': False },
                    ...
                ]
                },
                ...
            ]
            }
        ]
        }
        """

        # table_head と table_body を見つける
        table_head = None
        table_cell = None
        for child in token.get('children', []):
            if child['type'] == 'table_head':
                table_head = child
            elif child['type'] == 'table_body':
                table_cell = child
        
        # ヘッダー情報の取得
        # table_head["children"] は table_cell 群 (1行のみ)
        if table_head and 'children' in table_head:
            head_cells = table_head['children']
            col_count = len(head_cells)
        else:
            # ヘッダーなしの場合 → body の最初の row から列数を推定
            if table_cell and 'children' in table_cell and len(table_cell['children']) > 0:
                first_row = table_cell['children'][0]
                if first_row.get('type') == 'table_row':
                    col_count = len(first_row.get('children', []))
                else:
                    col_count = 0
            else:
                col_count = 0

        # テーブルを作成
        if col_count == 0:
            return

        table = self.document.add_table(rows=1, cols=col_count)
        table.style = 'Table Grid'

        # (1) テーブルヘッダーの描画
        if table_head and table_head.get('children'):
            # rows[0] にヘッダーを設定
            hdr_cells = table.rows[0].cells
            for i, cell_ast in enumerate(table_head['children']):
                cell_text = self._extract_text(cell_ast)  # 下記ヘルパーで子要素を走査
                hdr_cells[i].text = cell_text
                # ヘッダーは太字
                for p in hdr_cells[i].paragraphs:
                    for r in p.runs:
                        r.bold = True
        else:
            # ヘッダーなし → 空の見出し行を作るだけ
            pass

        # (2) テーブルボディの描画
        if table_cell and table_cell.get('children'):
            for row_ast in table_cell['children']:
                # row_ast: { 'type': 'table_row', 'children': [ ... table_cell ... ] }
                if row_ast.get('type') != 'table_row':
                    continue
                
                # 新規行を追加
                row_cells = table.add_row().cells
                for col_idx, cell_ast in enumerate(row_ast.get('children', [])):
                    cell_text = self._extract_text(cell_ast)
                    row_cells[col_idx].text = cell_text
        else:
            # 本文なし
            pass

    #######################################################
    # テキスト抽出用のヘルパー
    #######################################################
    def _extract_text(self, node):
        if 'text' in node:
            return node['text']
        elif 'children' in node:
            return ''.join(self._extract_text(child) for child in node['children'])
        return ''

def delete_file(path: str):
    """
    FileResponse返却後にバックグラウンドで削除するための後処理
    """
    try:
        os.remove(path)
        logging.info(f"[delete_file] Deleted file: {path}")
    except Exception as e:
        logging.error(f"[delete_file] Error: {e}")


def generate_word_file(
    background_tasks: BackgroundTasks,
    summaries: dict,            # 例: { "Perplexity": {...}, "ChatGPT": {...} }
    valuation_data: Optional[dict],
    company_name: str,
    file_name: Optional[str] = "result",
) -> FileResponse:
    """
    1. MarkdownをDocxRenderer (Mistune) でWordに変換
    2. valuation_dataを表形式で出力
    3. Wordファイルを保存し、FileResponseで返却 (後処理でファイル削除)
    """
    
    ########################################
    # 1) Wordドキュメント作成
    ########################################
    
    file_name = f"{company_name}_summary_report.docx"
    document = Document()
    
    # ドキュメントのデフォルトフォントを設定
    style = document.styles['Normal']
    font = style.font
    font.name = 'Noto Sans JP'
    font.element.rPr.rFonts.set(qn('w:eastAsia'), 'Noto Sans JP')

    # タイトル段落
    title_para = document.add_paragraph(f"{company_name} - 要約レポート")
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_para.runs[0]
    run_title.font.size = Pt(18)
    run_title.bold = True

    ########################################
    # 2) カスタムDocxRendererでMarkdownを解析
    ########################################

    docx_renderer = DocxRenderer(document)
    md_parser = create_markdown(renderer=docx_renderer, plugins=[plugin_table])

    # カテゴリ見出しのマッピング (任意)
    category_mapping = {
        "Perplexity": "Perplexity 分析",
        "ChatGPT": "ChatGPT+SPEEDA 分析",
    }

    # セクション見出しのマッピング (任意)
    reverse_key_mapping = {
        "current_situation": "現状",
        "future_outlook": "将来性と課題",
        "investment_advantages": "競合と差別化",
        "investment_disadvantages": "Exit先検討",
        "value_up": "バリューアップ施策",
        "use_case": "M&A事例",
        "swot_analysis": "SWOT分析",
    }


    for main_category, sections in summaries.items():
        # カテゴリ見出し
        cat_jp = category_mapping.get(main_category, main_category)
        cat_para = document.add_paragraph(cat_jp)
        cat_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        cat_run = cat_para.runs[0]
        cat_run.font.size = Pt(16)
        cat_run.bold = True
        

        # 各セクション
        for idx, (sec_key, sec_text) in enumerate(sections.items(), start=1):
            sec_jp = reverse_key_mapping.get(sec_key, sec_key)
            sec_para = document.add_paragraph(f"{idx}. {sec_jp}")
            sec_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            sec_run = sec_para.runs[0]
            sec_run.font.size = Pt(14)
            sec_run.bold = True

            # Markdown → Word
            # 1) パースのみ（Markdown → ASTトークン）
            tokens = md_parser.parse(sec_text or "内容がありません")

            # 2) レンダリング実行（ASTトークン → docx_renderer）
            #   state は省略可・または {} などでOK
            docx_renderer.render(tokens, state={})
            



    ########################################
    # 3) バリュエーションテーブル追加 (オプション)
    ########################################
    if valuation_data:
        val_para = document.add_paragraph("バリュエーション")
        val_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        val_run = val_para.runs[0]
        val_run.font.size = Pt(16)
        val_run.bold = True

        table = document.add_table(rows=1, cols=3)
        table.style = "Table Grid"

        hdr = table.rows[0].cells
        hdr[0].text = "項目"
        hdr[1].text = "直近実績"
        hdr[2].text = "進行期見込"
        for c_ in hdr:
            for p_ in c_.paragraphs:
                for r_ in p_.runs:
                    r_.font.bold = True
                    r_.font.size = Pt(10)

        # データ行追加
        for label_, val_obj in valuation_data.items():
            row_cells = table.add_row().cells
            row_cells[0].text = label_
            if isinstance(val_obj, dict):
                row_cells[1].text = val_obj.get("current", "不明")
                row_cells[2].text = val_obj.get("forecast", "不明")
            else:
                row_cells[1].text = str(val_obj)
                row_cells[2].text = "不明"

            # セル書式
            for cell_ in row_cells:
                for p_ in cell_.paragraphs:
                    p_.style = document.styles['Normal']
                    for r_ in p_.runs:
                        r_.font.size = Pt(10)

    ########################################
    # 4) Wordファイル保存 & FileResponse
    ########################################
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, file_name)
    document.save(output_path)

    background_tasks.add_task(delete_file, output_path)

    return FileResponse(
        output_path,
        filename=file_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    

