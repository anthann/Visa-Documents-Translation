#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Document Generator Service
Generates Word documents with translations
"""

from pathlib import Path
from typing import Dict, List, Optional

from docx import Document
from docx.shared import Emu, Inches, Pt, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Visible labels follow the translation's target language.
_LABELS = {
    'zh': {
        'table_title': '翻译对照表',
        'headers': ['项目', '内容'],
        'translation': '译文',
        'failed': '翻译失败',
        'translator': '译者信息',
        'name': '姓名',
        'phone': '电话',
        'address': '地址',
        'signature': '签名',
        'date': '日期',
        'confirm': '我确认本人能够从事翻译工作，且本译文是原件的准确翻译。',
    },
    'en': {
        'table_title': 'Translation table',
        'headers': ['Field', 'Text'],
        'translation': 'Translation',
        'failed': 'Translation failed',
        'translator': 'Translator details',
        'name': 'Full Name',
        'phone': 'Telephone number',
        'address': 'Address',
        'signature': 'Signature',
        'date': 'Date',
        'confirm': 'I confirm that I am capable of translation work and that this is an accurate translation of the original document.',
    },
    'ja': {
        'table_title': '対訳表',
        'headers': ['項目', '内容'],
        'translation': '訳文',
        'failed': '翻訳に失敗しました',
        'translator': '翻訳者情報',
        'name': '氏名',
        'phone': '電話番号',
        'address': '住所',
        'signature': '署名',
        'date': '日付',
        'confirm': '私は翻訳業務を行う能力があることを確認し、これは原文書の正確な翻訳です。',
    },
    'ko': {
        'table_title': '번역 대조표',
        'headers': ['항목', '내용'],
        'translation': '번역',
        'failed': '번역 실패',
        'translator': '번역자 정보',
        'name': '성명',
        'phone': '전화번호',
        'address': '주소',
        'signature': '서명',
        'date': '날짜',
        'confirm': '본인은 번역 업무를 수행할 능력이 있음을 확인하며, 이것은 원문의 정확한 번역입니다.',
    },
    'fr': {
        'table_title': 'Tableau de traduction',
        'headers': ['Champ', 'Texte'],
        'translation': 'Traduction',
        'failed': 'Échec de la traduction',
        'translator': 'Coordonnées du traducteur',
        'name': 'Nom complet',
        'phone': 'Téléphone',
        'address': 'Adresse',
        'signature': 'Signature',
        'date': 'Date',
        'confirm': "Je confirme que je suis capable d'effectuer un travail de traduction et qu'il s'agit d'une traduction exacte du document original.",
    },
    'de': {
        'table_title': 'Übersetzungstabelle',
        'headers': ['Feld', 'Text'],
        'translation': 'Übersetzung',
        'failed': 'Übersetzung fehlgeschlagen',
        'translator': 'Angaben zum Übersetzer',
        'name': 'Vollständiger Name',
        'phone': 'Telefonnummer',
        'address': 'Adresse',
        'signature': 'Unterschrift',
        'date': 'Datum',
        'confirm': 'Ich bestätige, dass ich zur Übersetzungsarbeit befähigt bin und dass dies eine genaue Übersetzung des Originaldokuments ist.',
    },
    'es': {
        'table_title': 'Tabla de traducción',
        'headers': ['Campo', 'Texto'],
        'translation': 'Traducción',
        'failed': 'Error de traducción',
        'translator': 'Datos del traductor',
        'name': 'Nombre completo',
        'phone': 'Teléfono',
        'address': 'Dirección',
        'signature': 'Firma',
        'date': 'Fecha',
        'confirm': 'Confirmo que soy capaz de realizar trabajos de traducción y que esta es una traducción exacta del documento original.',
    },
    'ru': {
        'table_title': 'Таблица перевода',
        'headers': ['Поле', 'Текст'],
        'translation': 'Перевод',
        'failed': 'Перевод не выполнен',
        'translator': 'Сведения о переводчике',
        'name': 'Полное имя',
        'phone': 'Телефон',
        'address': 'Адрес',
        'signature': 'Подпись',
        'date': 'Дата',
        'confirm': 'Подтверждаю, что обладаю квалификацией для выполнения перевода и что это точный перевод оригинального документа.',
    },
}

# Short document titles in the target language. Matched against the model's type text.
_DOC_TYPE_RULES = [
    (
        ('身份证', '居民身份证', 'id card', 'identity card', 'resident identity'),
        {
            'en': 'ID Card', 'zh': '身份证', 'ja': '身分証', 'ko': '신분증',
            'fr': "Carte d'identité", 'de': 'Personalausweis',
            'es': 'Documento de identidad', 'ru': 'Удостоверение личности',
        },
    ),
    (
        ('护照', 'passport'),
        {
            'en': 'Passport', 'zh': '护照', 'ja': '旅券', 'ko': '여권',
            'fr': 'Passeport', 'de': 'Reisepass', 'es': 'Pasaporte', 'ru': 'Паспорт',
        },
    ),
    (
        ('户口', '户籍', 'hukou', 'household register'),
        {
            'en': 'Household Register', 'zh': '户口本', 'ja': '戸籍簿', 'ko': '호구부',
            'fr': 'Registre de domicile', 'de': 'Haushaltsregister',
            'es': 'Registro de hogar', 'ru': 'Домовая книга',
        },
    ),
    (
        ('结婚', 'marriage'),
        {
            'en': 'Marriage Certificate', 'zh': '结婚证', 'ja': '結婚証明書', 'ko': '결혼증명서',
            'fr': 'Certificat de mariage', 'de': 'Heiratsurkunde',
            'es': 'Certificado de matrimonio', 'ru': 'Свидетельство о браке',
        },
    ),
    (
        ('出生证明', 'birth certificate'),
        {
            'en': 'Birth Certificate', 'zh': '出生证明', 'ja': '出生証明書', 'ko': '출생증명서',
            'fr': 'Acte de naissance', 'de': 'Geburtsurkunde',
            'es': 'Certificado de nacimiento', 'ru': 'Свидетельство о рождении',
        },
    ),
    (
        ('在职', '工作证明', 'employment certificate', 'work certificate'),
        {
            'en': 'Employment Certificate', 'zh': '在职证明', 'ja': '在職証明書', 'ko': '재직증명서',
            'fr': 'Certificat de travail', 'de': 'Arbeitsbescheinigung',
            'es': 'Certificado de empleo', 'ru': 'Справка с места работы',
        },
    ),
    (
        ('银行', '流水', 'bank statement'),
        {
            'en': 'Bank Statement', 'zh': '银行流水', 'ja': '銀行取引明細書', 'ko': '은행거래내역서',
            'fr': 'Relevé bancaire', 'de': 'Kontoauszug',
            'es': 'Extracto bancario', 'ru': 'Банковская выписка',
        },
    ),
    (
        ('毕业', '学历', '学位', 'diploma', 'degree certificate'),
        {
            'en': 'Diploma', 'zh': '学历证书', 'ja': '卒業証書', 'ko': '졸업증서',
            'fr': 'Diplôme', 'de': 'Diplom', 'es': 'Diploma', 'ru': 'Диплом',
        },
    ),
]

_LATIN_FONT = 'Arial'
_CJK_FONT = 'Microsoft YaHei'


class DocumentGenerator:
    """Generate Word documents with translations."""

    def __init__(self):
        pass

    def set_chinese_font(self, run, font_name: str = _LATIN_FONT, font_size: int = 11, bold: bool = False):
        """Set a black Arial run. Chinese falls back to Microsoft YaHei."""
        run.font.name = _LATIN_FONT
        run.font.size = Pt(font_size)
        run.bold = bold
        run.italic = False
        run.font.color.rgb = RGBColor(0, 0, 0)
        r = run._element
        rPr = r.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn('w:ascii'), _LATIN_FONT)
        rFonts.set(qn('w:hAnsi'), _LATIN_FONT)
        rFonts.set(qn('w:cs'), _LATIN_FONT)
        rFonts.set(qn('w:eastAsia'), _CJK_FONT)
        color = rPr.find(qn('w:color'))
        if color is None:
            color = OxmlElement('w:color')
            rPr.append(color)
        color.set(qn('w:val'), '000000')
        if color.get(qn('w:themeColor')) is not None:
            del color.attrib[qn('w:themeColor')]

    def _set_page_margins(self, doc: Document):
        """Use margins that leave a 7-inch text width on US Letter."""
        for section in doc.sections:
            section.left_margin = Inches(0.75)
            section.right_margin = Inches(0.75)
            section.top_margin = Inches(0.75)
            section.bottom_margin = Inches(0.75)

    def _content_width(self, doc: Document):
        section = doc.sections[0]
        return section.page_width - section.left_margin - section.right_margin

    def _set_table_column_widths(self, doc: Document, table, ratios):
        """Pin table and cell widths so Word cannot stretch the table past the page."""
        # Page measurements are EMUs; table widths are twips (dxa).
        content = int(Emu(self._content_width(doc)).twips)
        ratio_sum = sum(ratios)
        widths = [int(content * ratio / ratio_sum) for ratio in ratios]
        widths[-1] = content - sum(widths[:-1])

        tbl = table._tbl
        tblPr = tbl.tblPr
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = OxmlElement('w:tblW')
            tblPr.append(tblW)
        tblW.set(qn('w:type'), 'dxa')
        tblW.set(qn('w:w'), str(sum(widths)))

        tblInd = tblPr.find(qn('w:tblInd'))
        if tblInd is None:
            tblInd = OxmlElement('w:tblInd')
            tblPr.append(tblInd)
        tblInd.set(qn('w:type'), 'dxa')
        tblInd.set(qn('w:w'), '0')

        layout = tblPr.find(qn('w:tblLayout'))
        if layout is None:
            layout = OxmlElement('w:tblLayout')
            tblPr.append(layout)
        layout.set(qn('w:type'), 'fixed')

        tblCellMar = tblPr.find(qn('w:tblCellMar'))
        if tblCellMar is not None:
            tblPr.remove(tblCellMar)
        tblCellMar = OxmlElement('w:tblCellMar')
        for edge, dxa in (('top', 100), ('bottom', 100), ('left', 140), ('right', 140)):
            margin = OxmlElement(f'w:{edge}')
            margin.set(qn('w:w'), str(dxa))
            margin.set(qn('w:type'), 'dxa')
            tblCellMar.append(margin)
        tblPr.append(tblCellMar)

        grid = tbl.tblGrid
        for child, width in zip(list(grid), widths):
            child.set(qn('w:w'), str(width))

        for row in table.rows:
            for cell, width in zip(row.cells, widths):
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcW = tcPr.find(qn('w:tcW'))
                if tcW is None:
                    tcW = OxmlElement('w:tcW')
                    tcPr.append(tcW)
                tcW.set(qn('w:type'), 'dxa')
                tcW.set(qn('w:w'), str(width))
                tcMar = tcPr.find(qn('w:tcMar'))
                if tcMar is not None:
                    tcPr.remove(tcMar)

    def set_cell_border(self, cell, **kwargs):
        """Set cell border, replacing any previous tcBorders element."""
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        existing = tcPr.find(qn('w:tcBorders'))
        if existing is not None:
            tcPr.remove(existing)

        tcBorders = OxmlElement('w:tcBorders')
        for edge in ('top', 'left', 'bottom', 'right'):
            edge_data = kwargs.get(edge)
            if edge_data:
                tag = f'w:{edge}'
                element = OxmlElement(tag)
                element.set(qn('w:val'), edge_data.get('val', 'single'))
                element.set(qn('w:sz'), str(edge_data.get('sz', 4)))
                element.set(qn('w:color'), edge_data.get('color', '000000'))
                tcBorders.append(element)

        tcPr.append(tcBorders)

    def set_cell_shading(self, cell, fill_hex: str):
        """Fill a table cell background with a solid hex color."""
        tcPr = cell._tc.get_or_add_tcPr()
        existing = tcPr.find(qn('w:shd'))
        if existing is not None:
            tcPr.remove(existing)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), fill_hex)
        tcPr.append(shd)

    def generate_document(self, translations: List[Dict], output_path: str,
                         source_lang: str = 'zh', target_lang: str = 'en',
                         merge: bool = True) -> str:
        """
        Generate Word document with translations.

        Args:
            translations: List of translation results
            output_path: Path to save the document
            source_lang: Source language code
            target_lang: Target language code
            merge: Whether to merge into single document

        Returns:
            Path to generated document
        """
        doc = Document()

        # Arial for Latin text, Microsoft YaHei for Chinese, both black.
        style = doc.styles['Normal']
        style.font.name = _LATIN_FONT
        style.font.size = Pt(11)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.line_spacing = 1.15
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn('w:ascii'), _LATIN_FONT)
        rFonts.set(qn('w:hAnsi'), _LATIN_FONT)
        rFonts.set(qn('w:cs'), _LATIN_FONT)
        rFonts.set(qn('w:eastAsia'), _CJK_FONT)
        self._set_page_margins(doc)

        # Add each translation section
        for i, trans in enumerate(translations):
            self._add_section(doc, trans, i + 1, target_lang)

            # Add page break between sections (except last)
            if i < len(translations) - 1:
                doc.add_page_break()

        # Visitor-visa translator details
        self._add_certification(doc, target_lang)

        # Save document
        doc.save(output_path)
        return output_path

    def _labels(self, target_lang: str) -> Dict[str, object]:
        """Return UI strings for the target language."""
        return _LABELS.get(target_lang, _LABELS['en'])

    def _document_title(self, translation: Dict, target_lang: str) -> str:
        """Return a short document-type name in the target language."""
        blob = ' '.join([
            str(translation.get('document_type_original') or ''),
            str(translation.get('document_type') or ''),
        ]).lower()
        for keys, names in _DOC_TYPE_RULES:
            if any(key.lower() in blob for key in keys):
                return names.get(target_lang, names['en'])
        typed = str(translation.get('document_type') or '').strip()
        if typed and typed.lower() not in ('unknown', 'unknown document'):
            return typed
        return {'zh': '文件', 'ja': '文書', 'ko': '문서', 'fr': 'Document',
                'de': 'Dokument', 'es': 'Documento', 'ru': 'Документ'}.get(target_lang, 'Document')

    def _add_title(self, doc: Document, text: str, space_before_pt: float = 0):
        """Section title shared by the document type and the translator details."""
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.space_before = Pt(space_before_pt)
        para.paragraph_format.space_after = Pt(8)
        para.paragraph_format.keep_with_next = True
        run = para.add_run(text)
        self.set_chinese_font(run, font_size=16, bold=True)
        return para

    def _style_table_cell(self, cell, min_height_pt: Optional[float] = None):
        """Center cell contents vertically and strip default paragraph spacing."""
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for para in cell.paragraphs:
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.line_spacing = 1.15
        if min_height_pt is not None:
            trPr = cell._tc.getparent().get_or_add_trPr()
            if trPr.find(qn('w:trHeight')) is None:
                tr_height = OxmlElement('w:trHeight')
                tr_height.set(qn('w:val'), str(int(min_height_pt * 20)))
                tr_height.set(qn('w:hRule'), 'atLeast')
                trPr.append(tr_height)
            if trPr.find(qn('w:cantSplit')) is None:
                trPr.append(OxmlElement('w:cantSplit'))

    def _add_section(self, doc: Document, trans: Dict, section_num: int,
                    target_lang: str):
        """Add a translation section."""
        translation = trans.get('translation', {})

        self._add_title(doc, self._document_title(translation, target_lang))

        labels = self._labels(target_lang)

        # Add translation table
        trans_title = doc.add_paragraph()
        trans_title.paragraph_format.space_before = Pt(4)
        trans_title.paragraph_format.space_after = Pt(6)
        trans_title.paragraph_format.keep_with_next = True
        run = trans_title.add_run(labels['table_title'])
        self.set_chinese_font(run, font_size=13, bold=True)

        fields = translation.get('fields', [])
        if fields:
            self._add_translation_table(doc, fields, target_lang)
        else:
            # If no structured fields, show full text
            full_translated = translation.get('full_text_translated', '')

            if full_translated:
                para = doc.add_paragraph()
                run = para.add_run(labels['translation'])
                self.set_chinese_font(run, "SimHei", 11, bold=True)

                trans_para = doc.add_paragraph()
                run = trans_para.add_run(full_translated)
                self.set_chinese_font(run, "SimSun", 10)
            else:
                raw = translation.get('raw_response', labels['failed'])
                para = doc.add_paragraph()
                run = para.add_run(raw)
                self.set_chinese_font(run, "SimSun", 10)

    def _add_translation_table(self, doc: Document, fields: List[Dict], target_lang: str):
        """Add translation comparison table."""
        if not fields:
            return

        # Header plus two rows (original, translation) for each field.
        table = doc.add_table(rows=len(fields) * 2 + 1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        self._set_table_column_widths(doc, table, (2.2, 4.8))

        outer_heavy = {'val': 'single', 'sz': 8, 'color': '000000'}
        group_line = {'val': 'single', 'sz': 6, 'color': '000000'}
        pair_inner = {'val': 'single', 'sz': 2, 'color': 'D0D0D0'}
        col_divider = {'val': 'single', 'sz': 4, 'color': 'BFBFBF'}

        # Header row
        headers = self._labels(target_lang)['headers']
        header_row = table.rows[0]
        header_trPr = header_row._tr.get_or_add_trPr()
        if header_trPr.find(qn('w:tblHeader')) is None:
            header_trPr.append(OxmlElement('w:tblHeader'))
        for i, header_text in enumerate(headers):
            cell = header_row.cells[i]
            self._style_table_cell(cell, min_height_pt=24)
            self.set_cell_shading(cell, 'EAEAEA')
            self.set_cell_border(
                cell,
                top=outer_heavy,
                bottom=outer_heavy,
                left=group_line if i == 0 else col_divider,
                right=col_divider if i == 0 else group_line,
            )
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(header_text)
            self.set_chinese_font(run, font_size=10, bold=True)

        # Each field is an original row followed by its translation, visually grouped as a pair.
        for i, field in enumerate(fields):
            original_row = table.rows[i * 2 + 1]
            translated_row = table.rows[i * 2 + 2]
            band_fill = 'F6F7F8' if i % 2 == 1 else 'FFFFFF'
            is_last_group = (i == len(fields) - 1)

            for col_idx in (0, 1):
                left_edge = group_line if col_idx == 0 else col_divider
                right_edge = col_divider if col_idx == 0 else group_line
                self.set_cell_shading(original_row.cells[col_idx], band_fill)
                self.set_cell_shading(translated_row.cells[col_idx], band_fill)
                self.set_cell_border(
                    original_row.cells[col_idx],
                    top=outer_heavy if i == 0 else group_line,
                    bottom=pair_inner,
                    left=left_edge,
                    right=right_edge,
                )
                self.set_cell_border(
                    translated_row.cells[col_idx],
                    top=pair_inner,
                    bottom=outer_heavy if is_last_group else group_line,
                    left=left_edge,
                    right=right_edge,
                )

            pairs = (
                (original_row, field.get('label_original', ''), field.get('value_original', ''), True),
                (translated_row, field.get('label_translated', ''), field.get('value_translated', ''), False),
            )
            for row, label, value, keep_pair_together in pairs:
                self._style_table_cell(row.cells[0], min_height_pt=22)
                self._style_table_cell(row.cells[1], min_height_pt=22)
                if keep_pair_together:
                    row.cells[0].paragraphs[0].paragraph_format.keep_with_next = True
                    row.cells[1].paragraphs[0].paragraph_format.keep_with_next = True
                label_run = row.cells[0].paragraphs[0].add_run(label)
                self.set_chinese_font(label_run, font_size=10, bold=True)
                value_run = row.cells[1].paragraphs[0].add_run(value)
                self.set_chinese_font(value_run, font_size=10)

    def _add_certification(self, doc: Document, target_lang: str):
        """Add blank translator details in the target language."""
        labels = self._labels(target_lang)

        self._add_title(doc, labels['translator'], space_before_pt=18)

        confirm = doc.add_paragraph()
        confirm.paragraph_format.space_before = Pt(0)
        confirm.paragraph_format.space_after = Pt(8)
        confirm.paragraph_format.keep_with_next = True
        run = confirm.add_run(labels['confirm'])
        self.set_chinese_font(run, font_size=11)

        items = (
            labels['signature'],
            labels['name'],
            labels['phone'],
            labels['address'],
            labels['date'],
        )
        for idx, label in enumerate(items):
            para = doc.add_paragraph()
            # Extra room directly above Signature for placing an e-signature, without blank paragraphs.
            para.paragraph_format.space_before = Pt(28 if idx == 0 else 0)
            para.paragraph_format.space_after = Pt(8)
            if idx < len(items) - 1:
                para.paragraph_format.keep_with_next = True
            run = para.add_run(f"{label}: ________________________")
            self.set_chinese_font(run, font_size=11)
