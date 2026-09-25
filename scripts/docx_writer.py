#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx_writer —— 从零构造 .docx 的最小写入器（零依赖）。

用途：生成**新的** Word 文档，典型场景是「改写对照表」——横向页面 + 大表格 +
标题层级。不负责编辑已有文档（那是 :mod:`docx_io` 的职责）。

刻意采用「全内联格式」：字体、字号、颜色、边框、底纹全部写在 ``rPr`` /
``pPr`` / ``tcPr`` 里，不依赖 styles.xml 的样式继承。好处是产出的文件在
Word / WPS / LibreOffice / 各种在线预览器里表现一致，不会因为对方缺样式定义
而走形。styles.xml 只提供 docDefaults 与 Normal 兜底。

依赖：Python 3.8+ 标准库。
"""
from __future__ import annotations

import zipfile

NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS_DC = 'http://purl.org/dc/elements/1.1/'
NS_CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
NS_DCTERMS = 'http://purl.org/dc/terms/'
NS_XSI = 'http://www.w3.org/2001/XMLSchema-instance'
NS_EP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'

#: 常用中文字体（Word 通过 eastAsia 属性匹配中文字形）
SONG = '宋体'
HEI = '黑体'
KAI = '楷体'


def _esc(s: str) -> str:
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;'))


class DocxBuilder:
    """极简 docx 构造器。

    :param orientation: ``portrait`` 或 ``landscape``
    :param font: 默认中文字体名
    :param size: 默认字号（磅），10.5 = 五号
    """

    def __init__(self, orientation: str = 'portrait', font: str = SONG,
                 size: float = 10.5):
        self.orientation = orientation
        self.font = font
        self.size = size
        self._body: list[str] = []

    # ── 底层片段 ────────────────────────────────────────────────────────
    def _rpr(self, bold=False, italic=False, size=None, color=None,
             font=None, underline=False) -> str:
        f = font or self.font
        sz = int(round((size or self.size) * 2))       # half-points
        parts = ['<w:rPr>']
        parts.append('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
                     % (_esc(f), _esc(f), _esc(f), _esc(f)))
        if bold:
            parts.append('<w:b/><w:bCs/>')
        if italic:
            parts.append('<w:i/><w:iCs/>')
        if underline:
            parts.append('<w:u w:val="single"/>')
        if color:
            parts.append('<w:color w:val="%s"/>' % _esc(color))
        parts.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz))
        parts.append('</w:rPr>')
        return ''.join(parts)

    def _run(self, text: str, **kw) -> str:
        return ('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                % (self._rpr(**kw), _esc(text)))

    def _ppr(self, align=None, indent_first=None, space_before=None,
             space_after=None, line=None, outline=None, keep_next=False) -> str:
        parts = ['<w:pPr>']
        if keep_next:
            parts.append('<w:keepNext/>')
        if outline is not None:
            parts.append('<w:outlineLvl w:val="%d"/>' % outline)
        if indent_first:
            parts.append('<w:ind w:firstLineChars="%d" w:firstLine="%d"/>'
                         % (indent_first, int(indent_first * self.size * 20 / 100)))
        if space_before is not None or space_after is not None or line:
            a = ['<w:spacing']
            if space_before is not None:
                a.append(' w:before="%d"' % int(space_before * 20))
            if space_after is not None:
                a.append(' w:after="%d"' % int(space_after * 20))
            if line:
                a.append(' w:line="%d" w:lineRule="auto"' % int(line * 240))
            a.append('/>')
            parts.append(''.join(a))
        if align:
            parts.append('<w:jc w:val="%s"/>' % align)
        parts.append('</w:pPr>')
        return ''.join(parts)

    # ── 内容 ────────────────────────────────────────────────────────────
    def _p_xml(self, text: str = '', bold=False, size=None, color=None,
               align=None, indent_first=0, space_before=None,
               space_after=2, font=None, italic=False, outline=None,
               keep_next=False) -> str:
        """生成一个 ``w:p`` 的 XML 片段（不落盘、不产生副作用）。

        单元格、正文都复用这个函数——它必须是纯函数，否则表格内容会
        被错误地追加到文档顶部。
        """
        runs = []
        lines = str(text).split('\n')
        for i, ln in enumerate(lines):
            if i:
                runs.append('<w:r><w:br/></w:r>')
            if ln:
                runs.append(self._run(ln, bold=bold, size=size, color=color,
                                      font=font, italic=italic))
        return '<w:p>%s%s</w:p>' % (
            self._ppr(align=align, indent_first=indent_first,
                      space_before=space_before, space_after=space_after,
                      outline=outline, keep_next=keep_next),
            ''.join(runs))

    def heading(self, text: str, level: int = 1) -> 'DocxBuilder':
        """标题。level 1/2/3 对应字号 16/14/12 磅，并写入 outlineLvl 便于导航。"""
        cfg = {1: (16, 8, 6), 2: (14, 7, 5), 3: (12, 6, 4)}
        size, sb, sa = cfg.get(level, cfg[3])
        self._body.append(self._p_xml(text, bold=True, size=size, font=HEI,
                                      space_before=sb, space_after=sa,
                                      outline=level - 1, keep_next=True))
        return self

    def paragraph(self, text: str = '', **kw) -> 'DocxBuilder':
        """普通段落。支持用 ``\\n`` 在段内换行。"""
        self._body.append(self._p_xml(text, **kw))
        return self

    def spacer(self, size: float = 10.5) -> 'DocxBuilder':
        self._body.append('<w:p>%s</w:p>' % self._ppr(space_after=0))
        return self

    def page_break(self) -> 'DocxBuilder':
        self._body.append(
            '<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        return self

    def _cell(self, text, width, bold=False, fill=None, size=None,
              align=None, color=None, valign='top') -> str:
        shd = ('<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % fill) if fill else ''
        return (
            '<w:tc><w:tcPr>'
            '<w:tcW w:w="%d" w:type="dxa"/>'
            '%s'
            '<w:vAlign w:val="%s"/>'
            '</w:tcPr>'
            '%s'
            '</w:tc>' % (
                int(width), shd, valign,
                self._p_xml(text, bold=bold, size=size, align=align,
                            color=color, space_after=0)))

    def table(self, header: list[str], rows: list[list], widths: list[float] | None = None,
              header_fill: str = 'DCE6F1', total_width: int = 14570,
              size: float = 9, repeat_header: bool = True,
              zebra: str | None = None) -> 'DocxBuilder':
        """插入表格。

        :param widths: 各列宽度比例（和为 1 即可），默认等分
        :param total_width: 表格总宽（dxa，1/20 磅）。A4 横向正文宽约 14570
        """
        n = len(header)
        ratios = widths or [1.0 / n] * n
        cols = [int(total_width * r) for r in ratios]
        cols[-1] = total_width - sum(cols[:-1])

        borders = ('<w:tblBorders>'
                   + ''.join('<w:%s w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>' % s
                             for s in ('top', 'left', 'bottom', 'right',
                                       'insideH', 'insideV'))
                   + '</w:tblBorders>')
        out = ['<w:tbl><w:tblPr>'
               '<w:tblW w:w="%d" w:type="dxa"/>' % total_width,
               borders,
               '<w:tblLayout w:type="fixed"/>'
               '<w:tblCellMar>'
               '<w:top w:w="40" w:type="dxa"/><w:left w:w="70" w:type="dxa"/>'
               '<w:bottom w:w="40" w:type="dxa"/><w:right w:w="70" w:type="dxa"/>'
               '</w:tblCellMar>'
               '</w:tblPr><w:tblGrid>']
        out += ['<w:gridCol w:w="%d"/>' % c for c in cols]
        out.append('</w:tblGrid>')

        # 表头
        tr = ['<w:tr><w:trPr>']
        if repeat_header:
            tr.append('<w:tblHeader/>')
        tr.append('</w:trPr>')
        for i, h in enumerate(header):
            tr.append(self._cell(h, cols[i], bold=True, fill=header_fill,
                                 size=size, align='center', valign='center'))
        tr.append('</w:tr>')
        out.append(''.join(tr))

        for ri, row in enumerate(rows):
            fill = zebra if (zebra and ri % 2 == 1) else None
            tr = ['<w:tr>']
            for i in range(n):
                v = row[i] if i < len(row) else ''
                tr.append(self._cell(v, cols[i], size=size, fill=fill))
            tr.append('</w:tr>')
            out.append(''.join(tr))

        out.append('</w:tbl>')
        self._body.append(''.join(out))
        self._body.append('<w:p>%s</w:p>' % self._ppr(space_after=0))
        return self

    # ── 输出 ────────────────────────────────────────────────────────────
    def _sectpr(self) -> str:
        if self.orientation == 'landscape':
            pg = '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
            mar = ('<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" '
                   'w:left="1134" w:header="851" w:footer="992" w:gutter="0"/>')
        else:
            pg = '<w:pgSz w:w="11906" w:h="16838"/>'
            mar = ('<w:pgMar w:top="1440" w:right="1418" w:bottom="1440" '
                   'w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>')
        return ('<w:sectPr>%s%s<w:cols w:space="425"/>'
                '<w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>' % (pg, mar))

    def document_xml(self) -> bytes:
        body = list(self._body)
        if not body:
            body.append('<w:p/>')
        body.append(self._sectpr())
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
               '<w:document xmlns:w="%s" xmlns:r="%s">'
               '<w:body>%s</w:body></w:document>'
               % (NS_W, NS_R, ''.join(body)))
        return xml.encode('utf-8')

    def save(self, path: str, title: str = '', author: str = '') -> str:
        parts = {
            '[Content_Types].xml': self._content_types(),
            '_rels/.rels': self._rels(),
            'word/document.xml': self.document_xml().decode('utf-8'),
            'word/_rels/document.xml.rels': self._doc_rels(),
            'word/styles.xml': self._styles(),
            'docProps/core.xml': self._core(title, author),
            'docProps/app.xml': self._app(),
        }
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            for name, data in parts.items():
                z.writestr(name, data)
        return path

    # ── 包结构 ──────────────────────────────────────────────────────────
    def _content_types(self) -> str:
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Types xmlns="%s">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                '</Types>' % NS_CT)

    def _rels(self) -> str:
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="%s">'
                '<Relationship Id="rId1" Type="%s/officeDocument" Target="word/document.xml"/>'
                '<Relationship Id="rId2" Type="%s/metadata/core-properties" Target="docProps/core.xml"/>'
                '<Relationship Id="rId3" Type="%s/extended-properties" Target="docProps/app.xml"/>'
                '</Relationships>' % (NS_PKG, NS_R, NS_PKG, NS_R))

    def _doc_rels(self) -> str:
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="%s">'
                '<Relationship Id="rId1" Type="%s/styles" Target="styles.xml"/>'
                '</Relationships>' % (NS_PKG, NS_R))

    def _styles(self) -> str:
        f = _esc(self.font)
        sz = int(round(self.size * 2))
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:styles xmlns:w="%s">'
                '<w:docDefaults><w:rPrDefault><w:rPr>'
                '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>'
                '<w:sz w:val="%d"/><w:szCs w:val="%d"/>'
                '</w:rPr></w:rPrDefault>'
                '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="288" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
                '</w:docDefaults>'
                '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
                '<w:name w:val="Normal"/><w:qFormat/></w:style>'
                '<w:style w:type="table" w:default="1" w:styleId="TableNormal">'
                '<w:name w:val="Normal Table"/></w:style>'
                '</w:styles>' % (NS_W, f, f, f, f, sz, sz))

    def _core(self, title: str, author: str) -> str:
        import datetime
        now = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<cp:coreProperties xmlns:cp="%s" xmlns:dc="%s" xmlns:dcterms="%s" xmlns:xsi="%s">'
                '<dc:title>%s</dc:title><dc:creator>%s</dc:creator>'
                '<cp:lastModifiedBy>%s</cp:lastModifiedBy>'
                '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
                '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
                '</cp:coreProperties>' % (NS_CP, NS_DC, NS_DCTERMS, NS_XSI,
                                          _esc(title), _esc(author), _esc(author), now, now))

    def _app(self) -> str:
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Properties xmlns="%s"><Application>docx-dedup-aigc-pipeline</Application>'
                '</Properties>' % NS_EP)


if __name__ == '__main__':
    b = DocxBuilder(orientation='landscape')
    b.heading('DocxBuilder 自检', 1)
    b.paragraph('这是一个从零构造的 docx，用来验证写入器是否可用。')
    b.table(['列 A', '列 B'], [['1', '2'], ['3', '4']], widths=[0.5, 0.5])
    b.save('_writer_selftest.docx', title='自检')
    print('已生成 _writer_selftest.docx')
