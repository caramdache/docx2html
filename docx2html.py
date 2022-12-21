from collections import defaultdict
from dataclasses import dataclass
from io import StringIO
import re

from docx import Document
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


# https://stackoverflow.com/questions/58324814/how-to-get-cell-background-color-in-python-docx
def get_fill_color(cell):
    # fill_color = cell._tc.get_or_add_tcPr().find(qn('w:shd'))

    pattern = re.compile('w:fill=\"(\S*)\"')
    match = pattern.search(cell._tc.xml)
    if match:
        return match.group(1)


@dataclass
class Range:
    rowspan: int = 0
    colspan: int = 0


class DocxHTMLGenerator:
    def __init__(self, path):
        self.doc = Document(path)
        self.out = StringIO()

    def write(self, string):
        self.out.write(string)

    def to_html(self):
        self.write('<table class="table table-bordered table-hover">\n<tbody>\n')

        for table in self.doc.tables:
            self.write("</tbody>\n</table>\n\n\n<table class='table table-bordered table-hover'>\n<tbody>\n")

            self.table_to_html(table, level=0)

        self.write('</tbody>\n</table>')

        return self.out.getvalue()

    def table_to_html(self, table, level):
        def get_spans():
            merged_cells = defaultdict(lambda: False)
            spans = defaultdict(lambda: Range())

            for i, row in enumerate(table.rows):
                last_cell = None
                last_j = 0
                for j, cell in enumerate(row.cells):
                    if cell == last_cell:
                        merged_cells[(i, j)] = True

                        span = spans[(i, last_j)]
                        span.colspan += 1
                        spans[(i, last_j)] = span

                    else:
                        last_cell = cell
                        last_j = j

            for j, column in enumerate(table.columns):
                last = None
                last_i = 0
                for i, cell in enumerate(column.cells):
                    if cell == last_cell:
                        merged_cells[(i, j)] = True

                        span = spans[(last_i, j)]
                        span.rowspan += 1
                        spans[(last_i, j)] = span

                    else:
                        last_cell = cell
                        last_i = i

            return merged_cells, spans

        def span_to_html():
            span = spans[(i, j)]

            colspan = f'colspan="{span.colspan + 1}"' if span.colspan else ''
            rowspan = f'rowspan="{span.rowspan + 1}"' if span.rowspan else ''

            return f'{colspan} {rowspan}'

        def style_to_html():
            s = f'{background_color_to_html()}{alignment_to_html()}'

            return f'style="{s}"' if s else ''

        def background_color_to_html():
            fill_color = get_fill_color(cell)
            
            if fill_color and fill_color != 'auto':
                return f'background-color:#{fill_color}'

            return ''

        def alignment_to_html():
            return ''

        def value_to_html(para):
            self.write('<p>')

            left_indent = para.paragraph_format.left_indent
            if left_indent:
                self.write('&nbsp;&nbsp;&nbsp;&nbsp;' * int(left_indent.pt // 18))

            if 'Heading'in para.style.name:
                self.write('<b>')            
            
            for run in para.runs:
                text = run.text
                color = run.font.color.rgb

                if run.bold:
                    self.write('<b>')
                if run.italic:
                    self.write('<i>')
                if run.underline:
                    self.write('<u>')
                if run.font.strike:
                    self.write('<strike>')
                if color:
                    self.write(f'<span style="color:#{str(color)}">')

                text = re.sub(r"^( )+", lambda m: '&nbsp;' * len(m.group(1)), text)
                text = re.sub(r"( {4,})", lambda m: '&nbsp;' * len(m.group(1)), text)
                text = text.replace('\t', '&nbsp;&nbsp;&nbsp;&nbsp;')
                text = text.replace('\n', '</p><p>')

                self.write(text) # TODO escape

                if color:
                    self.write(f'</span>')
                if run.font.strike:
                    self.write('</strike>')
                if run.underline:
                    self.write('</u>')
                if run.italic:
                    self.write('</i>')
                if run.bold:
                    self.write('</b>')

            if 'Heading'in para.style.name:
                self.write('</b>')            
    
            self.write('</p>')

        def iterchildren(cell):
            for child in cell._tc.iterchildren():
                if isinstance(child, CT_P):
                    yield Paragraph(child, cell)

                elif isinstance(child, CT_Tbl):
                    yield Table(child, cell)

        merged_cells, spans = get_spans()

        for i, row in enumerate(table.rows):
            self.write('<tr>')

            for j, cell in enumerate(row.cells):
                if merged_cells[(i, j)]:
                    continue

                self.write(f'<td {span_to_html()} {style_to_html()}>')

                for child in iterchildren(cell):
                    if isinstance(child, Paragraph):
                        value_to_html(child)

                    elif isinstance(child, Table):
                        self.write('<table>\n<tbody>')
                        self.table_to_html(child, level=level+1)
                        self.write('</tbody>\n</table>')

                self.write('</td>')

            self.write('</tr>')
