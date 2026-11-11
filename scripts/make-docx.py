# 기획서 Markdown → Word(docx):  python3 scripts/make-docx.py
# 제목·소제목·표·글머리·번호 목록·인용·**굵게**·`코드` 만 다룹니다. mermaid 그림은 글자 그대로 고정폭으로 넣습니다.
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / 'docs' / '01_프로젝트_기획서.md'
OUT = ROOT / 'docs' / '01_프로젝트_기획서.docx'
FONT = '맑은 고딕'


def set_font(run, size=None, bold=None, mono=False, color=None):
    name = 'Consolas' if mono else FONT
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def add_inline(par, text, size=None):
    # **굵게** 와 `코드` 를 나눠 넣습니다
    for part in re.split(r'(\*\*[^*]+\*\*|`[^`]+`)', text):
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            set_font(par.add_run(part[2:-2]), size, bold=True)
        elif part.startswith('`') and part.endswith('`'):
            set_font(par.add_run(part[1:-1]), size, mono=True)
        else:
            set_font(par.add_run(part), size)


def cells(line):
    return [c.strip() for c in line.strip().strip('|').split('|')]


def main():
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = FONT
    st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    st.font.size = Pt(10)
    for s in ('Title', 'Heading 1', 'Heading 2', 'Heading 3', 'List Bullet', 'List Number'):
        doc.styles[s].font.name = FONT
        doc.styles[s].element.rPr.rFonts.set(qn('w:eastAsia'), FONT)

    lines = SRC.read_text(encoding='utf-8').split('\n')
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.startswith('```'):
            j = i + 1
            block = []
            while j < len(lines) and not lines[j].startswith('```'):
                block.append(lines[j])
                j += 1
            p = doc.add_paragraph()
            set_font(p.add_run('(흐름도 — Markdown 판에서 그림으로 보입니다)'), 9, color=(0x56, 0x61, 0x6F))
            for b in block:
                q = doc.add_paragraph()
                q.paragraph_format.space_after = Pt(0)
                set_font(q.add_run(b.replace('<br/>', ' ')), 8, mono=True)
            i = j + 1
            continue
        if line.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].startswith('|'):
                if not re.match(r'^\|\s*-', lines[i]):
                    rows.append(cells(lines[i]))
                i += 1
            ncol = max(len(r) for r in rows)
            t = doc.add_table(rows=len(rows), cols=ncol)
            t.style = 'Table Grid'
            for r, row in enumerate(rows):
                for c in range(ncol):
                    cell = t.cell(r, c)
                    cell.text = ''
                    add_inline(cell.paragraphs[0], row[c] if c < len(row) else '', 9)
                    if r == 0:
                        for run in cell.paragraphs[0].runs:
                            run.bold = True
            doc.add_paragraph()
            continue
        if line.startswith('# '):
            p = doc.add_paragraph(style='Title')
            set_font(p.add_run(line[2:].strip()), 20, bold=True)
        elif line.startswith('### '):
            p = doc.add_heading(level=2)
            set_font(p.add_run(line[4:].strip()), 12, bold=True)
        elif line.startswith('## '):
            p = doc.add_heading(level=1)
            set_font(p.add_run(line[3:].strip()), 14, bold=True)
        elif line.strip() == '---' or not line.strip():
            pass
        elif line.startswith('> '):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(12)
            add_inline(p, line[2:])
        elif re.match(r'^\s*- ', line):
            p = doc.add_paragraph(style='List Bullet')
            if line.startswith('  '):
                p.paragraph_format.left_indent = Pt(36)
            add_inline(p, re.sub(r'^\s*- ', '', line))
        elif re.match(r'^\s*\d+\. ', line):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(24 if line.startswith('  ') else 12)
            add_inline(p, line.strip())
        else:
            add_inline(doc.add_paragraph(), line)
        i += 1
    doc.save(OUT)
    print('저장:', OUT.relative_to(ROOT))


if __name__ == '__main__':
    sys.exit(main())
