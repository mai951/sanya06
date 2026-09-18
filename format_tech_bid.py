# -*- coding: utf-8 -*-
"""
Adjust generated technical bid docx to match 投标须知及投标文件格式.docx style:
- All text black (RGB 0,0,0) except citation runs (containing '（依据：') set to blue (0,0,255)
- Add statement "本技术标符合国家现行相关标准及规范要求。" after 编制说明 section
- Ensure cover page matches 技术标封面 (already similar)
"""
import sys
from docx import Document
from docx.shared import RGBColor

sys.stdout.reconfigure(encoding='utf-8')

IN_PATH = r'D:\罗\2026\三亚\001\三亚06地块商业改造外装综合幕墙工程_技术标_施工组织设计_v1.docx'
OUT_PATH = r'D:\罗\2026\三亚\001\三亚06地块商业改造外装综合幕墙工程_技术标_施工组织设计_v2.docx'

doc = Document(IN_PATH)

# Helper to set run color
def set_black(run):
    run.font.color.rgb = RGBColor(0,0,0)

def set_blue(run):
    run.font.color.rgb = RGBColor(0,0,255)

# First, set all runs to black
for para in doc.paragraphs:
    for run in para.runs:
        set_black(run)

# Then, for runs that look like citations, set to blue
for para in doc.paragraphs:
    full = para.text
    if '（依据：' in full:
        # We need to apply blue only to the citation part? Simpler: set whole para blue if contains citation
        for run in para.runs:
            set_blue(run)

# Add compliance statement after 编制说明
# Find paragraph with text '编制说明' (heading)
insert_idx = None
for i, para in enumerate(doc.paragraphs):
    if para.text.strip() == '编制说明':
        insert_idx = i
        break

if insert_idx is not None:
    # Insert after the heading and its following content until next heading or blank?
    # Simpler: insert a new paragraph right after the heading paragraph
    new_para = doc.paragraphs[insert_idx].insert_paragraph_after('本技术标符合国家现行相关标准及规范要求。')
    # Set its font to black (default) and maybe italic? Keep black.
    for run in new_para.runs:
        run.font.color.rgb = RGBColor(0,0,0)

# Ensure no other stray blue by rescanning and resetting non-citation to black
for para in doc.paragraphs:
    has_citation = '（依据：' in para.text
    for run in para.runs:
        if has_citation:
            # keep blue
            continue
        else:
            set_black(run)

doc.save(OUT_PATH)
print(f'Saved styled document to {OUT_PATH}')