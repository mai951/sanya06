# -*- coding: utf-8 -*-
"""
Rewrite technical bid to meet user requirements:
- Active voice, solution-oriented
- Chapter 1: Risk-Consequence-Our Approach
- Chapter 3: Focus on construction methods, remove calculation rules
- Chapter 4: Add specific implementation measures
- Throughout: Use "我们" perspective
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn
import re

def rewrite_chapter1_重难点(doc):
    """Rewrite chapter 1.4 to risk-consequence-our approach format"""
    # This will be done by reconstructing the chapter

def extract_current_content(doc_path):
    """Extract content from current document"""
    doc = Document(doc_path)
    content = []
    for p in doc.paragraphs:
        if p.text.strip():
            content.append({
                'text': p.text.strip(),
                'style': p.style.name if p.style else 'Normal'
            })
    return content

def is_heading(text):
    """Check if text is a heading"""
    heading_patterns = [
        r'^一、', r'^二、', r'^三、', r'^四、', r'^五、', r'^六、', r'^七、', r'^八、', r'^九、',
        r'^第[一二三四五六七八九十]+章',
        r'^[0-9]+\.[0-9]+',
        r'^编制说明$',
        r'^一、编制依据$',
        r'^附表'
    ]
    return any(re.match(pattern, text) for pattern in heading_patterns)

def rewrite_content_based_on_sections(old_content):
    """Rewrite content based on user requirements"""
    new_content = []

    i = 0
    while i < len(old_content):
        item = old_content[i]
        text = item['text']

        # Handle cover page and basic info (keep as is for now)
        if i < 10:
            new_content.append(item)
            i += 1
            continue

        # Handle 编制说明 section
        if '编制说明' in text and len(text) < 10:
            new_content.append(item)
            # Add compliance statement after 编制说明
            new_content.append({'text': '本技术标符合国家现行相关标准及规范要求及招标依据编制。', 'style': 'Normal'})
            i += 1
            continue

        # Handle 第一章  工程概况与编制范围
        if '第一章  工程概况与编制范围' in text:
            new_content.append({'text': '第一章  工程概况与编制范围', 'style': 'Heading 1'})
            # Add subsections
            new_content.append({'text': '1.1 工程基本信息', 'style': 'Heading 2'})
            new_content.append({'text': '工程名称：三亚06地块商业改造外装综合幕墙工程。', 'style': 'Normal'})
            new_content.append({'text': '工程地点：海南省三亚市铁轨东路与康庄路交汇处。', 'style': 'Normal'})
            new_content.append({'text': '招标人：海南万华商业管理有限公司。', 'style': 'Normal'})
            new_content.append({'text': '工程内容：外装玻璃幕墙、铝板幕墙、外墙涂料、雨蓬、栏杆、地弹门、电动平移门、氟碳漆、铝合金百叶、石材反坎、不锈钢水槽、后置埋件等相关幕墙材料采购及安装，以及成品保护、幕墙保洁、淋水试验等。', 'style': 'Normal'})
            new_content.append({'text': '现场状况：三亚06地块大区工程幕墙、景观地面已全部施工完成，属商业改造项目。', 'style': 'Normal'})
            new_content.append({'text': '1.2 工期要求', 'style': 'Heading 2'})
            new_content.append({'text': '总工期212日历天，预计开工时间2026年10月1日，竣工时间2027年4月30日，正式开街2027年10月1日。', 'style': 'Normal'})
            new_content.append({'text': '临街位置要求2027年1月30日前完成。', 'style': 'Normal'})
            new_content.append({'text': '总工期和各阶段工期的起算时间以招标人书面通知的现场开工令为准。', 'style': 'Normal'})
            new_content.append({'text': '1.3 承包范围', 'style': 'Heading 2'})
            new_content.append({'text': '施工范围：三亚06地块商业改造外装综合幕墙工程包括但不限于：图纸上所有（不含甲供）外装图纸中包括的玻璃幕墙、铝板幕墙、外墙涂料、雨蓬、栏杆、地弹门、电动平移门、氟碳漆、铝合金百叶、石材反坎、不锈钢水槽、后置埋件等幕墙材料采购、安装工程，以及成品保护、幕墙保洁、淋水试验等，包含现场实际交叉产生的部位处理施工、脚手架，以及所有考虑措施实施的条件（含高空作业、防台风、防雨等措施）。', 'style': 'Normal'})
            new_content.append({'text': '承包范围包含：深化设计（对招标图纸的深化）、材料采购、加工制作、运输、安装、检验试验、成品保护、幕墙保洁、淋水试验、竣工验收、质保期维修，以及与其他专业（土建、内装、机电、景观）的界面配合。', 'style': 'Normal'})
            # Rewrite 1.4 with risk-consequence-our approach
            new_content.append({'text': '1.4 工程重点与难点识别', 'style': 'Heading 2'})
            new_content.append({'text': '重点难点一：工期刚性约束强。风险：总工期仅212天，临街节点仅120天。后果：延误将导致违约金及声誉损失。我方对策：编制分阶段流水施工计划，临街段资源倾斜，采用分段施工+资源前置，关键路径压缩。', 'style': 'Normal'})
            new_content.append({'text': '重点难点二：改造工程拆除与基层复核先行。风险：原有结构不满足规范要求导致额外处理。后果：施工中发现隐蔽工程增加费用及工期。我方对策：进场即进行测量放线与主体结构复核，形成放线图报招标人复核；发现结构尺寸偏差超过30mm立即上报招标人确认调整方案后再下料。', 'style': 'Normal'})
            new_content.append({'text': '重点难点三：系统类型多、工艺跨度大。风险：深化设计工作量大，易漏项。后果：材料加工周期不匹配导致工期延误。我方对策：按清单项目特征完成各系统深化，二次深化图纸确认后必须与招标人签字确认。优化设计仅限于结构构件，所有针对招标图纸优化设计必须在技术标中明确列项说明并附相关图纸。', 'style': 'Normal'})
            new_content.append({'text': '重点难点四：成品保护与原有系统搭接责任重。风险：改造需保护性拆除且需为全专业兜底收口。后果：保护不到位导致返工及索赔。我方对策：建立全周期成品保护与工序移交制度。施工区域及原外装搭接部位，严格落实“一层软保护+一层硬保护”。各批次完工后，牵头各分包以书面工序移交单确认保护责任，相关费用已综合考虑在措施费中。', 'style': 'Normal'})
            new_content.append({'text': '重点难点五：淋水试验为强制性工序。风险：渗水导致整改及经济责任。后果：渗水点由我方承担全部返工责任并按2000元/点支付违约金。我方对策：严格执行二次淋水试验（塞缝一次、挂管一次），参照万华集团外装淋水检验操作指引；淋水记录留痕作为隐蔽资料。', 'style': 'Normal'})
            new_content.append({'text': '重点难点六：海南气候条件影响。风险：台风、高温雨季影响施工。后果：积水及高温天气减少有效施工时间。我方对策：综合考虑雨季及台风施工，制定专项施工措施；材料进场提前综合考虑，部分材料提前订货。', 'style': 'Normal'})
            i += 1
            continue

        # Handle 第二章  施工部署
        elif '第二章  施工部署' in text:
            new_content.append({'text': '第二章  施工部署', 'style': 'Heading 1'})
            # Keep most of this section but adjust perspective
            i += 1
            continue

        # Handle 第三章  主要分项工程施工方案 - THIS NEEDS MAJOR REWRITE
        elif '第三章  主要分项工程施工方案' in text:
            new_content.append({'text': '第三章  主要分项工程施工方案', 'style': 'Heading 1'})
            new_content.append({'text': '本章施工方案聚焦于施工工艺和技术控制要点，删除纯工程量计算规则和清单抄录内容。', 'style': 'Normal'})

            # Rewrite each subsection with focus on construction methods
            subsections = [
                ('3.1 拆除工程', '采用保护性拆除工艺。搭接部位人工配合小型机具精细拆除，保护原有结构及保留幕墙；非搭接部位机械拆除。拆除残值材料按铝材、铁钢材分类回收，抵扣工程款。'),
                ('3.2 结构封补与零星工程', '严格执行基层处理规范。混凝土浇筑振实密实，砂浆抹平饱满。零星工程按图集做法进行，确保与主体结构可靠连接。'),
                ('3.3 玻璃幕墙系统', '全玻座装安装，座槽采用Q235热浸镀锌U型槽。嵌缝采用硅酮耐候胶室外嵌缝打胶+三元乙丙胶条。防水处理：幕墙洞口采用单组份聚合物水泥防水砂浆塞缝处理。'),
                ('3.4 铝板幕墙系统', '铝单板局部弯折造型，采用专项模具加工。龙骨防腐处理到位，基层防锈处理满足设计要求。'),
                ('3.5 水泥纤维板外墙', '板间留缝按规范处理，嵌缝采用硅酮耐候胶室外嵌缝打胶。'),
                ('3.6 铝合金防雨百叶', '百叶安装前定位到位，安装后进行二次喷涂处理。角码转接件严格执行节点要求。'),
                ('3.7 精致钢龙骨造型雨棚', '焊接工艺专项评定，焊工持证上岗。表面氟碳喷漆均匀，转接件等节点详细做法见节点图。'),
                ('3.8 防护栏杆与不锈钢水沟', '栏杆立柱埋设深度满足要求，水沟坡度达标确保排水畅通。'),
                ('3.9 后置埋件', '埋件预埋深度及方位严格按图，混凝土浇筑振实后二次捣实。螺栓扭矩达到设计要求。'),
                ('3.10 门窗与电动装置', '门窗安装找正固定，四周填缝密封。电动门厂家深化，安全装置齐全。'),
                ('3.11 泛光工程', '灯具安装防水处理到位，线管敷设符合要求。'),
                ('3.12 幕墙清洗', '采用蜘蛛人或登高车方式，清洗剂不腐蚀幕墙材料。')
            ]

            for title, method in subsections:
                new_content.append({'text': title, 'style': 'Heading 2'})
                new_content.append({'text': method, 'style': 'Normal'})
            i += 1
            # Skip original detailed content
            while i < len(old_content) and not is_heading(old_content[i]['text']):
                i += 1
            continue

        # Handle 第四章  关键施工措施 - NEED TO ADD SPECIFIC IMPLEMENTATION
        elif '第四章  关键施工措施' in text:
            new_content.append({'text': '第四章  关键施工措施', 'style': 'Heading 1'})
            # Keep the structure but add specific implementation
            i += 1
            continue

        # Handle other chapters - keep most but adjust perspective
        else:
            # Adjust perspective language
            adjusted_text = text
            # Replace passive expressions with active ones
            replacements = [
                ('请投标单位综合考虑', '我方已在方案中充分考量'),
                ('请投标单位严格按照', '我方将严格按照'),
                ('应投标单位', '我方'),
                ('投标人应当', '我方将'),
                ('需投标单位', '我方需要'),
                ('由投标单位负责', '由我方负责'),
                ('投标单位负责', '我方负责'),
                ('建议投标单位', '我方建议'),
                ('要求投标单位', '我方要求'),
                ('应由投标单位', '应由我方'),
                ('由投标单位自行', '由我方自行'),
                ('投标单位应', '我方应'),
                ('投标单位须', '我方须'),
                ('投标人须', '我方须'),
                ('投标人应', '我方应')
            ]

            for old, new in replacements:
                adjusted_text = adjusted_text.replace(old, new)

            # Remove excessive quoting and calculation rules
            # Remove lines that are purely calculation rules or pure quoting
            lines = adjusted_text.split('\n')
            filtered_lines = []
            for line in lines:
                line = line.strip()
                # Skip pure calculation rules
                if ('计算规则：' in line and len(line) < 50) or \
                   ('工程量按' in line and '计算' in line and len(line) < 100) or \
                   ('综合单价包含' in line and len(line) < 100) or \
                   (line.count('：') >= 2 and len(line) < 80 and '，' in line):  # Likely just quoting
                    continue
                filtered_lines.append(line)

            if filtered_lines:
                new_content.append({'text': '\n'.join(filtered_lines), 'style': 'Normal'})

            i += 1
            continue

    return new_content

def main():
    input_path = r'D:\罗\2026\三亚\001\三亚06地块商业改造外装综合幕墙工程_技术标_施工组织设计_v3.docx'
    output_path = r'D:\罗\2026\三亚\001\三亚06地块商业改造外装综合幕墙工程_技术标_施工组织设计_v4.docx'

    # Extract current content
    old_content = extract_current_content(input_path)

    # Rewrite content
    new_content = rewrite_content_based_on_sections(old_content)

    # Create new document
    new_doc = Document()

    # Set default style to match reference
    style = new_doc.styles['Normal']
    style.font.name = '宋体'
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.color.rgb = RGBColor(0, 0, 0)

    # Configure heading styles
    for level in range(1, 4):
        h_style = new_doc.styles[f'Heading {level}']
        h_style.font.name = '黑体'
        h_style._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        h_style.font.color.rgb = RGBColor(0, 0, 0)
        if level == 1:
            h_style.font.size = Pt(16)
            h_style.paragraph_format.space_before = Pt(12)
            h_style.paragraph_format.space_after = Pt(6)
        elif level == 2:
            h_style.font.size = Pt(14)
            h_style.paragraph_format.space_before = Pt(10)
            h_style.paragraph_format.space_after = Pt(4)
        else:
            h_style.font.size = Pt(12)
            h_style.paragraph_format.space_before = Pt(8)
            h_style.paragraph_format.space_after = Pt(4)

    # Add content
    for item in new_content:
        text = item['text']
        style_name = item.get('style', 'Normal')

        if style_name.startswith('Heading'):
            level = int(style_name.split()[-1])
            p = new_doc.add_heading(text, level=level)
        else:
            p = new_doc.add_paragraph()
            p.paragraph_format.first_line_indent = Cm(0.74)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(text)
            run.font.name = '宋体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0, 0, 0)

    # Save
    new_doc.save(output_path)
    print(f'Rewritten document saved to {output_path}')

if __name__ == '__main__':
    from docx.shared import Pt, RGBColor, Cm
    main()