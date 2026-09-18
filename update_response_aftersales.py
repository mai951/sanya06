# -*- coding: utf-8 -*-
"""
Update the response document with enhanced after-sales service section based on contract warranty terms.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

def main():
    input_path = r'D:\罗\2026\三亚\001\三亚06地块商业改造外装综合幕墙工程_技术标_施工组织设计_响应内容.docx'
    output_path = r'D:\罗\2026\三亚\001\三亚06地块商业改造外装综合幕墙工程_技术标_施工组织设计_响应内容_v2.docx'

    doc = Document(input_path)

    # Find and replace the "售后服务体系" section
    new_paragraphs = []
    i = 0
    while i < len(doc.paragraphs):
        p = doc.paragraphs[i]
        text = p.text.strip()
        if text == '售后服务体系':
            # Skip this and next few paragraphs, replace with enhanced version
            # Collect current paragraphs to skip
            new_paragraphs.append(p)  # Keep the heading
            i += 1
            # Skip old content until next heading
            while i < len(doc.paragraphs):
                if doc.paragraphs[i].text.strip() and doc.paragraphs[i].style.name.startswith('Heading'):
                    break
                i += 1
            # Now insert enhanced content
            continue
        new_paragraphs.append(p)
        i += 1

    # Create new document with enhanced content
    new_doc = Document()

    # Set default style
    style = new_doc.styles['Normal']
    style.font.name = '宋体'
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.color.rgb = RGBColor(0, 0, 0)

    # Copy all content, but enhance the after-sales section
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            new_doc.add_paragraph()
            continue

        # Check if this is the 售后服务体系 section
        if text == '售后服务体系':
            # Add heading
            h = new_doc.add_heading(text, level=1)
            # Enhanced content
            enhanced_items = [
                '我方承诺严格履行合同约定的维修及保修义务，建立完善的售后服务体系，确保工程交付后的质量保障与运营配合。',
                '',
                '一、保修期限与范围',
                '1. 保修期自全部工程竣工验收达到合格标准并交付发包人后，从发包人集中交房期满之日起算，质保期详见本合同附件《施工合同保修维修条款》的约定。',
                '2. 保修范围覆盖本工程所有外装系统：玻璃幕墙、铝板幕墙、水泥纤维板外墙、铝合金百叶、精致钢龙骨雨棚、防护栏杆、不锈钢水沟、后置埋件、门窗及电动装置、泛光工程等。',
                '3. 保修内容包括但不限于：渗漏水维修、幕墙五金更换、玻璃更换、铝板变形修复、嵌缝胶老化更换、表面涂层修补、百叶/门窗运行故障维修、埋件锚固检查等。',
                '',
                '二、维修响应机制',
                '1. 接到维修通知后，无论质量缺陷属承包人、发包人或业主责任，我方均遵守时间性要求不问理由地进行维修。',
                '2. 一般质量问题：接到通知后 2 小时内响应，24 小时内到达现场，48 小时内完成维修。',
                '3. 紧急渗漏/安全隐患：接到通知后 1 小时内响应，4 小时内到达现场，实施应急处置。',
                '4. 维修完毕后，负责将施工现场清理干净，取得发包人验收签字，经发包人签字确认方视为完成维修义务。',
                '5. 所维修项目如在六个月内再次出现同样或类似缺陷的，无论是否在质保期内仍由我方维修。',
                '',
                '三、责任界定与费用承担',
                '1. 因我方工程质量问题造成发包人和业主的全部直接损失（包括退房、赔偿等），由我方无条件承担。',
                '2. 因质量问题引发的业主或第三方索赔，授权发包人全权代表我方与业主或第三方进行索赔谈判并确定赔偿金额，结果经签字后对我方即刻生效。',
                '3. 维修费用由发包人通过《扣款通知单》从应付款项中直接扣除，不足部分我方另行补足。',
                '4. 对于涉及结构安全的质量问题，按《房屋建设工程质量保修办法》规定立即向建设行政主管部门报告，由原设计单位或有资质设计单位提出保修方案，我方实施并承担全部费用。',
                '',
                '四、终身维修承诺',
                '1. 承诺在工程保修期满后，对该工程进行终身维修，维修时只收工料成本费。',
                '2. 如我方不履行维修义务，发包人有权聘请其它专业公司进行维修，由此多付出的费用由我方负责赔偿。',
                '',
                '五、维修服务制度与保障',
                '1. 设立专门售后服务小组：配备项目经理 1 人（总负责）、维修技术员 2 人（含幕墙/五金/电气专业）、资料员 1 人（维修档案管理）、巡检员 2 人。',
                '2. 建立维修档案管理制度：每次维修建立《维修记录表》，记录缺陷描述、处理措施、用料规格、验收签字、照片影像，归档备查。',
                '3. 定期巡检制度：保修期内每季度开展 1 次全面巡检，重点检查渗漏、五金运行、嵌缝胶老化、表面涂层剥落等情况，出具《季度巡检报告》报发包人。',
                '4. 应急物资储备：项目部常备常用维修物资（结构胶、密封胶、五金配件、玻璃吸盘、注胶枪、测漏仪等），确保 4 小时内应急处置。',
                '5. 维修配合制度：配合发包人、物业管理单位的检查验收，提供必要的检测工具、人员陪同，签字确认检查结果。',
                '',
                '六、交付运营期间配合工作',
                '1. 销售包装配合：按发包人规定时间完成有关工作（销售配合时间由发包人提前通知）。',
                '2. 现场开放配合：开放日前对施工现场进行清理，作好相应安全措施。',
                '3. 提前施工配合：为配合销售而提前需要施工的项目，按发包人确认工期安排施工，费用双方协商。',
                '',
                '七、质保金管理',
                '1. 质保金按结算总价 3% 划扣，质保金发票需在支付结算款时提供。',
                '2. 发包人退还质保金截止时间：从发包人集中交房期满之日起算届满 2 年。',
                '3. 经监理工程师、物业管理单位及发包人对工程质量无异议，且共同书面确认无任何质量问题后 28 天内办理完成保修金结算。',
                '4. 应付剩余质量保修金 = 工程结算总价 3% 质保金 - 保修期内发包人签发的《扣款通知单》总额。'
            ]

            for item in enhanced_items:
                if item == '':
                    new_doc.add_paragraph()
                elif item.startswith('一、') or item.startswith('二、') or item.startswith('三、') or item.startswith('四、') or item.startswith('五、') or item.startswith('六、') or item.startswith('七、'):
                    new_doc.add_heading(item, level=2)
                else:
                    p = new_doc.add_paragraph()
                    p.paragraph_format.first_line_indent = Cm(0.74)
                    p.paragraph_format.space_after = Pt(2)
                    run = p.add_run(item)
                    run.font.name = '宋体'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                    run.font.size = Pt(10.5)
                    run.font.color.rgb = RGBColor(0, 0, 0)
            continue

        # Copy other paragraphs
        if p.style.name.startswith('Heading'):
            new_doc.add_heading(text, level=int(p.style.name.split()[-1]))
        else:
            new_p = new_doc.add_paragraph()
            new_p.paragraph_format.first_line_indent = Cm(0.74)
            new_p.paragraph_format.space_after = Pt(2)
            run = new_p.add_run(text)
            run.font.name = '宋体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0, 0, 0)

    new_doc.save(output_path)
    print(f'Enhanced document saved to {output_path}')

if __name__ == '__main__':
    main()