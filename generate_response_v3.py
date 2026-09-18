# -*- coding: utf-8 -*-
"""
Generate response document based on the updated technical bid (技术标_5.docx)
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

def main():
    doc = Document()

    # Set default style
    style = doc.styles['Normal']
    style.font.name = '宋体'
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.color.rgb = RGBColor(0, 0, 0)

    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('三亚06地块商业改造外装综合幕墙工程技术标施工组织设计编制要求响应内容')
    run.font.size = Pt(16)
    run.font.name = '黑体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    run.bold = True
    doc.add_paragraph()  # empty line

    # Extract content from the technical bid
    tech_doc = Document(r'D:\罗\2026\三亚\001\三亚06地块商业改造外装综合幕墙工程_技术标_5.docx')
    tech_paras = [p.text for p in tech_doc.paragraphs if p.text.strip()]
    tech_full = '\n'.join(tech_paras)

    # Section 1: 履行现场管理职责
    doc.add_heading('履行现场管理职责', level=1)
    p = doc.add_paragraph()
    run = p.add_run('在总平面图上对场地进行围挡、门岗、材料进出场通道、人员通道、电梯管理、非机动车停放、材料库房、加工房等区域进行合理划分并悬挂张贴相应的管理制度。有临时用电布置详细图纸和管理措施。对甲供材、自身单位进行协调及管理的相关措施。')
    doc.add_paragraph()

    # Section 2: 现场管理及人员配备分工安排
    doc.add_heading('现场管理及人员配备分工安排', level=1)
    p = doc.add_paragraph()
    run = p.add_run('对工地内各单位（外装/景观/土建/内装）的现场管理进行详细阐述，明确人员配备分工安排。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('履行政府要求的一系列监管举措，包括安全文明施工、环境保护等。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('体现现场垂直/水平运输动线/施工动线，制定专项运输方案。')
    doc.add_paragraph()

    # Section 3: 施工工期计划
    doc.add_heading('施工工期计划', level=1)
    p = doc.add_paragraph()
    run = p.add_run('响应工期要求（总工期212日历天，临街位置2027年1月30日前完成），以project编制进度计划，各工序节点的先后逻辑关联正确清楚，各甲供材、甲分包节点时间清楚有可执行性。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('需体现专项的各楼层施工计划、公区、大堂、露台区域施工计划。')
    doc.add_paragraph()

    # Section 4: 施工难点、重点及控制措施
    doc.add_heading('施工难点、重点及控制措施', level=1)
    difficulties = [
        '工期刚性约束与海南气候条件叠加影响：总工期仅212天，临街节点仅120天；海南台风、高温雨季影响施工。控制措施：综合考虑雨季及台风施工，制定专项施工措施；材料进场提前综合考虑，部分材料提前订货；编制分段流水施工计划。',
        '改造工程拆除与基层复核先行：原有结构不满足规范要求导致额外处理。控制措施：进场即进行测量放线与主体结构复核，形成放线图报招标人复核；发现结构尺寸偏差超过30mm立即上报招标人确认调整方案后再下料。',
        '系统类型多、工艺跨度大：深化设计工作量大，易漏项。控制措施：按清单项目特征完成各系统深化，二次深化图纸确认后必须与招标人签字确认。优化设计仅限于结构构件，所有针对招标图纸优化设计必须在技术标中明确列项说明并附相关图纸。',
        '成品保护与原有系统搭接责任重：改造需保护性拆除且需为全专业兜底收口。控制措施：建立全周期成品保护与工序移交制度。施工区域及原外装搭接部位，严格落实“软+硬双层保护”。各批次完工后，牵头各分包以书面工序移交单确认保护责任，相关费用已综合考虑在措施费中。',
        '淋水试验为强制性工序：渗水导致整改及经济责任。控制措施：严格执行二次淋水试验（塞缝一次、挂管一次），参照万华集团外装淋水检验操作指引；淋水记录留痕作为隐蔽资料。'
    ]
    for diff in difficulties:
        p = doc.add_paragraph(style='List Bullet')
        run = p.add_run(diff)

    # Section 5: 玻璃幕墙系统、铝板幕墙系统、水泥纤维板外墙等与场馆内的功能与检修
    doc.add_heading('玻璃幕墙系统、铝板幕墙系统、水泥纤维板外墙等与场馆内的功能与检修', level=1)
    p = doc.add_paragraph()
    run = p.add_run('玻璃幕墙系统：全玻座装安装，座槽采用Q235热浸镀锌U型槽。嵌缝采用硅酮耐候胶室外嵌缝打胶+三元乙丙胶条。防水处理：幕墙洞口采用单组份聚合物水泥防水砂浆塞缝处理。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('铝板幕墙系统：铝单板局部弯折造型，采用专项模具加工。龙骨防腐处理到位，基层防锈处理满足设计要求。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('水泥纤维板外墙：板间留缝按规范处理，嵌缝采用硅酮耐候胶室外嵌缝打胶。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('铝合金防雨百叶：百叶安装前定位到位，安装后进行二次喷涂处理。角码转接件严格执行节点要求。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('精致钢龙骨造型雨棚：焊接工艺专项评定，焊工持证上岗。表面氟碳喷漆均匀，转接件等节点详细做法见节点图。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('防护栏杆与不锈钢水沟：栏杆立柱埋设深度满足要求，水沟坡度达标确保排水畅通。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('后置埋件：埋件预埋深度及方位严格按图，混凝土浇筑振实后二次捣实。螺栓扭矩达到设计要求。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('幕墙清洗：采用蜘蛛人或登高车方式，清洗剂不腐蚀幕墙材料。')
    doc.add_paragraph()

    # Section 6: 材料计划部署安排
    doc.add_heading('材料计划部署安排', level=1)
    p = doc.add_paragraph()
    run = p.add_run('对主要材料的采购有切实可行的实施计划方案，有专门的材料计划时间表（含甲供材）。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('幕墙铝型材、玻璃、五金等主要材料采购时间表，确保与施工进度匹配。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('木饰面加工进度保证。（如适用）')
    doc.add_paragraph()

    # Section 7: 质量控制措施或保证体系
    doc.add_heading('质量控制措施或保证体系', level=1)
    p = doc.add_paragraph()
    run = p.add_run('对石材有系统性的排版、下单的技术控制措施和石材防病变专项措施。（如适用）')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('有木饰面质量、细部品质控制措施，并写明基材品牌、规格，及油漆工艺，下单准确性保证措施。（如适用）')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('配图文介绍不锈钢加工、施工品质控制措施。（如适用）')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('有瓷砖、石材防空鼓专项授程。（如适用）')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('配图文对卫生间防水编制专项施工方案，从基层处理、材料准备、施工、成品保护、检验批及干湿分区返坎及卫生间门口防潮进行详细介绍并确保落实。（如适用）')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('施工过程中定期自检、整改落实、复检等具体质量控制实施所需的制度、流程、表格、人员配置等。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('木饰面精细度控制，收边收口质量控制。（如适用）')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('不锈钢-构件类不锈钢拼接控制，不锈钢门包封。（如适用）')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('艺术涂料/金属漆质量控制。（如适用）')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('镜面类不锈钢平整度控制。（如适用）')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('木地板铺设的平整度，拼接缝隙均匀等质量控制。（如适用）')
    doc.add_paragraph()

    # Section 8: 现场深化图纸能力
    doc.add_heading('现场深化图纸能力', level=1)
    p = doc.add_paragraph()
    run = p.add_run('有稳定的驻场设计人员对图纸进行深化，并对深化内容先后顺序及程度进行详细描述。')
    doc.add_paragraph()

    # Section 9: 深化设计项目履历
    doc.add_heading('深化设计项目履历', level=1)
    p = doc.add_paragraph()
    run = p.add_run('（可提供过往深化设计项目的履历证明）')
    doc.add_paragraph()

    # Section 10: 项目部管理架构组织
    doc.add_heading('项目部管理架构组织', level=1)
    p = doc.add_paragraph()
    run = p.add_run('人员配置满足技术文件内要求，招标单位在技术标答辩阶段将约谈面试主要负责人是否满足要求。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('项目经理：1人，驻场，在职期限10个月，全职。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('安全专员（兼资料员）：1人，驻场，在职期限7个月，全职。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('商务经理：月度三单核对及结算期间驻场，其他时间无硬性要求，非全职。')
    doc.add_paragraph()

    # Section 11: 安全文明生产目标及保证授程
    doc.add_heading('安全文明生产目标及保证授程', level=1)
    p = doc.add_paragraph()
    run = p.add_run('有详细的管理授程和具体的目标，以及切实可行的实施方案，如库房（危险品单独设立库房/堆放）、临水临电、材料堆码、工完场清制度、焊接用火管理制度、禁烟管理制度等，包括对各甲分单位的安全文明施工管理授程制度。')
    doc.add_paragraph()

    # Section 12: 拟投入劳动力计划
    doc.add_heading('拟投入劳动力计划', level=1)
    p = doc.add_paragraph()
    run = p.add_run('能确保本项目在各阶段有充分的劳动力，满足项目质量进度需要。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('以表格形式体现')
    doc.add_paragraph()

    # Section 13: 拟投入机械计划
    doc.add_heading('拟投入机械计划', level=1)
    p = doc.add_paragraph()
    run = p.add_run('能确保本项目在各阶段有充分的机具；垂直运输/室内电梯运力测算及保障计划。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('脚手架搭设：双排脚手架搭设3个月。')
    doc.add_paragraph()

    # Section 14: 资金保证
    doc.add_heading('资金保证', level=1)
    p = doc.add_paragraph()
    run = p.add_run('若中标后中标单位为保证项目的正常进展所需的资金保证措施。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('附匹配项目施工节奏的资金计划（垫资能力及拟定逐月完成产值）')
    doc.add_paragraph()

    # Section 15: 成品保护措施
    doc.add_heading('成品保护措施', level=1)
    p = doc.add_paragraph()
    run = p.add_run('必需图文并茂并针对本项目所有环节进行详细的做法、维护、交接手续等方面进行阐述。以及对各甲分包单位的成品进行管理措施。')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('施工区域及原外装搭接部位，严格落实“软+硬双层保护”。各批次完工后，牵头各分包以书面工序移交单确认保护责任。')
    doc.add_paragraph()

    # Section 16: 交付运营期间的配合工作
    doc.add_heading('交付运营期间的配合工作', level=1)
    p = doc.add_paragraph()
    run = p.add_run('有详细的实施授程、制度和充足的人员、机具、材料安排。')
    doc.add_paragraph()

    # Section 17: 售后服务体系
    doc.add_heading('售后服务体系', level=1)
    p = doc.add_paragraph()
    run = p.add_run('有详细的实施授程、制度和充足的人员、机具、材料安排。')
    doc.add_paragraph()

    # Section 18: 打样工作
    doc.add_heading('打样工作', level=1)
    p = doc.add_paragraph()
    run = p.add_run('中标单位进场后，须先对项目重难点及指定特殊工艺开展工艺打样工作，指定特殊工艺及打样要求如下：')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('①玻璃幕墙节点打样；')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('②铝板幕墙弯折节点打样；')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('③百叶安装节点打样；')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('④雨棚节点打样；')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('⑤后置埋件节点打样；')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('⑥门窗及电动装置节点打样；')
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run('打样完成时限：自进场之日起30天以内。')
    doc.add_paragraph()

    # Save the document
    output_path = r'D:\罗\2026\三亚\001\三亚06地块商业改造外装综合幕墙工程_技术标_施工组织设计_响应内容_v4.docx'
    doc.save(output_path)
    print(f'Response document saved to {output_path}')

if __name__ == '__main__':
    main()