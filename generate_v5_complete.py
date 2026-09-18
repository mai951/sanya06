# -*- coding: utf-8 -*-
"""
Generate COMPLETE formatted technical bid v5.docx matching reference style
with ALL content from the reference technical bid.
"""
import sys
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.stdout.reconfigure(encoding='utf-8')

doc = Document()

# ---- Styles ----
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(10.5)
style._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
style.font.color.rgb = RGBColor(0, 0, 0)
style.paragraph_format.first_line_indent = Cm(0.74)
style.paragraph_format.space_after = Pt(3)
style.paragraph_format.space_before = Pt(0)
style.paragraph_format.line_spacing = Pt(20)

for level in range(1, 4):
    h = doc.styles[f'Heading {level}']
    h.font.name = 'Times New Roman'
    h.font.color.rgb = RGBColor(0, 0, 0)
    h._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    if level == 1:
        h.font.size = Pt(16)
        h.bold = True
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(6)
    elif level == 2:
        h.font.size = Pt(14)
        h.bold = True
        h.paragraph_format.space_before = Pt(10)
        h.paragraph_format.space_after = Pt(4)
    else:
        h.font.size = Pt(12)
        h.bold = True
        h.paragraph_format.space_before = Pt(8)
        h.paragraph_format.space_after = Pt(4)

def add_para(text, bold=False, indent=True):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0.74) if indent else Cm(0)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
    run.font.size = Pt(10.5)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.bold = bold
    return p

# ===== COVER PAGE =====
for _ in range(6):
    doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('三亚06地块商业改造外装工程')
run.font.size = Pt(22)
run.bold = True
run.font.name = 'Times New Roman'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
run.font.color.rgb = RGBColor(0, 0, 0)

doc.add_paragraph()
doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('投\n标\n文\n件')
run.font.size = Pt(26)
run.bold = True
run.font.name = 'Times New Roman'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
run.font.color.rgb = RGBColor(0, 0, 0)

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('【技术标】')
run.font.size = Pt(18)
run.bold = True
run.font.name = 'Times New Roman'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
run.font.color.rgb = RGBColor(0, 0, 0)

for _ in range(4):
    doc.add_paragraph()

info_lines = [
    '项 目 名 称：三亚06地块商业改造外装工程',
    '投   标  人：                    （加盖公章）',
    '法定代表人或其委托代理人：       （签字或盖章）',
    '日       期：         年         月        日',
]
for line in info_lines:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(line)
    run.font.size = Pt(10.5)
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
    run.font.color.rgb = RGBColor(0, 0, 0)

doc.add_page_break()

# ===== CONTENT - FULL CONTENT FROM REFERENCE =====

# 第一章 编制依据及原则
doc.add_heading('第一章  编制依据及原则', level=1)

doc.add_heading('1.1  编制依据', level=2)
add_para('国家现行相关技术、材料、验收等规范要求')
add_para('招标文件、施工图、材料表及技术指标要求')
add_para('现场踏勘资料及招标答疑澄清文件')

doc.add_heading('1.2  编制原则', level=2)
add_para('满足招标文件规定的工期、质量、安全要求')
add_para('结合改造工程特点与三亚气候特征，确保方案切实可行')
add_para('【待补充：企业自身相关标准或类似工程施工经验】')

# 第二章 工程概况与重难点分析
doc.add_heading('第二章  工程概况与重难点分析', level=1)

doc.add_heading('2.1  工程概况', level=2)
add_para('工程名称：三亚06地块商业改造外装综合幕墙工程')
add_para('工程内容：玻璃幕墙、铝板幕墙、外墙涂料、雨蓬、栏杆、地弹门、电动平移门、氟碳漆、铝合金百叶、石材反坎、不锈钢水槽、后置埋件等相关幕墙材料采购及安装，以及成品保护、幕墙拆改回收、垃圾外运、搭拆措施工作等')
add_para('现场状况：前期幕墙已全部施工完成，改造需保护性拆除，其余破坏性拆除，景观地面石材已完成，现场处于完工状态')
add_para('工期要求：暂定212日历天（2026.10.1 - 2027.4.30），临街位置2027.1.30前完成（仅120天），实际开工以开工令为准')
add_para('人员配置：项目经理、安全员（兼职资料员）全过程在岗，商务阶段调配，大面展开需增加管理人员')

doc.add_heading('2.2  重难点分析及对策', level=2)
add_para('重难点一：工期极其紧张（实际有效施工时间不足）', bold=True)
add_para('分析：总工期212天，但扣除台风（10-11月）、连续降雨、高温降效、整改排查、春节假期（2027年1月底），有效施工时间可能缩水至150天左右。临街120天节点实际安装时间可能仅60-70天')
add_para('对策：我方采用分段流水施工、春节抢工预案、增加人员设备、材料提前下单排产计划')

add_para('重难点二：改造工程拆除与结构偏差风险', bold=True)
add_para('分析：原幕墙已完工，需保护性拆除，其余破坏性拆除。结构尺寸偏差超过30mm需报甲方确认调整方案后方可下料')
add_para('对策：进场后首先复核主体结构，形成放线图。内外装图纸放样合图。原结构偏差处理方案提前报审')

add_para('重难点三：高标准工艺与防渗漏要求', bold=True)
add_para('分析：立柱横梁平齐交接需工厂高精度铣削+现场拼装（5mm胶缝）。防渗漏是红线，每做完一道防水必须淋水，飞检不合格必重罚')
add_para('对策：严格执行氟碳漆施工环境控制（10-30℃，湿度≤80%，7天固化）。淋水试验按万华指引两阶段执行（喷壶淋水+布管淋水）')

add_para('重难点四：界面划分与隐性成本全包干', bold=True)
add_para('分析：土建只提供基础条件，所有交接面防水、塞缝、打胶、收边、保护、垃圾外运、全域精保洁全部由外装包干。22条交底文件中，隐性成本不可在后期增加费用')
add_para('对策：投标时设立"专项措施费"科目，逐项测算摊入综合单价。履约阶段做好书面移交记录')

# 第三章 施工总体部署
doc.add_heading('第三章  施工总体部署', level=1)

doc.add_heading('3.1  施工目标', level=2)
add_para('质量目标：符合国家规范上限验收条件值')
add_para('安全目标：重伤、死亡人数为零，无重大机械设备、急性中毒事故，扬尘治理达标')
add_para('工期目标：确保212天总工期，临街位置2027.1.30前完成')

doc.add_heading('3.2  施工区段划分与流水段', level=2)
add_para('【待补充：结合A/B/C/D区及临街/非临街划分具体流水段，建议与深化设计图纸对应】')

doc.add_heading('3.3  施工顺序', level=2)
add_para('总体顺序：拆除及回收 → 测量放线及结构复核 → 深化设计与下料 → 后置埋件安装 → 龙骨安装 → 面板安装 → 塞缝防水及打胶 → 淋水试验 → 成品保护 → 精保洁')

doc.add_heading('3.4  施工准备与资源配置计划', level=2)
add_para('劳动力计划：【待补充：按施工阶段配置各工种人数】')
add_para('材料进场计划：【重点提示：玻璃、铝板定制加工周期45-60天，幕墙铝型材需提前下单排产】')

# 第四章 施工进度计划与保证措施
doc.add_heading('第四章  施工进度计划与保证措施', level=1)

doc.add_heading('4.1  进度计划', level=2)
add_para('【待补充：附横道图或网络计划图，须体现临街120天节点及总工期212天】')

doc.add_heading('4.2  工期保证措施', level=2)
add_para('台风及雨季应对：10-11月台风季，提前加固堆放材料，停止高空作业，建立应急响应机制')
add_para('高温降效应对：调整作息（做两头，歇中间），午间11点-15点停止露天作业')
add_para('春节抢工预案：提前储备劳动力，制定春节前赶工专项方案')
add_para('材料前置：中标后立即启动深化设计和主材下单，避免因材料延误导致工期滞后')

# 第五章 主要施工方案与技术措施
doc.add_heading('第五章  主要施工方案与技术措施', level=1)

doc.add_heading('5.1  测量放线及结构复核方案', level=2)
add_para('进场后首先复核主体结构，放出控制线形成放线图，经甲方复核后方可施工')
add_para('内外装图纸放样合图，结构尺寸偏差超过30mm的，报甲方确认调整方案后方可下料')

doc.add_heading('5.2  玻璃幕墙施工方案', level=2)
add_para('玻璃配置：隐框，加工厂施打双组份结构胶（指定品牌：安泰、白云、之江等），静电膜双面保护')
add_para('玻璃安装：安装前100%复检尺寸与外观，注胶饱满连续，无气泡断胶')
add_para('洞口处理：单组份聚合物水泥防水砂浆塞缝，防火位置做铁皮封堵')

doc.add_heading('5.3  氟碳漆施工方案', level=2)
add_para('基材：镀锌钢幕墙龙骨/镀锌构件')
add_para('配比：底漆4:1:0.3~0.5；氟碳面漆10:1:0.3~0.6')
add_para('工序：表面处理 → 喷涂环氧底漆（1-2遍，干膜30-40μm） → 中间处理 → 喷涂氟碳面漆（2遍，干膜40-60μm） → 干燥养护')
add_para('环境要求：温度10-30℃，湿度≤80%，雨天、雾天、大风严禁施工。完全固化7天，期间禁止磕碰、淋雨')

doc.add_heading('5.4  铝合金百叶施工方案', level=2)
add_para('涂层厚度：采用欧标（平均膜厚≥60μm，局部≥48μm）')
add_para('组装要求：需在专业加工厂组装，严禁现场组装。防雨百叶设不锈钢防虫网')
add_para('安装流程：点位复核 → 深化图纸确认 → 成品进场 → 安装固定 → 成品保护 → 确认验收')

doc.add_heading('5.5  淋水试验专项方案', level=2)
add_para('按万华集团外装淋水检验操作指引执行')
add_para('第一阶段（过程控制）：门窗、抹灰及外墙防水完成后，分区段自上而下进行。外墙淋水≥24h，外窗淋水≥6h。整改后需局部淋水检验')
add_para('第二阶段（全面性）：外墙饰面完成后或分户验收时进行。外墙淋水≥12h，外窗淋水≥6h')
add_para('喷壶淋水与布管淋水要求：【待补充：水压、孔径、孔距、淋水管布置等具体参数】')

doc.add_heading('5.6  后置埋件与防雷接驳方案', level=2)
add_para('后置埋件供应、安装，安装时的抹灰切除、基层找平、抹灰修复由外装负责')
add_para('负责屋面幕墙金属构件与防雷连接点的接驳、检测工作')

# 第六章 质量管理体系与保证措施
doc.add_heading('第六章  质量管理体系与保证措施', level=1)

doc.add_heading('6.1  质量目标与标准', level=2)
add_para('竣工验收条件按国家现行相关验收规范要求上限执行')
add_para('凡乙供材料，品质以甲方主要管理人员验收认可（平整度、色泽、均匀度、拼接方式、收缝关系、切割几何尺寸）为准')

doc.add_heading('6.2  质量管理措施', level=2)
add_para('严格执行"施工单位分项完工自检合格后甲方监理仅做分项验收"流程')
add_para('材料进场后按政府验收要求送检，费用由我方综合考虑')
add_para('资料必须按工序及时报送纸质版，不得后补')

# 第七章 安全管理体系与保证措施
doc.add_heading('第七章  安全管理体系与保证措施', level=1)

doc.add_heading('7.1  安全目标', level=2)
add_para('重伤、死亡人数为零，无重大机械设备、急性中毒事故')
add_para('特种设备备案强检合格率100%，特种作业人员持证上岗率100%')
add_para('施工现场扬尘治理必须达标')

doc.add_heading('7.2  安全管理措施', level=2)
add_para('签订《安全生产综合目标管理责任书》，缴纳安全保证金（合同额3‰，1万起，10万封顶）')
add_para('全面实施两项许可制度（安全许可证、三类人员任职资格）')
add_para('对深基坑、起重机械、脚手架、模板、施工用电、四口五临边等重点危险源全面监控')
add_para('幕墙安装工程属于超危大工程，专项施工方案需经专家论证')

# 第八章 文明施工与环境保护
doc.add_heading('第八章  文明施工与环境保护', level=1)

doc.add_heading('8.1  成品保护方案', level=2)
add_para('施工过程及工序衔接保护由我方自行负责，保护不力将追溯处罚')
add_para('施工区域地面及与原外装搭接位置做一层软保护+一层硬保护（不限于木工板等硬质材料）')
add_para('各甲分包单位完成并做好成品保护后，以工序移交形式书面移交')

doc.add_heading('8.2  建渣清运与保洁', level=2)
add_para('甲方现场指定临时场地堆放，满一车/斗（约8.5方）及时清运，每周不少于一次外运')
add_para('开业前对施工区域内成品保护拆除，以及全地块整体外立面精保洁一次（含未改造区域）')

doc.add_heading('8.3  水电及围挡管理', level=2)
add_para('甲方提供取水点、电箱，移交后由我方维护管理，水电费自行承担（含整改期间）')
add_para('施工围挡及脚手架自行搭设与拆除，脚手架必须提前报专项方案，持证上岗')

# 第九章 项目管理机构与人员配备
doc.add_heading('第九章  项目管理机构与人员配备', level=1)

doc.add_heading('9.1  项目组织架构', level=2)
add_para('项目经理（1人，全过程在岗）：统筹协调材料生产、现场施工、商务及外部协调')
add_para('安全员兼资料员（1人，全过程在岗，持证上岗）')
add_para('商务（根据阶段调配）：配合组价报价、三单梳理、进度款申报、竣工结算')

doc.add_heading('9.2  人员管理', level=2)
add_para('未经甲方批准，不得更换项目经理等主要人员')
add_para('大面展开时，必须增加管理人员，工作面全覆盖管理')

# 第十章 配合与服务承诺
doc.add_heading('第十章  配合与服务承诺', level=1)

doc.add_heading('10.1  农民工工资支付管理', level=2)
add_para('中标后开设农民工工资专用账户')
add_para('当期进度款的30%由甲方直接划转至专户，实现资金流向与用工管理双重闭环监管')
add_para('每月15日前提交劳务清单、考勤、工资发放表等资料')

doc.add_heading('10.2  配合甲方管理要求', level=2)
add_para('配合办理施工许可证（实名制通道、消纳证）')
add_para('配合现场打样（不排除多次打样确认）')
add_para('配合三单管理制度，服从限价要求')
add_para('配合甲供材料进场验收、卸货、盘点、保存、转运及上墙')

doc.add_heading('10.3  资料管理', level=2)
add_para('负责整理、收集分包工程范围内所有工程资料，配合总包完成竣工资料归档与备案')
add_para('施工过程资料必须按工序及时报送，不得后补')

# 第十一章 应急预案
doc.add_heading('第十一章  应急预案', level=1)
add_para('11.1  台风及极端天气应急预案')
add_para('11.2  高温中暑应急预案')
add_para('11.3  春节返乡及节后复工保障措施')
add_para('11.4  突发安全事故应急预案')

# ===== Save =====
output_path = r'D:\罗\2026\三亚\001\三亚06地块商业改造外装综合幕墙工程_技术标_v5_complete.docx'
doc.save(output_path)
print(f'Complete v5 saved to {output_path}')