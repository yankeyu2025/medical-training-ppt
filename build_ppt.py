# -*- coding: utf-8 -*-
"""31-slide Chinese medical training presentation. Run: pip install python-pptx && python build_ppt.py"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

OUT = "Heart_Failure_Antipsychotics_Training.pptx"
NAVY = RGBColor(23, 54, 85)
BLUE = RGBColor(33, 104, 163)
LIGHT = RGBColor(234, 244, 251)
GRAY = RGBColor(75, 90, 106)
RED = RGBColor(176, 58, 64)
WHITE = RGBColor(255,255,255)
FONT = "Microsoft YaHei"
prs=Presentation()
prs.slide_width=Inches(13.333)
prs.slide_height=Inches(7.5)

slides=[
("封面|心力衰竭与抗精神病药物相关心脏不良反应",
 ["医学培训 · 心力衰竭症状与体征、治疗目标","抗精神病药物相关心脏风险、发生机制与处置","适用对象：精神科医师、临床药师、护理人员"],"ESC 2021/2023; AHA/ACC/HFSA 2022; ESC 2025"),
("课程导览|学习目标与培训路径",
 ["识别：心力衰竭典型症状、体征及危急表现","理解：心力衰竭治疗目标及主要治疗原则","比较：不同抗精神病药物的心脏不良反应类型与风险","实践：心电图与生物标志物监测、异常结果处置"],""),
("第一部分｜基础|什么是心力衰竭？",
 ["心力衰竭（HF）是由于心脏结构和/或功能异常导致的临床综合征。","临床表现：呼吸困难、疲乏、运动耐量下降、水肿等。","诊断不能仅凭症状：结合体征、利钠肽和超声心动图。","精神科患者新出现持续心动过速、气促或水肿时应考虑心脏病因。"],"ESC HF 2021; AHA/ACC/HFSA 2022"),
("第一部分｜基础|按射血分数分类",
 ["HFrEF：LVEF ≤40%，射血分数降低型心衰","HFmrEF：LVEF 41%–49%，射血分数轻度降低型心衰","HFpEF：LVEF ≥50%，射血分数保留型心衰，须有充盈压升高等证据","HFimpEF：既往LVEF ≤40%，后续改善至 >40%（AHA/ACC/HFSA定义）","分型决定疾病修饰治疗策略，但均需评估淤血与诱因。"],"ESC HF 2021; AHA/ACC/HFSA 2022"),
("第一部分｜症状|心力衰竭的主要症状",
 ["劳力性呼吸困难、活动耐量下降、疲乏","端坐呼吸、夜间阵发性呼吸困难","踝部/下肢水肿、体重短期增加、腹胀","严重者静息气促、咳粉红色泡沫痰、少尿、意识改变","注意鉴别焦虑、镇静、肺部疾病与药物相关症状。"],"ESC HF 2021"),
("第一部分｜体征|淤血与低灌注：床旁观察",
 ["淤血：颈静脉压升高、肺部湿啰音、外周水肿、肝大/腹水","低灌注：四肢湿冷、少尿、低血压、意识状态改变","其他：心动过速、第三心音、呼吸频率增快、血氧下降","单一体征敏感性不足，应结合病史、生命体征及检查。"],"ESC HF 2021"),
("第一部分｜评估|NYHA心功能分级",
 ["Ⅰ级：一般体力活动不引起症状","Ⅱ级：日常活动轻度受限，普通活动出现症状","Ⅲ级：明显受限，低于普通活动量即出现症状","Ⅳ级：静息也有症状，任何活动均加重","NYHA评估功能受限程度，不等同于LVEF分型。"],"AHA/ACC/HFSA 2022"),
("第一部分｜评估|诊断流程与常用检查",
 ["询问病史及查体：高血压、冠心病、心肌炎、用药史等","心电图：心律、传导、缺血；胸片：肺淤血或心脏扩大","BNP/NT-proBNP：协助排除或支持心衰，受年龄、肾功能等影响","超声心动图：LVEF、室壁运动、瓣膜及结构功能","肌钙蛋白、电解质、肾功能、甲状腺功能用于寻找病因与诱因。"],"ESC HF 2021"),
("第一部分｜目标|心力衰竭的治疗目标",
 ["改善预后：降低心血管死亡和全因死亡风险","减少失代偿：降低心衰住院与再入院风险","改善症状：减轻气促、水肿与疲乏，提高生活质量","兼顾病因治疗、合并症管理、依从性和长期随访。"],"AHA/ACC/HFSA 2022"),
("第一部分｜治疗|HFrEF：改善预后的四类基石药物",
 ["ARNI（或ACEI/ARB）：抑制神经内分泌异常激活","有证据的β受体阻滞剂：如比索洛尔、卡维地洛、琥珀酸美托洛尔","MRA：螺内酯、依普利酮（监测血钾与肾功能）","SGLT2抑制剂：达格列净、恩格列净","有淤血时使用袢利尿剂改善症状；方案须结合血压、肾功能个体化。"],"ESC 2021/2023; AHA/ACC/HFSA 2022"),
("第一部分｜治疗|HFmrEF与HFpEF治疗重点",
 ["SGLT2抑制剂：减少HFmrEF/HFpEF心衰住院或心血管死亡复合风险","利尿剂：缓解容量负荷过多和水肿","积极治疗高血压、房颤、肥胖、糖尿病、睡眠呼吸暂停等","HFmrEF部分患者可考虑HFrEF相关治疗，依据指南和个体情况。"],"ESC HF Focused Update 2023"),
("第一部分｜急症|急性心衰的危险信号与初步处理",
 ["红旗：静息严重气促/低氧、肺水肿、低血压伴低灌注","红旗：胸痛、持续恶性心律失常、晕厥、意识障碍","立即呼叫急救团队；连续监测心电、血压、脉搏血氧","按当地急救规范评估氧疗、无创通气、静脉利尿及病因处置","疑似药物相关心肌损伤时同步回顾药物并尽快会诊。"],"ESC HF 2021"),
("第二部分｜背景|为什么精神科要重视心脏安全？",
 ["严重精神障碍人群心血管疾病负担较高。","抗精神病药可通过电生理、血流动力学、炎症或代谢途径影响心脏。","药物风险常与基础心脏病、吸烟、肥胖、合并用药等共同作用。","临床管理强调基线风险评估、动态监测与精神疾病控制之间的平衡。"],"ESC Clinical Consensus on Mental Health & CVD 2025"),
("第二部分｜全景|抗精神病药相关心脏不良反应",
 ["电生理：QTc延长、尖端扭转型室性心动过速（TdP）","心肌：心肌炎、心肌病、心功能不全","血流动力学：体位性低血压、反射性心动过速","长期风险：体重增加、糖脂代谢异常及动脉粥样硬化风险","应按具体结局比较药物，而不是给所有心脏不良反应一个统一排名。"],"ESC 2025; BHRS 2019"),
("第二部分｜机制|机制①：hERG/IKr阻滞与QTc延长",
 ["某些抗精神病药阻断心肌hERG（IKr）钾通道。","心室复极延迟 → 动作电位持续时间延长 → QTc延长。","复极不稳定可产生早期后除极，诱发多形性室速/TdP。","低钾、低镁、心动过缓、先天长QT及药物相互作用增加风险。"],"BHRS QT Guideline 2019"),
("第二部分｜机制|机制②：自主神经、炎症与代谢",
 ["α1受体阻断：外周血管舒张，体位性低血压及反射性心动过速","抗胆碱能作用：窦性心动过速，部分患者明显","氯氮平相关心肌炎：可能涉及免疫/炎症过程，机制未完全阐明","H1/5-HT2C等相关作用及生活方式因素：体重、胰岛素抵抗与血脂变化","药物相关风险受剂量、滴定速度、个体易感性及并用药影响。"],"ESC 2025"),
("第二部分｜比较|哪些药物更需要关注QTc？",
 ["QTc影响较明确/需重点关注：齐拉西酮、氨磺必利等；部分低效价第一代药物亦需关注。","氯丙嗪、喹硫平等具有QT延长风险；风险随剂量和患者情况变化。","阿立哌唑一般对QTc平均影响较小，但不意味着零风险。","氯氮平的独特高关注点是心肌炎/心肌病，不能仅凭QTc排序判断。","药物风险分层参照具体指南；不同研究与药物标签结果可能不同。"],"BHRS 2019; ESC 2025"),
("第二部分｜风险|QTc风险为何在患者间差异很大？",
 ["患者因素：结构性心脏病、心衰、女性、老龄、先天性长QT","可纠正因素：低钾、低镁、低钙、脱水、呕吐/腹泻","药物因素：高剂量、快速加量、多种QT延长药合用","相互作用：代谢抑制使抗精神病药暴露升高","心动过速时Bazett校正可能高估QTc，可考虑Fridericia校正。"],"BHRS 2019; ESC 2025"),
("第二部分｜监测|QTc测量、判读与复查",
 ["治疗前评估晕厥史、家族猝死史、基础心脏病和合并用药。","高风险方案建议基线12导联心电图及电解质（K、Mg、Ca）。","增加剂量或新增相互作用药物后，在适当时点复查心电图。","尽量在标准导联与一致的QT校正公式下进行纵向比较。","QTc >500 ms或较基线显著增长，需立即升级评估。"],"BHRS 2019"),
("第二部分｜处置|QTc延长分级处置",
 ["QTc <440 ms（男性）或 <470 ms（女性）：通常无需因QT单独调整，结合风险评估。","高于以上阈值且 <500 ms：考虑减量或换低风险药、纠正可逆因素、复查心电图。","QTc ≥500 ms：立即审查并通常停用可疑致QT药物，紧急心脏科评估与监护。","伴晕厥/前晕厥：立即心电监护并按心律失常急症处理。","阈值来自BHRS指导，临床应综合QRS、校正方法和具体情境。"],"BHRS 2019"),
("第二部分｜急症|尖端扭转型室速（TdP）的紧急处置",
 ["立即启动急救：评估意识、脉搏、血流动力学并持续监护。","无脉搏或不稳定恶性室速：按照高级生命支持流程电除颤/复苏。","停用致QT药，检查并积极纠正低钾和低镁。","静脉硫酸镁为TdP常用急症治疗，具体用法依医院急救规范。","反复发作需心脏专科评估心动过缓/停搏依赖因素及起搏等措施。"],"AHA ACLS; BHRS 2019"),
("第二部分｜氯氮平|心肌炎与心肌病：独立关注点",
 ["氯氮平相关心肌炎多在治疗早期，尤其开始后的前4周出现。","表现可不典型：持续心动过速、低热、胸痛、气促、疲乏。","心肌病可发生于更晚阶段，表现为活动耐量下降及心衰。","不能将滴定期心动过速均视为良性药理反应。","早识别、及时停药评估是降低严重后果的关键。"],"Ronaldson et al. 2011; ESC 2025"),
("第二部分｜机制|氯氮平心肌炎的可能机制",
 ["目前病理机制尚未完全确立，可能涉及免疫过敏及炎症反应。","可能出现心肌细胞损伤和炎症生物标志物升高。","快速滴定、个体因素和某些合并用药在部分研究中与风险相关。","肌钙蛋白、CRP及症状动态变化比单看心率更有意义。"],"Ronaldson et al. 2011; ESC 2025"),
("第二部分｜监测|氯氮平启动前4周：监护思路",
 ["开始前：评估心脏史及症状；记录心率/体温，考虑基线ECG、CRP、肌钙蛋白。","前4周：依据本院方案至少每周监测CRP、肌钙蛋白，同时观察生命体征。","若持续心动过速、胸痛、低热或CRP异常：立即升级监测频次与医学评估。","参考Ronaldson方案：肌钙蛋白 >正常上限2倍或CRP >100 mg/L支持立即停用并评估。","监测流程应结合当前本地指南、实验室参考范围及心脏科意见。"],"Ronaldson et al. Aust N Z J Psychiatry 2011"),
("第二部分｜处置|疑似氯氮平心肌炎/心肌病",
 ["1. 识别症状：胸痛、气促、持续心动过速、发热、乏力或水肿。","2. 急查ECG、肌钙蛋白、CRP、BNP/NT-proBNP及超声心动图（视情况）。","3. 高度怀疑或严重异常：及时停用氯氮平，并紧急心脏科评估。","4. 排查感染、肺栓塞、急性冠脉综合征等鉴别诊断。","5. 后续精神症状替代治疗与是否再挑战由多学科审慎决定。"],"ESC 2025; Ronaldson 2011"),
("第二部分｜其他风险|体位性低血压与心动过速",
 ["可关注药物：氯氮平、氯丙嗪、喹硫平等（α1阻断/自主神经相关）。","症状：站立头晕、心悸、跌倒或晕厥；与脱水、降压药联用有关。","处理：测卧/立位血压和心率，排查感染、贫血、心肌炎与心律失常。","必要时调整滴定速度、给药方案和合并用药，强化跌倒预防。"],"ESC 2025; medication labeling"),
("第二部分｜长期风险|代谢风险与长期心血管管理",
 ["氯氮平、奥氮平通常更需关注体重增加与代谢异常。","持续监测体重/BMI、腰围、血压、空腹血糖或HbA1c、血脂。","评估吸烟、运动、饮食及合并代谢疾病，必要时多学科干预。","长期心血管风险不等同于QT延长风险，换药应综合精神科疗效。"],"ESC 2025"),
("第三部分｜落实|精神科病区心脏安全监护清单",
 ["治疗前：家族史/心脏史、晕厥、基础ECG、危险合并用药、电解质。","治疗中：胸痛、气促、心悸、晕厥、下肢水肿、生命体征变化。","针对QT风险：复查ECG、校正电解质、审查剂量和相互作用。","针对氯氮平：重点监测初始阶段心动过速、CRP/肌钙蛋白。","长期：体重、血压、血糖、血脂及综合心血管风险。"],"ESC 2025; BHRS 2019"),
("第三部分｜病例|病例讨论：氯氮平第18天出现心动过速",
 ["病例：患者开始氯氮平第18天，静息心率120次/分，伴低热、乏力、轻度胸闷。","问题1：能否仅将其解释为氯氮平抗胆碱能效应？","问题2：应立即评估哪些指标，如何排查心肌炎？","建议：生命体征、12导联ECG、肌钙蛋白、CRP及必要的心脏超声。","若心肌损伤证据明显或高度怀疑，应停止可疑药并紧急会诊。"],"Ronaldson 2011; ESC 2025"),
("总结|五个临床带走点",
 ["① 心衰：同时观察呼吸困难、水肿、肺淤血及低灌注征象。","② 目标：降低死亡和住院风险，改善症状与生活质量。","③ QT风险：关注药物本身、合并用药、电解质和既往心脏病。","④ 氯氮平：持续心动过速伴低热/胸闷须警惕心肌炎。","⑤ 监护：基线评估、动态复测、分级处置和多学科协作。"],"ESC 2021/2023; BHRS 2019; ESC 2025"),
("参考文献|主要指南与资料",
 ["McDonagh TA, et al. 2021 ESC Guidelines for the diagnosis and treatment of acute and chronic heart failure. Eur Heart J. 2021. DOI:10.1093/eurheartj/ehab368.",
 "McDonagh TA, et al. 2023 Focused Update of the 2021 ESC heart failure guidelines. Eur Heart J. 2023. DOI:10.1093/eurheartj/ehad195.",
 "Heidenreich PA, et al. 2022 AHA/ACC/HFSA Guideline for the Management of Heart Failure. Circulation. 2022.",
 "Lambiase PD, et al. British Heart Rhythm Society Clinical Practice Guidelines on QT prolongation with antipsychotics. Arrhythm Electrophysiol Rev. 2019.",
 "Ronaldson KJ, et al. A new monitoring protocol for clozapine-induced myocarditis. Aust N Z J Psychiatry. 2011.",
 "European Society of Cardiology. Clinical consensus statement on mental health and cardiovascular disease. Eur Heart J. 2025."],""),
]

def tb(slide,x,y,w,h,txt,size=21,bold=False,color=GRAY):
 sh=slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
 tf=sh.text_frame; tf.word_wrap=True
 for k,line in enumerate(txt.split("\n")):
  p=tf.paragraphs[0] if k==0 else tf.add_paragraph()
  p.text=line; p.font.name=FONT; p.font.size=Pt(size); p.font.bold=bold; p.font.color.rgb=color
 return sh

for idx,(heading,lines,ref) in enumerate(slides,1):
 slide=prs.slides.add_slide(prs.slide_layouts[6])
 slide.background.fill.solid(); slide.background.fill.fore_color.rgb=WHITE
 band=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(.13))
 band.fill.solid(); band.fill.fore_color.rgb=BLUE; band.line.fill.background()
 section,title=heading.split("|",1)
 tb(slide,.72,.38,11.8,.36,section,13,True,BLUE)
 tb(slide,.72,.93,12.0,1.05,title,30,True,NAVY)
 line=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(.72), Inches(2.02), Inches(11.8), Inches(.025))
 line.fill.solid(); line.fill.fore_color.rgb=RGBColor(213,225,234); line.line.fill.background()
 maxsize=20 if len(lines)<6 else 17
 if idx==31:maxsize=13
 top=2.32
 available=4.52
 itemheight=min(.87,available/len(lines))
 for j,s in enumerate(lines):
  y=top+j*itemheight
  circle=slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(.83), Inches(y+.14), Inches(.13), Inches(.13))
  circle.fill.solid();circle.fill.fore_color.rgb=BLUE;circle.line.fill.background()
  tb(slide,1.15,y,11.25,itemheight-.02,s,maxsize,False,GRAY)
 if ref:tb(slide,.73,7.08,11.75,.23,ref,9,False,RGBColor(116,131,145))
 tb(slide,12.5,7.08,.42,.23,f"{idx:02d}",10,True,BLUE)

prs.save(OUT)
print(f"Created {OUT} with {len(prs.slides)} slides")
