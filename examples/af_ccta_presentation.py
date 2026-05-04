"""
使用 McKinsey PPTX 工具生成心房颤动患者冠状动脉CT血管造影研究的学术演示文稿
基于 medical-imaging-research 技能的深度分析结果
"""
from __future__ import annotations
from pathlib import Path

import sys
sys.path.append(r'c:\Users\tiech\.trae-cn\mckinsey-pptx')

from mckinsey_pptx import PresentationBuilder


def build_af_ccta_presentation(output_path: str = "output/af_ccta_presentation.pptx") -> str:
    """构建心房颤动患者冠状动脉CT血管造影研究的麦肯锡风格演示文稿"""

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    b = PresentationBuilder(default_section_marker="核心发现")

    # ============================================
    # 1. 执行摘要 - 关键结论
    # ============================================
    b.add("executive_summary_takeaways",
          title="执行摘要",
          sections=[
              {
                  "takeaway": "研究背景",
                  "bullets": [
                      "心房颤动（AF）是成人最常见的心律失常",
                      "CAD在AF患者中的患病率约为20%-40%",
                      "CCTA在AF患者中存在技术挑战"
                  ]
              },
              {
                  "takeaway": "核心发现",
                  "bullets": [
                      "98%冠状动脉节段具有诊断价值",
                      "患者水平ROC曲线下面积达0.97",
                      "前瞻性触发方案减少约50%辐射剂量"
                  ]
              },
              {
                  "takeaway": "临床建议",
                  "bullets": [
                      "推荐使用≥128层CT扫描仪",
                      "优先采用前瞻性ECG触发方案",
                      "个体化对比剂和电压方案可优化图像质量"
                  ]
              }
          ],
          final_conclusion="370 mgI/ml碘普罗胺增强CCTA在AF患者中提供满意的诊断性能，是一种可行且有效的CAD诊断方法")

    # ============================================
    # 2. 评估表格 - 研究特征概览
    # ============================================
    b.add("assessment_table",
          title="研究特征概览",
          categories=[
              {
                  "name": "研究规模",
                  "rows": [
                      {"kpi": "纳入研究数", "target": "14项", "actual": "14项", "status_label": "达标", "status": "green"},
                      {"kpi": "患者总数", "target": "601例", "actual": "601例", "status_label": "达标", "status": "green"},
                      {"kpi": "随机对照试验", "target": "2项", "actual": "2项", "status_label": "达标", "status": "green"},
                      {"kpi": "观察性研究", "target": "12项", "actual": "12项", "status_label": "达标", "status": "green"}
                  ]
              },
              {
                  "name": "患者特征",
                  "rows": [
                      {"kpi": "中位年龄", "target": "64±13.0岁", "actual": "64±13.0岁", "status_label": "达标", "status": "green"},
                      {"kpi": "BMI范围", "target": "22.5-29.3", "actual": "22.5-29.3", "status_label": "达标", "status": "green"},
                      {"kpi": "心率范围", "target": "67-96 bpm", "actual": "67-96 bpm", "status_label": "达标", "status": "green"},
                      {"kpi": "心率变异性", "target": "23-71 bpm", "actual": "23-71 bpm", "status_label": "达标", "status": "green"}
                  ]
              },
              {
                  "name": "CCTA技术参数",
                  "rows": [
                      {"kpi": "对比剂", "target": "370 mgI/ml", "actual": "370 mgI/ml", "status_label": "达标", "status": "green"},
                      {"kpi": "流速", "target": "3.5-6.0 ml/s", "actual": "3.5-6.0 ml/s", "status_label": "达标", "status": "green"},
                      {"kpi": "前瞻性触发", "target": "10项研究", "actual": "10项研究", "status_label": "达标", "status": "green"},
                      {"kpi": "≥128层扫描仪", "target": "9项研究", "actual": "9项研究", "status_label": "达标", "status": "green"}
                  ]
              }
          ])

    # ============================================
    # 3. 柱状图 - 图像质量结果
    # ============================================
    b.add("column_comparison",
          title="图像质量分析",
          categories=[
              "总体诊断价值",
              "优秀质量",
              "HRV>50bpm",
              "BMI≥24",
              "瓣膜疾病",
              "≥128层扫描仪",
              "个体化方案"
          ],
          values=[98, 54, 98, 99, 99, 98, 98],
          focus_index=0,
          data_label="诊断质量",
          data_unit="%",
          description="冠状动脉节段诊断质量比较",
          takeaways=[
              "98%的冠状动脉节段具有诊断价值",
              "54%的节段达到优秀图像质量",
              "各亚组分析结果一致（95%-99%）",
              "≥128层扫描仪显著提高图像质量（98% vs 94%）"
          ])

    # ============================================
    # 4. 柱状图 - 诊断性能
    # ============================================
    b.add("column_comparison",
          title="诊断性能分析",
          categories=[
              "患者-敏感性",
              "患者-特异性",
              "节段-敏感性",
              "节段-特异性"
          ],
          values=[93, 90, 86, 98],
          focus_index=1,
          data_label="诊断效能",
          data_unit="%",
          description="患者水平 vs 节段水平诊断效能",
          takeaways=[
              "患者水平：敏感性93%，特异性90%",
              "节段水平：敏感性86%，特异性98%",
              "特异性在患者和节段水平均保持高水平",
              "ROC曲线下面积：患者水平0.97，节段水平0.99"
          ])

    # ============================================
    # 5. 柱状图 - 辐射剂量对比
    # ============================================
    b.add("column_comparison",
          title="辐射剂量分析",
          categories=[
              "总体剂量",
              "前瞻性触发",
              "回顾性门控"
          ],
          values=[821, 558, 1239],
          focus_index=1,
          data_label="辐射剂量",
          data_unit="mSv",
          description="不同扫描方案的辐射剂量比较",
          takeaways=[
              "总体剂量：8.21 mSv（95% CI: 6.28-10.13）",
              "前瞻性触发方案：5.58 mSv（95% CI: 4.11-7.05）",
              "回顾性门控方案：12.39 mSv（95% CI: 10.36-14.42）",
              "前瞻性触发方案减少约50%辐射剂量"
          ])

    # ============================================
    # 6. 优先级矩阵 - 技术优化策略
    # ============================================
    b.add("prioritization_matrix",
          title="技术优化策略",
          items=[
              {"name": "≥128层CT", "x_band": 2, "y_band": 0, "ox": 0.25, "oy": 0.5, "status": "green"},
              {"name": "前瞻性ECG触发", "x_band": 2, "y_band": 0, "ox": 0.7, "oy": 0.5, "status": "green"},
              {"name": "个体化电压", "x_band": 1, "y_band": 0, "ox": 0.65, "oy": 0.6, "status": "green"},
              {"name": "个体化对比剂", "x_band": 1, "y_band": 1, "ox": 0.3, "oy": 0.5, "status": "amber"},
              {"name": "碘递送率1.85 gI/s", "x_band": 1, "y_band": 1, "ox": 0.85, "oy": 0.85, "status": "amber"},
              {"name": "β受体阻滞剂", "x_band": 0, "y_band": 2, "ox": 0.55, "oy": 0.30, "status": "amber"}
          ])

    # ============================================
    # 7. BCG矩阵 - 特殊人群分析
    # ============================================
    b.add("growth_share",
          title="特殊人群分析",
          bus=[
              {"name": "瓣膜疾病\n99%", "x": 75, "y": 85, "size": 4},
              {"name": "BMI≥24\n99%", "x": 70, "y": 80, "size": 3.5},
              {"name": "HRV>50bpm\n98%", "x": 55, "y": 70, "size": 3},
              {"name": "慢性/永久性AF\n96%", "x": 35, "y": 55, "size": 2.5},
              {"name": "64层扫描\n94%", "x": 20, "y": 40, "size": 2}
          ])

    # ============================================
    # 8. 三个关键发现
    # ============================================
    b.add("three_trends_icons",
          title="关键发现",
          subtitle="370 mgI/ml碘普罗胺增强CCTA的核心优势",
          trends=[
              {
                  "label": "高质量成像",
                  "icon": "✓",
                  "bullets": [
                      "98%节段具有诊断价值",
                      "54%节段达到优秀质量",
                      "与窦性心律患者相当"
                  ]
              },
              {
                  "label": "优异诊断性能",
                  "icon": "◎",
                  "bullets": [
                      "患者水平AUC: 0.97",
                      "节段水平AUC: 0.99",
                      "敏感性93%，特异性90%"
                  ]
              },
              {
                  "label": "低辐射暴露",
                  "icon": "↓",
                  "bullets": [
                      "前瞻性触发仅5.58 mSv",
                      "减少约50%辐射剂量",
                      "符合ALARA原则"
                  ]
              }
          ])

    # ============================================
    # 9. 五个关键领域
    # ============================================
    b.add("five_key_areas",
          title="临床实践建议",
          subtitle="基于系统评价和元分析的建议",
          areas=[
              {
                  "name": "扫描仪选择",
                  "description": "优先使用≥128层CT扫描仪，可显著提高图像质量和诊断准确性"
              },
              {
                  "name": "采集方案",
                  "description": "推荐前瞻性ECG触发方案，在保证图像质量的同时显著降低辐射剂量"
              },
              {
                  "name": "对比剂方案",
                  "description": "采用370 mgI/ml碘普罗胺，个体化对比剂和电压方案优化图像质量"
              },
              {
                  "name": "特殊人群",
                  "description": "高BMI、瓣膜疾病、高心率变异性患者均可获得满意的诊断图像"
              },
              {
                  "name": "辐射防护",
                  "description": "前瞻性触发方案减少约50%辐射剂量，特别适合需要重复成像的AF患者"
              }
          ])

    # ============================================
    # 10. 执行摘要 - 结论
    # ============================================
    b.add("executive_summary_takeaways",
          title="结论",
          sections=[
              {
                  "takeaway": "主要结论",
                  "bullets": [
                      "370 mgI/ml碘普罗胺增强CCTA在AF患者中提供满意的诊断性能",
                      "高性能扫描仪（≥128层）、个体化对比剂方案可进一步提高性能",
                      "前瞻性ECG触发CCTA可维持足够图像质量同时显著减少辐射剂量"
                  ]
              },
              {
                  "takeaway": "临床应用价值",
                  "bullets": [
                      "CCTA是AF患者CAD诊断的可行且有效方法",
                      "图像质量和诊断性能与窦性心律患者相当",
                      "β受体阻滞剂在许多情况下可安全省略，简化检查流程"
                  ]
              },
              {
                  "takeaway": "未来研究方向",
                  "bullets": [
                      "需要更大规模的头对头研究比较不同采集方案",
                      "对比剂类别效应的验证需要更大规模分析",
                      "定量图像质量参数的标准化报告"
                  ]
              }
          ],
          final_conclusion="对于AF患者，CCTA是一种可行且有效的CAD诊断方法，应优先采用前瞻性触发方案和个体化技术参数")

    # 保存演示文稿
    b.save(output_path)
    return output_path


if __name__ == "__main__":
    output = build_af_ccta_presentation()
    print(f"麦肯锡风格PPT已生成: {output}")
