"""cv_software.py —— 软件方向简历（AI 应用开发 / Agent / 后端）

侧重：Agent/MCP 工具开发、大模型应用架构（AI 看板）、60+ REST 接口、
Python/Java 全栈工程能力；能源类实习/科研压缩，科研保留量化与算法含量。

用法:
    python cv_software.py <输出路径.docx> [照片路径.png] [pdf路径.pdf]

默认输出: 脚本同目录下 Jason_CV_2026_Software.docx（同目录同时生成 .pdf）
"""

import os

from cv_common import runGenerate

DEFAULT_OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Jason_CV_2026_Software.docx')

HEADER_LEFT = (
    '姚冯杰',
    '电话：+86 18692250432   ｜   邮箱：fengjieyao330@gmail.com',
    '求职意向：AI 应用开发（Agent / 大模型应用 / 后端）   ｜   出生年月：2002.10',
)

SECTIONS = [
    ('教育经历', [
        ('entry', '2024.09-2025.11', '香港大学（HKU）工程学院', '电气工程（硕士）'),
        ('bullet', '主修课程', '机器学习、智慧电网、可再生能源科学、投资学'),
        ('entry', '2020.09-2024.06', '中南大学（CSU）能源科学与工程学院', '能源与动力工程（本科）'),
        ('bullet', '主修课程', '电子技术、机械设计、C语言、工程流体力学、热力学'),
    ]),
    ('工作/实习经历', [
        ('entry', '2025.12-至今', '宁德时代新能源科技股份有限公司', '智能体开发工程师'),
        ('bullet', '质量智能体（2026.06-至今｜Agent 开发）', '开发大模型函数调用的核心 MCP 工具：Neo4j FTA 失效模式/失效因子知识图谱读写（670 节点，多跳关系图/相邻节点/关系类型白名单防注入）、Milvus 向量检索（3000+ chunks、768 维，向量+SHA256 哈希双模）、相关性系数矩阵（Pearson/Spearman/Kendall+显著性分级）；从零开发 13 个 Agent 工具 REST 接口；设计结果展示闭环（自动拆解落库、admin/负责人/共享名单三级权限共享）'),
        ('bullet', '智能 FA 系统 · AI 看板（2026.03-至今｜独立完成）', '大模型对话式数据分析：Java(Spring Boot)+Python(FastAPI) 分工架构、串行任务队列+异步调用（504 容错/6 分钟超时治理）、多轮会话（Redis TTL 1h）、强约束提示词驱动 LLM 直接产出 ECharts 图表（存 S3 增量轮询）、clarify 追问澄清状态机、9 个 REST 接口，支持 xlsx/txt 上传分析'),
        ('bullet', '智能 FA 系统 · 业务后端（2026.03-至今｜30+ 接口）', '冷压断带智能监测（共享 CSV 增量同步判重、恶化 Top 两阶段环比+判责分析、多组多维度下钻、EasyExcel 流式导出）、FTA 因子推理（失效特征沿树回溯根因+部门级数据权限，从零开发）、风险监测日志、改善库（STOP/SKIP/OVERWRITE 三策略批量导入）、MU 产品线、桌面云外网示警等；持续 SQL 调优、导出异步化（导出成功率 99%+）；量产改善定时任务新增“根因推荐暂存”，单条失败不阻断整批（全岗位累计独立开发 60+ REST 接口）'),
        ('bullet', '智能工艺设计（2026.03-2026.05｜Python 后端）', '工艺文档向量化管理平台（FastAPI）：切块→768 维向量→Milvus(HNSW) 千条分批插入+失败回滚、SHA256 文件级严格查重、query+upsert 实现启用/停用/删除管理'),
        ('bullet', '厂房用电能耗项目（2025.12-2026.03｜数据开发）', '多源数据（设备编码/工序量测/用电单元清单）Hive/Iceberg/Doris 数据治理（统一指标口径/数据质量标准）、SQL+Airflow 数据流构建，参与 Python 数据分析平台服务接口封装'),
        ('bullet', 'AI 提效', '日常运用公司自研 OpenCode（DeepSeek/GLM/Qwen/MiniMax）与 GitHub Copilot 辅助 SQL/Python/Java 开发，OpenCode 亦为 AI 看板大模型执行引擎'),
        ('entry', '2025.06-2025.09', '图迹科技有限公司', '新能源解决方案实习生'),
        ('bullet', '零碳园区 ToG 项目', '政府客户（涞源/临平区政府）方案：能源资产聚合电力交易（AI 大模型电价预测、中长期+现货结合策略）、AI 零碳低空经济基础设施网（光伏+储能+智能调度）；完成客户对接方案 3 次、提案宣讲 2 次'),
    ]),
    ('科研经历', [
        ('entry', '2025.01-2025.08', '基于需求响应的能源市场机制设计', '硕士毕设 · 独立完成'),
        ('bullet', '', 'Python+Gurobi 构建 TOU/RTP/动态定价电力市场-HEMS 仿真平台：系统波动 -36.6%、可再生能源利用率 76.8%、用户电费 -18.4%，机制可迁移至虚拟电厂/储能 EMS'),
        ('entry', '2023-2024', '两篇独立作者（唯一作者）国际会议论文', '2 篇均录用（EI）'),
        ('bullet', '', '① Torque Distribution Optimization Strategy for Independently Driven Trams with Front and Rear Axles（AETS 2024，PID+PSO 优化）；② Brief Analysis of Aeroengine（CONF-CIAP 2023）'),
        ('entry', '2023.10-2024.10', '基于析氧电催化剂的锂电池正极材料高值化回收', '国家级大创 · 成员'),
        ('bullet', '', '去锂渣 SEM/EDS/XPS/XRD 表征，确定最优配比 SP:PVDF:去锂渣 = 1:1:9'),
    ]),
    ('专业技能', [
        ('bullet', '开发', 'Python、Java（Spring Boot/MyBatis-Plus）、SQL、FastAPI、Airflow、Git、Postman'),
        ('bullet', 'AI 工程', 'LLM 应用开发（提示词工程/Agent 工具/MCP）、RAG、Milvus、Neo4j、Langfuse、Redis'),
        ('bullet', '数据与存储', 'S3、Doris/Hive/Iceberg、EasyExcel'),
        ('bullet', '语言', 'IELTS 7.5（流利商务沟通）、CET-6 504'),
    ]),
]

if __name__ == '__main__':
    runGenerate(DEFAULT_OUTPUT, HEADER_LEFT, SECTIONS)
