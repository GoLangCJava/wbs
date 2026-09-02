# -*- coding: utf-8 -*-
"""新增「功能拆解估算」页：把26项功能下钻到子任务/工作包级，
每个子任务按 BA / 开发A(Wade) / 开发B(DevB) / 测试 / SLC文档 直估人天（独立估算，不锚定185人天）。
用法：python3 func_breakdown.py
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PATH = "/home/user/wbs/RIMS项目WBS工作分解_AI Native版.xlsx"
wb = openpyxl.load_workbook(PATH)

# ---------- 子任务数据：(功能编号, [(子任务, 工作内容要点, BA, Wade, DevB, Test, SLC), ...]) ----------
B = [
("F01-01", [
 ("原系统基本信息录入", "来源系统档案：系统编码/名称/业务域/责任人/退役批次；录入与维护界面+存储API", .5, 1, .5, .5, .25),
 ("源库DB连接配置", "DB类型/连接串录入，账号密码Key Vault加密存储，连通性测试", .25, 1, 1, .5, .25),
 ("同步表勾选", "拉取源表清单、勾选同步范围、数据量预估与全选过滤", .25, 1, 1, .5, .25),
 ("保留期限与迁移参数配置", "保留策略绑定、迁移批次命名、并发/限速参数", .5, 1, .5, .25, .25),
 ("存储落地实现（Iceberg/UC）", "磁盘目录布局、Iceberg表属性与分区策略、UC目录/Schema/授权配置化", .25, 0, 2, .5, .5),
 ("SeaTunnel批量导入作业", "作业配置生成器、分片/断点续传/限速执行", 0, .5, 2, 1, .5),
 ("导入执行与监控", "任务状态/进度/失败重试/日志查看", 0, 1, .5, .5, .25),
]),
("F01-02", [
 ("文件来源登记与上传通道", "在线上传/共享路径挂载/离线介质登记", .25, 1.5, 1, .5, .25),
 ("大文件分片上传与断点续传", "前端分片上传、后端合并与校验", 0, 1.5, 1.5, 1, .25),
 ("文件包解析", "zip/tar/数据库备份包识别与展开、原始属性保留", 0, 0, 1.5, .75, .25),
 ("附件元数据与Blob存储布局", "附件索引表、路径规范、与结构化记录关联", .25, 0, 1.5, .5, .25),
]),
("F01-03", [
 ("字段映射规则配置", "源字段→标准字段映射界面、自动映射建议", .5, 1.5, .5, .5, .25),
 ("格式转换与清洗规则", "类型转换/编码/默认值/脏数据处理策略", .5, 0, 1.5, .75, .25),
 ("校验规则配置", "必填/唯一/范围/自定义表达式校验", 0, 1, 1, .5, .25),
 ("规则版本与试跑预览", "规则版本管理、样本试跑与结果预览", 0, .5, 1, .5, .25),
]),
("F01-04", [
 ("基准采集策略定义", "行数/校验和/SHA-256抽样口径", .5, 0, 1, .25, .25),
 ("基准采集作业与存储", "源端基准快照生成、基准库表", 0, 0, 1.5, .5, .25),
 ("基准报告查询展示", "基准结果展示与导出", 0, .5, .25, .25, .25),
]),
("F01-05", [
 ("比对引擎", "RAW/CURATED逐层行数与校验和比对、抽样SHA-256", 0, 0, 2, .75, .25),
 ("差异报告与处理流程", "差异清单、重迁/豁免记录", .25, 1, .5, .5, .25),
]),
("F02-01", [
 ("元数据字段体系定义", "字段/类型/必填/密级/保留等级字典", 1.5, 0, .5, .25, .5),
 ("元数据核心表与API", "核心元数据表实现、CRUD API、Migration", 0, 0, 2, .5, .25),
 ("元数据驱动配置机制", "查询配置模型、动态表单渲染协议", .25, 1.5, .5, .5, .25),
]),
("F02-02", [
 ("绑定规则与批次属性", "源编码引用、退役项目/迁移批次属性规则", .5, 0, .75, .25, .25),
 ("接入时自动绑定", "导入/上传时来源属性自动写入与校验", 0, 0, 1, .5, .25),
 ("按来源检索与统计", "来源维度筛选、统计报表", 0, 1, .5, .5, .25),
]),
("F03-01", [
 ("保留策略配置", "策略规则、继承关系、生效范围", .75, 1, .5, .5, .25),
 ("处置引擎", "到期扫描、豁免检查、DROP PARTITION、VACUUM物理清除", 0, 0, 2.5, 1, .5),
 ("处置计划审批与台账", "预演清单、审批流、处置留痕台账", .25, 1, .5, .5, .25),
]),
("F03-02", [
 ("Hold申请与审批流程", "涉诉/审计/调查场景、审批链定义", .75, 1, .25, .5, .25),
 ("Hold生效与解除实现", "记录级/表级冻结、处置豁免、操作留痕", 0, 0, 1.5, .75, .25),
 ("Hold清单管理与提醒", "在Hold清单、解除预警", 0, .5, .25, .25, .25),
]),
("F04-01", [
 ("加密策略定义", "SSE+CMK方案、PII字段清单与字段级加密规则", .5, 0, .75, .25, .25),
 ("存储层加密落地", "ADLS SSE、Key Vault CMK集成、密钥轮换", 0, 0, 1.5, .5, .25),
 ("PII字段级加密服务", "应用层加解密、密文索引策略、传输加密验证", 0, 0, 1.5, .75, .25),
]),
("F04-02", [
 ("密钥全流程管理", "创建/保管/更新/吊销/权限控制", .25, 0, 1.5, .5, .25),
 ("密钥操作审计与告警", "操作日志、异常使用告警", 0, .5, .5, .5, .25),
 ("密钥管理界面", "密钥清单、版本、权限分配界面", 0, 1, .25, .25, .25),
]),
("F04-03", [
 ("备份策略配置", "备份周期/范围/保留期策略", .5, .75, .5, .25, .25),
 ("备份作业与副本创建", "元数据库PITR、湖快照、附件副本", 0, 0, 1.5, .5, .25),
 ("备份监控与恢复操作", "备份状态监控、恢复流程", 0, .5, .5, .5, .25),
]),
("F04-04", [
 ("灾备方案与RTO/RPO定义", "灾备架构、恢复目标定义", .5, 0, .5, .25, .25),
 ("异地副本同步", "存储异地冗余/复制配置", 0, 0, 1.5, .5, .25),
 ("恢复切换演练", "演练脚本、切换/回切流程与报告", 0, 0, 1, .75, .25),
]),
("F05-01", [
 ("检索配置模型", "可检索字段、分词策略、OCR文本索引策略", .5, .5, 1, .25, .25),
 ("检索页动态表单", "元数据驱动渲染、关键词与组合查询", 0, 1.5, .25, .5, .25),
 ("查询代理与SERVE层对接", "查询服务、索引优化、结果高亮", 0, 0, 1.25, .75, .25),
]),
("F05-02", [
 ("来源系统筛选组件", "来源系统下拉/搜索/多选", 0, .75, 0, .25, .25),
 ("来源维度查询下推", "sys_id过滤、分区裁剪", 0, 0, .5, .25, .25),
]),
("F05-03", [
 ("时间维度筛选组件", "创建/修改/业务发生时间范围选择", 0, .75, 0, .25, .25),
 ("时间分区下推与索引", "分区裁剪、时间索引优化", 0, 0, .5, .25, .25),
]),
("F05-04", [
 ("列表展示与分页排序", "关键字段展示、分批加载、虚拟滚动", 0, 1.5, 0, .5, .25),
 ("大批量结果性能", "异步计数、游标分页、缓存", 0, .25, .75, .5, .25),
]),
("F05-05", [
 ("预览服务", "格式识别、图片/PDF/Office转换", 0, 0, 1.5, .5, .25),
 ("SAS直连与流式加载", "临时SAS链接、流式传输、权限校验", 0, .5, 1, .5, .25),
 ("前端预览组件", "格式适配、缩放、失败降级下载", 0, 1, 0, .25, .25),
]),
("F06-01", [
 ("埋点规范与日志模型", "操作对象/用户/时间/结果/IP模型定义", .25, 0, .5, .25, .25),
 ("操作拦截与落库", "归档/导入/上传统一拦截器、异步写入", 0, 0, 1, .5, .25),
 ("审计查询界面", "多维筛选与导出（覆盖归档/查询/导出三类日志）", 0, 1, .25, .5, .25),
]),
("F06-02", [
 ("查询/预览行为埋点", "查询条件脱敏记录、预览行为记录", 0, 0, .75, .5, .25),
 ("异步批量写入", "缓冲批量写入、不阻塞查询主链路", 0, 0, .5, .25, .25),
]),
("F06-03", [
 ("导出行为审计", "导出范围/条数/导出人/IP/下载链接追踪", 0, 0, .75, .5, .25),
 ("大批量导出告警", "阈值告警、审批联动", 0, .5, .25, .25, .25),
]),
("F08-01", [
 ("RBAC模型与权限矩阵", "角色/菜单/数据范围/记录类型/操作权限定义", 1, 0, .5, .25, .5),
 ("角色管理界面与API", "角色CRUD、权限分配界面与接口", 0, 1.5, .5, .5, .25),
 ("权限校验与UC授权映射", "菜单/接口级校验中间件、UC权限同步", 0, 0, 1, .75, .25),
]),
("F08-02", [
 ("账号CRUD与生命周期", "创建/启用禁用/密码重置/有效期", .25, 1, .75, .5, .25),
 ("账号与角色绑定", "分配、批量操作、人员交接", 0, .5, .25, .25, .25),
]),
("F08-03", [
 ("Entra ID对接配置", "应用注册、回调配置、claims映射", .25, .5, 1, .5, .25),
 ("SSO登录链路", "登录页、令牌交换、会话管理", 0, 1.5, .5, .75, .25),
 ("认证异常场景处理", "认证失败、账号锁定、多账号冲突", 0, .5, 0, .5, .25),
]),
("F08-04", [
 ("身份映射规则", "企业账号↔RIMS用户映射、自动建号规则", .5, 0, .5, .25, .25),
 ("首登自动关联与角色赋予", "默认角色、待分配队列", 0, .25, .75, .5, .25),
]),
("F08-05", [
 ("会话与超时策略", "超时控制、登出、单点登出", .25, 1, .25, .5, .25),
 ("异常访问控制", "账号锁定、认证失效限制访问、IP策略", 0, 0, .75, .5, .25),
]),
]

# ---------- 功能名（读覆盖对照页，保证一字不差） ----------
wst = wb["26项功能覆盖对照"]
FUNC_NAME, MOD_OF = {}, {}
for r in range(3, wst.max_row + 1):
    fid = wst.cell(row=r, column=1).value
    if fid and str(fid).startswith("F"):
        FUNC_NAME[str(fid)] = str(wst.cell(row=r, column=5).value)
        MOD_OF[str(fid)] = str(wst.cell(row=r, column=2).value)
assert set(FUNC_NAME) == {f[0] for f in B}, "功能清单与覆盖页不一致"

# ---------- 渲染（合并单元格版式）：功能名/功能小计纵向合并，不占汇总行 ----------
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_TITLE = Font(bold=True, size=14, color="FFFFFF")
F_HDR = Font(bold=True, color="FFFFFF")
F_L1 = Font(bold=True)
FILL_NAVY = PatternFill("solid", fgColor="305496")
FILL_L1 = PatternFill("solid", fgColor="D9E1F2")
CENTER = Alignment(horizontal="center", vertical="center")

MOD_FILL = {}
_cur = None
_palette = ["EDEDF7", "E2EFDA", "FFF2CC", "FCE4D6", "DDEBF7", "E7E6E6", "E4DFEC"]
for fid in FUNC_NAME:
    if fid[:3] != _cur:
        _cur = fid[:3]
        MOD_FILL[_cur] = _palette[len(MOD_FILL) % len(_palette)]

if "功能拆解估算" in wb.sheetnames:
    del wb["功能拆解估算"]
wsn = wb.create_sheet("功能拆解估算", wb.sheetnames.index("功能视角分解") + 1)
wsn.sheet_properties.tabColor = "8FA9DB"
wsn.merge_cells("A1:K1")
wsn["A1"] = "功能拆解估算（子任务级直估）：功能信息纵向合并对应多个子任务 — 独立估算，不锚定排期185人天"
wsn["A1"].font = F_TITLE; wsn["A1"].fill = FILL_NAVY; wsn["A1"].alignment = CENTER
wsn.row_dimensions[1].height = 26
GH = ["子任务编号", "功能名称", "子任务/工作包", "工作内容要点", "BA人天", "开发A·Wade", "开发B·DevB", "测试人天", "SLC文档", "小计", "功能小计"]
GW = [11, 20, 24, 42, 8, 10, 10, 8, 8, 8, 9]
for i, (h, w) in enumerate(zip(GH, GW), 1):
    wsn.column_dimensions[get_column_letter(i)].width = w
    c = wsn.cell(row=2, column=i, value=h); c.font = F_HDR; c.fill = FILL_NAVY; c.alignment = CENTER; c.border = BORDER

r = 3
gt = [0.0] * 5
for fcode, subs in B:
    fill = PatternFill("solid", fgColor=MOD_FILL[fcode[:3]])
    r0 = r
    for n, (name, desc, ba, wa, db, te, sl) in enumerate(subs, 1):
        vals = [f"{fcode}-{n}", None, name, desc, ba or None, wa or None, db or None, te or None, sl or None,
                round(ba + wa + db + te + sl, 2)]
        for i, v in enumerate(vals, 1):
            c = wsn.cell(row=r, column=i, value=v); c.border = BORDER
            c.alignment = CENTER if i in (1, 5, 6, 7, 8, 9, 10) else Alignment(wrap_text=True, vertical="center")
            if i in (5, 6, 7, 8, 9, 10): c.number_format = "0.##"
        wsn.row_dimensions[r].height = 26
        r += 1
    r1 = r - 1
    # 功能名称（B列）与功能小计（K列）纵向合并
    wsn.merge_cells(start_row=r0, start_column=2, end_row=r1, end_column=2)
    cB = wsn.cell(row=r0, column=2)
    cB.value = FUNC_NAME[fcode]
    cB.font = F_L1; cB.fill = fill
    cB.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
    wsn.merge_cells(start_row=r0, start_column=11, end_row=r1, end_column=11)
    fs = round(sum(s[i] for s in subs for i in range(2, 7)), 2)
    cK = wsn.cell(row=r0, column=11, value=fs)
    cK.font = F_L1; cK.fill = fill; cK.alignment = CENTER; cK.number_format = "0.##"
    for rr in range(r0, r1 + 1):  # 合并区域补边框
        wsn.cell(row=rr, column=2).border = BORDER
        wsn.cell(row=rr, column=11).border = BORDER
    for s in subs:
        for i, v in enumerate(s[2:7]):
            gt[i] += v
# 总计行
labels = ["", "总计", "26项功能 × " + str(sum(len(s) for _, s in B)) + "个子任务",
          "独立直估（不含项目初始化/架构设计/集成/UAT/上线/PM等公共工作）",
          gt[0], gt[1], gt[2], gt[3], gt[4], round(sum(gt), 2), ""]
for i, v in enumerate(labels, 1):
    c = wsn.cell(row=r, column=i, value=v); c.border = BORDER; c.fill = FILL_L1; c.font = F_L1
    c.alignment = CENTER if i not in (3, 4) else Alignment(horizontal="left", vertical="center")
    if i in (5, 6, 7, 8, 9, 10): c.number_format = "0.##"
wsn.freeze_panes = "C3"

# ---------- 与计划口径对照 ----------
wsg = wb["功能视角分解"]
plan_feat = plan_common = 0.0
for rr in range(3, wsg.max_row + 1):
    a = str(wsg.cell(row=rr, column=1).value)
    v = wsg.cell(row=rr, column=14).value
    if a.startswith("F") and isinstance(v, (int, float)):
        plan_feat += v
    elif a.startswith("G") and isinstance(v, (int, float)):
        plan_common += v
grand = round(sum(gt), 2)
note = ("对照说明（两个口径）：① 本表为子任务级独立直估，合计 %s 人天（BA %s / Wade %s / DevB %s / 测试 %s / SLC %s），"
        "其中「存储落地Iceberg/UC」「元数据驱动」「RBAC」等设计实现类子任务与计划公共行（架构与非功能设计等）存在重叠，且未假设跨功能复用（同一界面/中间件多处计价）。"
        "② 排期计划（WBS分解/功能视角分解）中：26项功能直接分摊仅 %s 人天（大量开发工作量按阶段放在架构设计/集成/专项等公共行，公共合计 %s 人天），总投入185人天。"
        "结论：功能直估 %s 人天已达计划总投入的 %.0f%%，而计划中还有需求8.5/初始化9/PM15及集成·UAT·上线等纯公共工作未包含在本直估内——按传统粒度估算，3人×3个月偏紧，"
        "需依赖AI Native复用与生成提效吸收约50人天量级缺口，建议将本表作为功能规模基线并纳入风险跟踪。"
        % (grand, gt[0], gt[1], gt[2], gt[3], gt[4], round(plan_feat, 2), round(plan_common, 2), grand, grand / 185 * 100))
c = wsn.cell(row=r + 2, column=1, value=note)
c.font = Font(italic=True, size=9, color="C00000")
c.alignment = Alignment(wrap_text=True, vertical="top")
wsn.merge_cells(start_row=r + 2, start_column=1, end_row=r + 2, end_column=11)
wsn.row_dimensions[r + 2].height = 52

# ---------- 使用说明：幂等更新 + 清理历史重复行 ----------
ws0 = wb["使用说明"]
desc = ("子任务级直估：26项功能下钻到%d个子任务/工作包（如F01-01拆为原系统录入/DB连接/表勾选/保留期配置/Iceberg·UC存储落地/导入作业/监控），"
        "按BA/开发A/开发B/测试/SLC文档独立估人天，不锚定185人天；功能名称与功能小计为纵向合并单元格；底部附与计划口径的对照与差异说明" % sum(len(s) for _, s in B))
seen_feat, seen_break = 0, 0
for rr in range(1, ws0.max_row + 1):
    a = ws0.cell(row=rr, column=1).value
    if a == "功能拆解估算":
        if seen_break == 0:
            ws0.cell(row=rr, column=2, value=desc).alignment = Alignment(wrap_text=True, vertical="top")
            ws0.row_dimensions[rr].height = 30
        else:
            ws0.cell(row=rr, column=1).value = None; ws0.cell(row=rr, column=2).value = None
        seen_break += 1
    elif a == "功能视角分解":
        seen_feat += 1
        if seen_feat > 1:
            ws0.cell(row=rr, column=1).value = None; ws0.cell(row=rr, column=2).value = None
if seen_break == 0:
    last = ws0.max_row + 1
    ws0.cell(row=last, column=1, value="功能拆解估算").font = Font(bold=True)
    v = ws0.cell(row=last, column=2, value=desc)
    v.alignment = Alignment(wrap_text=True, vertical="top")
    ws0.row_dimensions[last].height = 30

wb.save(PATH)
print(f"完成：功能拆解估算（合并版式）{len(B)}功能 / {sum(len(s) for _, s in B)}子任务 | 直估合计 {grand} 人天")
print(f"使用说明去重：功能视角分解保留1行（清理{seen_feat-1}行）、功能拆解估算保留1行（清理{seen_break-1}行）")
