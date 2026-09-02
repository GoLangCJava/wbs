# -*- coding: utf-8 -*-
"""「功能视角分解」页（功能×角色矩阵）· V2.0：功能人天直接取自「功能拆解估算」直估，
公共支撑行挂新WBS组编码并与「WBS分解」V2.0动态对账。用法：python3 add_function_view.py
"""
import ast, datetime as dt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PATH = "/home/user/wbs/RIMS项目WBS工作分解_AI Native版.xlsx"
wb = openpyxl.load_workbook(PATH)

# ---------- WBS分解V2.0 叶子读取 ----------
ws = wb["WBS分解"]
LEAVES = []          # (code, owner, eff, end, name)
for r in range(4, ws.max_row + 1):
    code = ws.cell(r, 1).value
    if code and ws.cell(r, 3).value == 3:
        LEAVES.append((str(code), ws.cell(r, 5).value, ws.cell(r, 9).value or 0,
                       ws.cell(r, 7).value, str(ws.cell(r, 2).value or "")))

def bucket(owner, name):
    if owner in ("Wade", "DevB"):
        return owner
    if owner == "Mark":
        return "PM"
    if "SLC" in name or "会议纪要" in name or "复盘" in name:
        return "文档"
    if any(k in name for k in ("测试", "验证", "用例", "走查", "UAT", "核对", "阶段评审", "回归")):
        return "测试"
    return "BA"

def gsum(prefixes, b=None):
    return round(sum(eff for code, owner, eff, _e, name in LEAVES
                     if any(code == p or code.startswith(p + ".") for p in prefixes)
                     and (b is None or bucket(owner, name) == b)), 2)

def gend(prefixes):
    ds = [e for code, _o, _f, e, _n in LEAVES if e and any(code == p or code.startswith(p + ".") for p in prefixes)]
    d = max(ds)
    return d.date() if isinstance(d, dt.datetime) else d

# ---------- 功能清单（覆盖对照页） ----------
wst = wb["26项功能覆盖对照"]
FMAP = []
for r in range(3, wst.max_row + 1):
    fid = wst.cell(r, 1).value
    if fid and str(fid).startswith("F"):
        FMAP.append((str(fid), str(wst.cell(r, 5).value),
                     str(wst.cell(r, 7).value), str(wst.cell(r, 11).value)))
assert len(FMAP) == 26

# ---------- V2.1：功能行人天直接取自 WBS分解 排期值（直估×AI执行系数后） ----------
FEATSUM = {}
for f in FMAP:
    fc, _, grp, _ = f
    if grp.replace(".", "").isdigit():
        FEATSUM[fc] = dict(BA=gsum([grp], "BA"), Wade=gsum([grp], "Wade"), DevB=gsum([grp], "DevB"),
                           测试=gsum([grp], "测试"), 文档=gsum([grp], "文档"))

# ---------- 公共支撑（挂新WBS组编码） ----------
PUBLIC = [
 ("G1", "工程基座与部署流水线（N1＋N4-1/2/3）", {"Wade": ["1.1", "1.2"], "DevB": ["1.1", "1.2"], "测试": ["1.1", "1.2"], "文档": ["1.1"]},
  {"Wade": "仓库/规范/前端骨架/四道门禁CI-CD/部署流水线/K8s清单", "DevB": "后端骨架/集群对接/测试环境先期部署", "测试": "基座与门禁验证/测试环境验证", "文档": "工程基座SLC"}),
 ("G2", "全局非功能（N3）", {"Wade": ["4.1"], "DevB": ["4.1"], "BA": ["4.1"], "测试": ["4.1"]},
  {"Wade": "前端性能/鉴权掩码/可观测插桩", "DevB": "查询性能/鉴权掩码服务侧/可观测/容量验证", "BA": "鉴权掩码策略澄清", "测试": "非功能专项用例"}),
 ("G3", "质量治理（N2）", {"Wade": ["5.1"], "DevB": ["5.1"], "测试": ["5.1"], "文档": ["5.1"]},
  {"Wade": "ESLint/SAST/SCA/技术债A侧", "DevB": "SonarLint/SAST/SCA/技术债B侧", "测试": "治理抽检", "文档": "质量治理SLC"}),
 ("G4", "开发侧文档（N5）", {"Wade": ["6.1"], "DevB": ["6.1"], "文档": ["6.1"]},
  {"Wade": "架构与总体设计/应用接口文档", "DevB": "数据设计与ETL作业文档", "文档": "N5合稿评审"}),
 ("G5", "部署收尾（N4-4/5）", {"Wade": ["6.2"], "DevB": ["6.2"], "测试": ["6.2"], "文档": ["6.2"]},
  {"Wade": "上线值守与运维交接", "DevB": "部署预演与回退验证", "测试": "值守验证", "文档": "部署与运维SLC"}),
]
# 归属完整性：功能组编码 ∪ 公共编码 = 全部叶子
_feat_grp = {f[2] for f in FMAP if f[2].replace(".", "").isdigit()}
_pub = {c for g in PUBLIC for codes in g[2].values() for c in codes}
_all = {code.split(".")[0] + "." + code.split(".")[1] for code, *_ in LEAVES if code.count(".") >= 1}
_miss = {g for g in _all if g not in _feat_grp and g not in _pub and g.split(".")[0] not in _pub}
assert not _miss, f"未归属组: {_miss}"

WT = {
 "F01-01": ("迁移范围/映射规则澄清与验收标准", "迁移任务管理界面（调度/进度）", "SeaTunnel分片/限速/断点批量导入", "导入用例+样本表验证", "SLC文档"),
 "F01-02": ("文件来源与离线介质场景澄清", "迁移任务界面（文件批次）", "多源文件/大文件/属性保留/包解析", "导入与中断恢复用例", "SLC文档"),
 "F01-03": ("源→标准字段映射规则定义", "规则配置界面", "规则配置服务（映射/转换/校验）", "规则执行用例", "SLC文档"),
 "F01-04": ("完整性基准口径定义", "校验结果展示", "迁移前基准记录（行数/校验和）", "基准记录用例", "SLC文档"),
 "F01-05": ("比对差异处理口径", "比对结果展示", "迁移后比对（SHA-256/差异报告）", "比对用例（CC对账预演）", "SLC文档"),
 "F02-01": ("元数据字段体系与必填规则", "—", "统一元数据模型API", "模块2走查", "SLC文档"),
 "F02-02": ("来源系统编码与绑定规则", "元数据管理界面（协作）", "来源系统信息绑定API", "模块2走查", "SLC文档"),
 "F03-01": ("保留期矩阵与到期处置策略", "保留策略配置界面", "处置引擎（扫描→DROP PARTITION→VACUUM）", "到期删除走查", "SLC文档"),
 "F03-02": ("Legal Hold审批流程定义", "Legal Hold冻结/解除界面", "Hold/Release+豁免+留痕", "冻结/解除走查", "SLC文档"),
 "F04-01": ("密级与加密合规要求", "—", "ADLS原生SSE/CMK配置+PII字段级加密", "加密落盘验证", "SLC文档"),
 "F04-02": ("密钥保管与审计要求", "—", "Key Vault原生(RBAC)+应用Managed Identity集成", "密钥轮换用例", "SLC文档"),
 "F04-03": ("备份策略要求（RPO）", "—", "SQL PITR/ADLS原生备份策略配置", "备份恢复用例", "SLC文档"),
 "F04-04": ("灾备要求（RTO）", "—", "Azure Blob原生冗余(GRS)配置+恢复演练", "灾备切换用例", "SLC文档"),
 "F05-01": ("检索需求与查询配置模型", "检索页动态表单+关键词检索", "—", "检索用例", "SLC文档"),
 "F05-02": ("来源系统筛选需求", "按系统名称检索", "—", "筛选用例", "SLC文档"),
 "F05-03": ("时间维度检索需求", "按业务时间检索", "—", "时间范围用例", "SLC文档"),
 "F05-04": ("列表字段/排序/导出需求", "结果列表（分页/排序）+流式导出", "—", "大批量用例", "SLC文档"),
 "F05-05": ("预览格式与权限要求", "在线预览（SAS直连）+大文件优化", "—", "预览用例", "SLC文档"),
 "F06-01": ("审计合规要求（归档日志）", "归档操作日志+审计查询界面", "—", "审计走查", "SLC文档"),
 "F06-02": ("查询行为审计要求", "查询操作日志记录", "—", "审计走查", "SLC文档"),
 "F06-03": ("导出审计要求（范围/IP）", "导出操作日志记录", "—", "审计走查", "SLC文档"),
 "F08-01": ("角色权限矩阵（RBAC）定义", "角色权限配置（界面+API）", "—", "越权用例", "SLC文档"),
 "F08-02": ("账号生命周期策略", "用户账号管理（界面+API）", "—", "账号用例", "SLC文档"),
 "F08-03": ("SSO对接需求（Entra ID）", "SSO集成（Entra ID+MSAL中间件）", "—", "SSO登录用例", "SLC文档"),
 "F08-04": ("身份映射规则", "身份自动识别与关联", "—", "身份关联用例", "SLC文档"),
 "F08-05": ("异常访问控制策略", "超时/锁定/失效控制", "—", "异常登录用例", "SLC文档"),
}
# ---------- 样式 ----------
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NAVY = "305496"; L1FILL = "D9E1F2"; GRAY = "F2F2F2"
F_TITLE = Font(bold=True, size=14, color="FFFFFF")
F_HDR = Font(bold=True, color="FFFFFF")
F_L1 = Font(bold=True)
FILL_NAVY = PatternFill("solid", fgColor=NAVY)
FILL_L1 = PatternFill("solid", fgColor=L1FILL)
FILL_G = PatternFill("solid", fgColor=GRAY)
CENTER = Alignment(horizontal="center", vertical="center")

if "功能视角分解" in wb.sheetnames:
    del wb["功能视角分解"]
wsg = wb.create_sheet("功能视角分解", wb.sheetnames.index("26项功能覆盖对照") + 1)
wsg.sheet_properties.tabColor = "C00000"
wsg.merge_cells("A1:P1")
wsg["A1"] = "功能视角分解（功能×角色矩阵）· V3.0：每个功能由谁做什么、投入多少人天 — 与「功能拆解估算」「WBS分解」完全对账（BA12.5/Wade60/DevB70.5/测试38.75/SLC13.75＝195.5）"
wsg["A1"].font = F_TITLE; wsg["A1"].fill = FILL_NAVY; wsg["A1"].alignment = CENTER
wsg.row_dimensions[1].height = 26
GH = ["功能编号", "模块", "具体功能（与源文件一致）",
      "BA·需求（Mandy）", "BA人天", "开发A·Wade", "A人天", "开发B·DevB", "B人天",
      "测试·验证（Mandy）", "测试人天", "文档（Mandy）", "文档人天", "小计人天", "关联WBS任务", "计划完成"]
GW = [8, 13, 20, 24, 7, 24, 7, 24, 7, 22, 7, 14, 7, 8, 16, 12]
for i, (h, w) in enumerate(zip(GH, GW), 1):
    wsg.column_dimensions[get_column_letter(i)].width = w
    c = wsg.cell(row=2, column=i, value=h); c.font = F_HDR; c.fill = FILL_NAVY; c.alignment = CENTER; c.border = BORDER
MOD_NAME = {"F01": "模块1·存量迁移", "F02": "模块2·元数据", "F03": "模块3·生命周期处置", "F04": "模块4·存储安全",
            "F05": "模块5·检索利用", "F06": "模块6·审计合规", "F08": "模块8·全局治理"}
_mod_fills = {"F01": "EDEDF7", "F02": "E2EFDA", "F03": "FFF2CC", "F04": "FCE4D6", "F05": "DDEBF7", "F06": "E7E6E6", "F08": "E4DFEC"}
rg = 3
tot_cols = {"BA": 0.0, "Wade": 0.0, "DevB": 0.0, "测试": 0.0, "文档": 0.0, "PM": 0.0}
for fcode, fname, main_w, plan_txt in FMAP:
    ev = {"BA": 0.0, "Wade": 0.0, "DevB": 0.0, "测试": 0.0, "文档": 0.0}
    if fcode in FEATSUM:
        ev.update(FEATSUM[fcode])
    vals = [fcode, MOD_NAME[fcode[:3]], fname,
            WT[fcode][0], ev["BA"] or None, WT[fcode][1], ev["Wade"] or None, WT[fcode][2], ev["DevB"] or None,
            WT[fcode][3], ev["测试"] or None, WT[fcode][4], ev["文档"] or None,
            f"=ROUND(E{rg}+G{rg}+I{rg}+K{rg}+M{rg},2)",
            main_w, plan_txt]
    fill = PatternFill("solid", fgColor=_mod_fills[fcode[:3]])
    for i, v in enumerate(vals, 1):
        c = wsg.cell(row=rg, column=i, value=v); c.border = BORDER
        c.alignment = CENTER if i in (1, 5, 7, 9, 11, 13, 14, 16) else Alignment(wrap_text=True, vertical="center")
        c.fill = fill
        if i in (5, 7, 9, 11, 13, 14): c.number_format = "0.##"
    for k in ("BA", "Wade", "DevB", "测试", "文档"):
        tot_cols[k] += ev[k]
    wsg.row_dimensions[rg].height = 34
    rg += 1
for gid, gname, cols, desc in PUBLIC:
    sums = {k: gsum(codes) if k == "PM" else gsum(codes, k) for k, codes in cols.items() if k != "PM"}
    sums["PM"] = gsum(cols.get("PM", []), "PM")
    all_ps = sorted({c for codes in cols.values() for c in codes})
    vals = [gid, "公共支撑", gname,
            desc.get("BA", "—"), sums.get("BA") or None, desc.get("Wade", "—"), sums.get("Wade") or None,
            desc.get("DevB", "—"), sums.get("DevB") or None, desc.get("测试", "—"), sums.get("测试") or None,
            desc.get("文档", "—"), sums.get("文档") or None, desc.get("PM", "—"),
            (f"=ROUND(E{rg}+G{rg}+I{rg}+K{rg}+M{rg}+{sums.get('PM', 0)},2)" if sums.get("PM") else f"=ROUND(E{rg}+G{rg}+I{rg}+K{rg}+M{rg},2)"),
            ",".join(all_ps[:4]) + ("…" if len(all_ps) > 4 else ""),
            gend(all_ps).strftime("%m/%d")]
    for i, v in enumerate(vals, 1):
        c = wsg.cell(row=rg, column=i, value=v); c.border = BORDER; c.fill = FILL_G
        c.alignment = CENTER if i in (1, 5, 7, 9, 11, 13, 14, 15, 16) else Alignment(wrap_text=True, vertical="center")
        if i in (5, 7, 9, 11, 13, 14, 15): c.number_format = "0.##"
    for k in ("BA", "Wade", "DevB", "测试", "文档", "PM"):
        tot_cols[k] += sums.get(k, 0)
    wsg.row_dimensions[rg].height = 34
    rg += 1
grand = round(sum(tot_cols[k] for k in ("BA", "Wade", "DevB", "测试", "文档")) + tot_cols["PM"], 2)
mandy_tot = round(tot_cols["BA"] + tot_cols["测试"] + tot_cols["文档"], 2)
_last = rg - 1
labels = ["合计", "", "26项功能 + 9类公共支撑",
          "Mandy·BA", "=ROUND(SUM(E3:E%d),2)" % _last, "Wade（开发A）", "=ROUND(SUM(G3:G%d),2)" % _last,
          "DevB（开发B）", "=ROUND(SUM(I3:I%d),2)" % _last,
          "Mandy·测试", "=ROUND(SUM(K3:K%d),2)" % _last, "Mandy·文档", "=ROUND(SUM(M3:M%d),2)" % _last,
          "=ROUND(SUM(N3:N%d),2)" % _last, "", "PM Mark " + str(tot_cols["PM"])]
for i, v in enumerate(labels, 1):
    c = wsg.cell(row=rg, column=i, value=v); c.border = BORDER; c.fill = FILL_L1; c.font = F_L1
    c.alignment = CENTER
    if i in (5, 7, 9, 11, 13, 14): c.number_format = "0.##"
wsg.freeze_panes = "D3"
wsg.auto_filter.ref = "A2:P" + str(rg)
note = wsg.cell(row=rg + 1, column=1,
    value="V3.0对账：Mandy三职 %s（BA %s + 测试 %s + 文档 %s）+ Wade %s + DevB %s = %s人天，与「功能拆解估算」「WBS分解」完全一致（195.5）；"
          "人天=拆解估算直估原值；开发WBS交付至2026/12/29（M1框架9/17·M2批次1 11/17·M3批次2 12/18·M4质量12/29·M5完成12/29）"
          % (mandy_tot, tot_cols["BA"], tot_cols["测试"], tot_cols["文档"], tot_cols["Wade"], tot_cols["DevB"], grand))
note.font = Font(italic=True, size=9, color="C00000")

# 与WBS分解动态对账（强断言）
wsx = wb["WBS分解"]
wbs_tot = {}
for r in range(4, wsx.max_row + 1):
    if wsx.cell(r, 3).value == 3 and isinstance(wsx.cell(r, 9).value, (int, float)):
        o = wsx.cell(r, 5).value
        wbs_tot[o] = round(wbs_tot.get(o, 0) + wsx.cell(r, 9).value, 2)
assert round(tot_cols["Wade"],2) == wbs_tot.get("Wade") and round(tot_cols["DevB"],2) == wbs_tot.get("DevB"), (tot_cols, wbs_tot)
if mandy_tot != wbs_tot.get("Mandy"):
    _cap = set()
    for _g in PUBLIC:
        for _b, _ps in _g[2].items():
            for _p in _ps:
                for code, owner, eff, _e, name in LEAVES:
                    if owner == "Mandy" and (code == _p or code.startswith(_p + ".")) and bucket(owner, name) == _b:
                        _cap.add(code)
    _lost = [(c, n[:26], e) for c, o, e, _x, n in LEAVES if o == "Mandy" and c not in _cap]
    raise AssertionError(f"Mandy对账失败 {mandy_tot} vs {wbs_tot.get('Mandy')}，漏捕: {_lost}")
assert round(tot_cols["PM"], 2) == wbs_tot.get("Mark", 0) == 0, tot_cols
assert grand == round(sum(wbs_tot.values()), 2) == 195.5, (grand, wbs_tot)
assert tot_cols["Wade"] == 60.0 and tot_cols["DevB"] == 70.5 and mandy_tot == 65.0, tot_cols

# 使用说明（幂等覆盖）
ws0 = wb["使用说明"]
seen = False
for rr in range(1, ws0.max_row + 1):
    if ws0.cell(row=rr, column=1).value == "功能视角分解":
        ws0.cell(row=rr, column=2).value = ("按功能维度查看：26项功能×（BA/开发A/开发B/测试/文档）五类角色各自的工作内容与人天（=「功能拆解估算」直估原值），"
                                            "另有5类公共支撑行（工程基座/全局非功能/质量治理/开发侧文档/部署收尾）；底部与WBS分解V3.0合计195.5人天自动对账")
        seen = True
if not seen:
    last = ws0.max_row + 1
    k = ws0.cell(row=last, column=1, value="功能视角分解"); k.font = Font(bold=True)
    v = ws0.cell(row=last, column=2, value="按功能维度查看：26项功能×五类角色人天（V2.0直估）+9类公共支撑行；与WBS分解V2.0合计297.5人天对账")
    v.alignment = Alignment(wrap_text=True, vertical="top")

wb.save(PATH)
print(f"完成：功能视角分解V2.0 | Mandy {mandy_tot}（BA {tot_cols['BA']}/测试 {tot_cols['测试']}/文档 {tot_cols['文档']}）"
      f"Wade {tot_cols['Wade']} DevB {tot_cols['DevB']} PM {tot_cols['PM']} = {grand}人天 ✓")
