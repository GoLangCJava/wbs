# -*- coding: utf-8 -*-
"""
生成《One-Abbott 共享应用平台 · 技术架构图集》PPTX（精简版 · 2 页）
  P1 总体技术架构 —— 能力复用视角（Build Once, Reuse Everywhere）
  P2 微前端架构设计 —— 复用与独立交付
绘图助手与视觉体系复用 gen_arch_ppt.py（PPT 原生形状，不嵌图片）。
内容基于 archi.md（§6 应用架构：当前模式 vs 目标模式 / 微前端）与 tech-architecture.md。
"""
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE

from gen_arch_ppt import (
    INK, SUB, MUT, FAINT, BORDER, BG,
    AZ, AZD, GRN, GRND, PUR, PURD, CYN, CYND, ORG, AMB, AMBD, ROS, SLA,
    L_BLUE, B_BLUE, L_GRN, B_GRN, L_PUR, B_PUR, L_CYN, B_CYN,
    L_AMB, B_AMB, L_ROS, B_ROS, L_GRNB, B_GRNB, L_PURB, B_PURB, L_CYNB, B_CYNB,
    rect, oval, seg, poly, txt, chip, title_bar,
    OPS, _cur,
)

DASH = MSO_LINE_DASH_STYLE.DASH
OUT_NAME = 'One-Abbott_共享应用平台_技术架构图集.pptx'


# ==================================================== P1 · 总体架构·复用视角 =
def slide_reuse(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6]); _cur['i'] = 0
    sl.background.fill.solid(); sl.background.fill.fore_color.rgb = RGBColor.from_string(BG)
    title_bar(sl, '总体技术架构 · Capability First —— Build Once, Reuse Everywhere',
              '复用三形态：API · 事件 · 微前端模块\n多租户：ADC / EPD / MD / APOC → 平台租户', AZ)

    LX, LW = 0.22, 8.34          # 左主列
    RX, RW = 8.66, 4.45          # 右复用面板

    # ---------- 01 渠道层 ----------
    rect(sl, LX, 0.62, LW, 0.78, L_GRN, B_GRN, 1.2, 0.08)
    txt(sl, 0.40, 0.69, 4.6, 0.17, '01 渠道层 · 同一能力 · 多渠道触达', sz=8.5, b=True, c=GRND)
    chs = [('微信公众号 · H5', '内容 / 会议 / 积分'), ('微信小程序 · Taro', '内容 / 签到 / AI'),
           ('企业微信', '侧边栏 · 拜访 / C360'), ('App · RN', '离线拜访 / 工单'), ('Portal', '运营后台')]
    xx = 0.40
    for t, s in chs:
        chip(sl, xx, 0.92, 1.55, 0.40, t, s, 'FFFFFF', 'BBF7D0', 1.0, 0.05, tsz=6.8, sc=MUT, ssz=5.8)
        xx += 1.615
    for ax in (2.60, 6.00):
        seg(sl, ax, 1.40, ax, 1.50, GRN, 1.8)
    txt(sl, 2.70, 1.395, 1.2, 0.11, 'HTTPS · SSO', sz=5.6, b=True, c=GRN)

    # ---------- 02 接入边缘 ----------
    rect(sl, LX, 1.50, LW, 0.50, 'F8FAFC', 'CBD5E1', 1.2, 0.07)
    txt(sl, 0.36, 1.61, 1.10, 0.28, '02 接入边缘', sz=7.6, b=True, c=SUB, anchor='m')
    chip(sl, 1.52, 1.61, 1.05, 0.28, 'CDN ＋ WAF', None, 'FFFFFF', 'CBD5E1', 1.0, 0.04, tc=INK, tsz=6.4)
    chip(sl, 2.63, 1.61, 1.50, 0.28, '外部 APIM · 公网 API', None, 'FFFFFF', AZ, 1.2, 0.04, tc=AZD, tsz=6.4)
    chip(sl, 4.19, 1.61, 1.60, 0.28, 'Entra ID · 内部 SSO', None, 'FFFFFF', B_BLUE, 1.0, 0.04, tc=AZD, tsz=6.4)
    chip(sl, 5.85, 1.61, 2.50, 0.28, 'identity-service 微信 / 验证码 → OneID', None,
         'FFFFFF', B_BLUE, 1.0, 0.04, tc=AZD, tsz=6.2)
    for ax in (2.60, 6.00):
        seg(sl, ax, 2.00, ax, 2.10, CYN, 1.8)

    # ---------- 03 BFF ----------
    rect(sl, LX, 2.10, LW, 0.58, L_CYN, B_CYNB, 1.2, 0.08)
    txt(sl, 0.40, 2.17, 4.6, 0.15, '03 BFF 聚合层 · 只做聚合裁剪 · 不重复实现业务', sz=8.2, b=True, c=CYND)
    bfs = ['hcp-bff · 小程序 / H5', 'wework-bff · 企微侧边栏', 'app-bff · 离线同步', 'portal-bff · 运营门户']
    xx = 0.40
    for s in bfs:
        chip(sl, xx, 2.36, 1.90, 0.26, s, None, 'FFFFFF', CYN, 1.1, 0.04, tc=INK, tsz=6.6)
        xx += 1.97
    for ax in (2.60, 6.00):
        seg(sl, ax, 2.68, ax, 2.78, CYN, 1.8)
    txt(sl, 6.10, 2.675, 1.0, 0.11, '聚合调用', sz=5.6, b=True, c=CYN)

    # ---------- 04 能力层 ----------
    rect(sl, LX, 2.78, LW, 1.88, 'FFFFFF', BORDER, 1.2, 0.08)
    txt(sl, 0.40, 2.85, 5.10, 0.15, '04 能力层 · Capability Hub —— 一次建设（Build Once）', sz=8.5, b=True, c='334155')
    txt(sl, 5.60, 2.86, 2.84, 0.13, '一 Hub 一限界上下文 · Database-per-Service', sz=5.9, c=FAINT, align='r')
    rect(sl, 0.36, 2.98, 0.30, 1.40, L_BLUE, B_BLUE, 1.2, 0.05)
    txt(sl, 0.36, 2.98, 0.30, 1.40, '内部 APIM', sz=6.0, b=True, c=AZD, align='c', anchor='m', vert=True)
    seg(sl, 0.66, 3.31, 0.72, 3.31, AZ, 1.3)
    seg(sl, 0.66, 4.05, 0.72, 4.05, AZ, 1.3)
    hubs = [
        ('Content Hub', '内容 · 学术 · DAM · 搜索', '复用：API · 事件 · MF', L_BLUE, B_BLUE),
        ('Event Hub', '会议 · 报名签到 · 积分', '复用：API · 事件 · MF', L_GRN, B_GRN),
        ('Customer Hub', 'HCP / HCO · OneID · Consent', '复用：API · 事件', L_CYN, B_CYN),
        ('Sales Hub', '拜访 · 打卡 · 离线同步', '复用：API · 事件', L_AMB, B_AMB),
    ]
    for i, (t, s1, s2, lf, bl) in enumerate(hubs):
        hx = 0.74 + i * 1.92
        rect(sl, hx, 2.98, 1.85, 0.66, lf, bl, 1.2, 0.05, shadow=True)
        txt(sl, hx + 0.05, 3.025, 1.75, 0.14, t, sz=7.2, b=True, c=INK, align='c')
        txt(sl, hx + 0.05, 3.205, 1.75, 0.13, s1, sz=5.7, c=MUT, align='c')
        txt(sl, hx + 0.05, 3.40, 1.75, 0.13, s2, sz=5.6, c=AZD, b=True, align='c')
    hubs2 = [
        ('Service Hub', '工单 · SLA · 备件 · IoT', '复用：API · 事件 · MF', 'F6F8FA', 'CBD5E1'),
        ('Plugin Hub', '插件 API · 租户级启用', '复用：API · 事件 · MF', 'FDF2F8', 'F9A8D4'),
        ('横切服务', '身份 · 通知 · 工作流 · 规则 · 审计 · 文件', '所有 Hub 共用的技术能力 · API / SDK 复用', 'FFF7ED', 'FED7AA'),
    ]
    for i, (t, s1, s2, lf, bl) in enumerate(hubs2):
        hx = 0.74 + i * 2.58
        rect(sl, hx, 3.72, 2.50, 0.66, lf, bl, 1.2, 0.05, shadow=True)
        txt(sl, hx + 0.05, 3.765, 2.40, 0.14, t, sz=7.2, b=True, c=INK, align='c')
        txt(sl, hx + 0.05, 3.945, 2.40, 0.13, s1, sz=5.7, c=MUT, align='c')
        txt(sl, hx + 0.05, 4.14, 2.40, 0.13, s2, sz=5.6, c=AZD, b=True, align='c')
    txt(sl, 0.40, 4.42, 8.00, 0.15,
        '消费三形态：① API 复用（BFF / 应用 → 内部 APIM · OpenAPI 契约 · 版本并存）　② 事件复用（订阅领域事件 · 解耦集成）　③ 模块复用（MF 远程模块多端挂载 → 详见下页）',
        sz=6.2, b=True, c=AMBD, align='c')
    for ax in (2.60, 6.00):
        seg(sl, ax, 4.66, ax, 4.76, AMB, 1.8)
    txt(sl, 2.70, 4.655, 1.2, 0.11, '事件 / CDC', sz=5.6, b=True, c=AMB)

    # ---------- 05 事件骨干 ----------
    rect(sl, LX, 4.76, LW, 0.46, 'FFF7ED', 'FED7AA', 1.2, 0.07)
    chip(sl, 0.36, 4.85, 0.94, 0.28, '05 事件骨干', None, 'FFFFFF', ORG, 1.2, 0.04, tc=AMBD, tsz=6.6)
    chip(sl, 1.38, 4.85, 3.44, 0.28, 'Service Bus · 领域事件', 'CloudEvents · Outbox · 幂等 / DLQ',
         'FFFFFF', ORG, 1.1, 0.04, tc=INK, tsz=6.4, sc=MUT, ssz=5.8)
    chip(sl, 4.90, 4.85, 3.44, 0.28, 'Event Hub · 行为流', '埋点 → 实时标签 → 湖仓',
         'FFFFFF', ORG, 1.1, 0.04, tc=INK, tsz=6.4, sc=MUT, ssz=5.8)
    for ax in (2.60, 6.00):
        seg(sl, ax, 5.22, ax, 5.32, AZ, 1.8)

    # ---------- 06 数据与智能底座 ----------
    rect(sl, LX, 5.32, LW, 0.92, L_BLUE, B_BLUE, 1.2, 0.08)
    txt(sl, 0.40, 5.39, 6.0, 0.15, '06 数据与智能底座 · Data / AI / Cloud —— 数据与 AI 同样只建一次', sz=8.2, b=True, c='1D4ED8')
    bases = [('HCP Pool', 'OneID · Consent · C360'), ('湖仓 · Fabric', 'B→S→G · C360 宽表'),
             ('AI 平台', '模型网关 · RAG · Copilot'), ('Landing Zone', 'AKS · GitOps · OTel')]
    xx = 0.40
    for t, s in bases:
        chip(sl, xx, 5.60, 1.95, 0.50, t, s, 'FFFFFF', '93C5FD', 1.1, 0.05, tc=INK, tsz=7.0, sc=MUT, ssz=5.8)
        xx += 2.02

    # ================= 右侧 · 复用面板 =================
    # ---- A 复用前后对比 ----
    rect(sl, RX, 0.62, RW, 1.90, 'FFFFFF', BORDER, 1.2, 0.07, shadow=True)
    txt(sl, RX + 0.12, 0.70, RW - 0.24, 0.16, '从重复建设到平台复用（archi.md §6）', sz=8.2, b=True, c=INK)
    rect(sl, RX + 0.12, 0.94, 1.86, 1.44, L_ROS, 'F87171', 1.1, 0.05, dash=DASH)
    txt(sl, RX + 0.22, 1.00, 1.66, 0.14, '当前 · 各自建设 ✗', sz=7.0, b=True, c='B91C1C')
    for i, s in enumerate(['ADC / EPD / MD / APOC', '各建 内容·会议·积分', '≈ 12 套重复实现', '数据孤岛 · 体验割裂']):
        txt(sl, RX + 0.22, 1.20 + i * 0.24, 1.68, 0.20, s, sz=5.9, c=SUB)
    seg(sl, RX + 2.02, 1.66, RX + 2.32, 1.66, MUT, 2.0)
    rect(sl, RX + 2.36, 0.94, 1.97, 1.44, L_GRN, B_GRN, 1.2, 0.05)
    txt(sl, RX + 2.46, 1.00, 1.77, 0.14, '目标 · 平台统一 ✓', sz=7.0, b=True, c=GRND)
    for i, s in enumerate(['Capability Hub 1 套能力', '应用 = 租户按需组合', '能力 / 数据 / AI 三重复用', 'Build Once · Reuse Everywhere']):
        txt(sl, RX + 2.46, 1.20 + i * 0.24, 1.79, 0.20, s, sz=5.9, c=SUB, b=(i == 3))

    # ---- B 复用矩阵 ----
    rect(sl, RX, 2.62, RW, 2.70, 'FFFFFF', BORDER, 1.2, 0.07, shadow=True)
    txt(sl, RX + 0.12, 2.70, RW - 0.24, 0.16, '能力复用矩阵 · 一次建设 → 处处消费', sz=8.2, b=True, c=INK)
    apps = [('小程序', 'H5'), ('企微', ''), ('App', ''), ('Portal', ''), ('插件', '')]
    cols_x = [RX + 1.32 + i * 0.60 for i in range(5)]
    for i, (a1, a2) in enumerate(apps):
        paras = [(a1, {'sz': 5.6, 'b': True, 'c': SUB})]
        if a2:
            paras.append((a2, {'sz': 5.6, 'b': True, 'c': SUB}))
        txt(sl, cols_x[i] - 0.28, 2.92, 0.56, 0.34, paras, align='c', anchor='m', leading=1.05)
    seg(sl, RX + 0.12, 3.26, RX + RW - 0.12, 3.26, BORDER, 0.8, arrow=False)
    # ● 复用  ○ 按需裁剪  – 不涉及
    matrix = [
        ('内容学术', [2, 2, 1, 2, 2]), ('会议积分', [2, 1, 1, 2, 1]), ('客户 360', [1, 2, 2, 2, 2]),
        ('拜访绩效', [0, 2, 2, 2, 0]), ('报修服务', [0, 0, 2, 2, 0]), ('AI 能力', [2, 2, 1, 2, 2]),
    ]
    for r, (cap, marks) in enumerate(matrix):
        ry = 3.32 + r * 0.29
        txt(sl, RX + 0.12, ry, 1.14, 0.24, cap, sz=6.2, c=INK, anchor='m')
        for c, m in enumerate(marks):
            cx = cols_x[c]
            if m == 2:
                oval(sl, cx - 0.05, ry + 0.07, 0.10, 0.10, GRN)
            elif m == 1:
                oval(sl, cx - 0.05, ry + 0.07, 0.10, 0.10, 'FFFFFF', GRN, 1.0)
            else:
                txt(sl, cx - 0.10, ry, 0.20, 0.24, '–', sz=6.0, c=FAINT, align='c', anchor='m')
    txt(sl, RX + 0.12, 5.06, RW - 0.24, 0.14,
        '● 直接复用　○ 按需裁剪　– 不涉及 ｜ 例：新插件「术后随访」= C360 ＋ 内容 ＋ 积分 ＋ 通知，零新建',
        sz=5.8, b=True, c=GRND)

    # ---- C 租户级复用 ----
    rect(sl, RX, 5.42, RW, 0.82, L_BLUE, B_BLUE, 1.1, 0.06)
    txt(sl, RX + 0.12, 5.49, RW - 0.24, 0.14, '租户级复用 · 治理与计量', sz=7.8, b=True, c='1D4ED8')
    ten = ['tid 全链路 · SQL 行级安全', 'feature flag 按租户启用能力', 'APIM 配额 / 限流', '成本按 tid 计量分摊']
    for i, s in enumerate(ten):
        chip(sl, RX + 0.12 + (i % 2) * 2.12, 5.66 + (i // 2) * 0.26, 2.04, 0.22, s, None,
             'FFFFFF', '93C5FD', 1.0, 0.04, tc=SUB, tsz=5.9)

    # ---- 底部结论 ----
    txt(sl, 0.22, 6.36, 12.89, 0.16,
        '复用成效目标：12 套重复建设 → 1 套平台能力 ｜ 新渠道 / 新应用接入周期减半（只建 BFF 与视图）｜ 新场景 = 组合已有能力 ｜ 分层对应 archi.md：Experience → Application → Capability → Data → AI → Cloud',
        sz=6.4, c=MUT, align='c')


# ============================================ P2 · 微前端架构设计（如何设计） =
def slide_mfe(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6]); _cur['i'] = 1
    sl.background.fill.solid(); sl.background.fill.fore_color.rgb = RGBColor.from_string(BG)
    title_bar(sl, '微前端架构设计 · 复用与独立交付（Module Federation）',
              'React 18 · TS 5 · Rspack / Vite · Taro 小程序\n独立开发 · 独立部署 · 独立灰度 · 共享治理', CYN)

    # ================= ① 分层设计 =================
    rect(sl, 0.22, 0.62, 12.89, 2.14, 'FFFFFF', BORDER, 1.2, 0.08, shadow=True)
    txt(sl, 0.36, 0.70, 6.0, 0.16, '① 分层设计 · 四层结构（谁负责什么）', sz=8.5, b=True, c=INK)
    # L1 Shell
    chip(sl, 0.36, 0.94, 1.36, 0.44, [('Shell · 宿主', {'sz': 7.0, 'b': True, 'c': CYND}),
                                      ('平台团队统一维护', {'sz': 5.4, 'c': MUT})],
         None, L_CYN, B_CYNB, 1.2, 0.05)
    l1 = [('路由 / 布局', 'RBAC 菜单下发'), ('SSO 会话', '全局权限 SDK'), ('Design Tokens', '主题 / 暗色'),
          ('事件总线', 'mitt 跨模块通信'), ('MF 注册表', 'manifest 版本 / 灰度'), ('错误兜底', '失败降级占位')]
    xx = 1.78
    for t, s in l1:
        chip(sl, xx, 0.94, 1.78, 0.44, [(t, {'sz': 6.8, 'b': True, 'c': INK}), (s, {'sz': 5.6, 'c': MUT})],
             None, 'FFFFFF', BORDER, 1.1, 0.05)
        xx += 1.86
    seg(sl, 6.70, 1.38, 6.70, 1.50, CYN, 1.8)
    txt(sl, 6.82, 1.385, 2.4, 0.11, '运行时远程加载（不是构建时打包）', sz=5.6, b=True, c=CYN)
    # L2 Remote MF
    chip(sl, 0.36, 1.50, 1.36, 0.52, [('Remote MF', {'sz': 7.0, 'b': True, 'c': PURD}),
                                      ('业务团队自治', {'sz': 5.4, 'c': MUT})],
         None, L_PUR, B_PUR, 1.2, 0.05)
    mfs = [('产品', 'product'), ('内容学术', 'content'), ('会议', 'event'), ('销售', 'sales'),
           ('售后', 'service'), ('客户', 'customer'), ('AI 助手', 'ai'), ('插件', 'plugin')]
    xx = 1.78
    for t, s in mfs:
        chip(sl, xx, 1.50, 1.30, 0.52, [(t, {'sz': 6.8, 'b': True, 'c': INK}), (s + '-mf', {'sz': 5.6, 'c': MUT})],
             None, 'FFFFFF', B_PUR, 1.1, 0.05)
        xx += 1.37
    # L3 Shared
    chip(sl, 0.36, 2.08, 1.36, 0.30, [('Shared · 单例', {'sz': 6.6, 'b': True, 'c': AMBD})],
         None, L_AMB, B_AMB, 1.2, 0.05)
    chip(sl, 1.78, 2.08, 11.08, 0.30,
         '白名单：react ^18 · design-system · auth-sdk · api-client · i18n · telemetry ｜ 白名单外一律独立打包 · singleton 版本协商（防版本地狱）',
         None, 'FFFFFF', 'FDE68A', 1.1, 0.04, tc=SUB, tsz=6.2)
    # L4 Contract
    chip(sl, 0.36, 2.42, 1.36, 0.26, [('Contract', {'sz': 6.6, 'b': True, 'c': GRND})],
         None, L_GRN, B_GRN, 1.2, 0.05)
    chip(sl, 1.78, 2.42, 11.08, 0.26,
         '@abbott/contracts：API DTO · 领域事件 · 权限码 —— 前后端同源引用 · CI 兼容性校验 · 破坏性变更即阻断合并',
         None, 'FFFFFF', 'BBF7D0', 1.1, 0.04, tc=SUB, tsz=6.2)

    # ================= ② 运行时加载 / ③ 通信隔离 =================
    rect(sl, 0.22, 2.88, 6.30, 1.54, 'FFFFFF', BORDER, 1.2, 0.07, shadow=True)
    txt(sl, 0.36, 2.96, 5.9, 0.15, '② 运行时加载与版本协商（怎么跑起来）', sz=8.2, b=True, c=INK)
    steps = [('① Shell 启动', '引导 / 鉴权'), ('② 拉取 manifest', '版本 ＋ 灰度比例'), ('③ 解析远程入口', 'CDN URL'),
             ('④ 加载 MF', 'import remote'), ('⑤ 版本协商', 'react singleton'), ('⑥ 挂载渲染', '路由 / 权限校验')]
    for i, (t, s) in enumerate(steps):
        sx = 0.36 + (i % 3) * 1.98
        sy = 3.16 + (i // 3) * 0.52
        chip(sl, sx, sy, 1.90, 0.42, [(t, {'sz': 6.6, 'b': True, 'c': INK}), (s, {'sz': 5.6, 'c': MUT})],
             None, L_CYN, B_CYNB, 1.1, 0.05)
        if i % 3 < 2:
            seg(sl, sx + 1.90, sy + 0.21, sx + 1.98, sy + 0.21, CYN, 1.4)
    txt(sl, 0.36, 4.20, 6.02, 0.14, '失败降级：模块加载失败 → 占位提示不白屏 · 菜单按 RBAC 动态下发',
        sz=6.0, b=True, c=CYND, align='c')

    rect(sl, 6.66, 2.88, 6.45, 1.54, 'FFFFFF', BORDER, 1.2, 0.07, shadow=True)
    txt(sl, 6.80, 2.96, 6.0, 0.15, '③ 跨模块通信与隔离规则（怎么不腐化）', sz=8.2, b=True, c=INK)
    rules = [
        ('事件总线', 'mf:<module>:<event>（例 content.published → 积分刷新 · 通知）', PUR),
        ('禁止直连', '跨 MF 直接 import 一律禁止 · CI 扫描阻断', ROS),
        ('样式隔离', 'Design Tokens ＋ 模块级 class 前缀 · 全局样式仅 Shell 注入一次', AZ),
        ('跳转协议', '跨模块跳转走 Shell 统一路由 /mfe/<module>/…', CYN),
        ('全局一致', 'Loading / 错误 / 主题 / i18n 由 Shell 统一，保证体验不割裂', GRN),
    ]
    for i, (tag, s, c) in enumerate(rules):
        txt(sl, 6.80, 3.14 + i * 0.25, 6.18, 0.22,
            [dict(runs=[(tag + '：', {'c': c, 'sz': 6.3, 'b': True}), (s, {'c': SUB, 'sz': 6.3})])])

    # ================= ④ 一次开发多端挂载 =================
    rect(sl, 0.22, 4.52, 6.30, 1.74, 'FFFFFF', BORDER, 1.2, 0.07, shadow=True)
    txt(sl, 0.36, 4.60, 6.0, 0.15, '④ 复用：一次开发 · 多端挂载（以 content-mf 为例）', sz=8.2, b=True, c=INK)
    rect(sl, 0.36, 4.82, 2.36, 0.72, L_PUR, B_PUR, 1.3, 0.06, shadow=True)
    txt(sl, 0.44, 4.82, 2.20, 0.72,
        [('content-mf', {'sz': 7.6, 'b': True, 'c': PURD}),
         ('内容学术模块', {'sz': 6.6, 'b': True, 'c': INK}),
         ('含 core 业务 / 状态 / 契约层', {'sz': 5.6, 'c': MUT})], align='c', anchor='m', leading=1.15)
    ends = [('Portal Shell', '运营后台 · 全功能'), ('企微 / 公众号 H5', '销售侧 · 嵌入侧边栏'),
            ('小程序 · Taro', 'HCP 侧 · 转译复用 core'), ('App · React Native', '工程师侧 · 复用 core')]
    for i, (t, s) in enumerate(ends):
        ey = 4.78 + i * 0.34
        chip(sl, 3.42, ey, 2.98, 0.28, [(t + '　', {'sz': 6.4, 'b': True, 'c': INK}),
                                        (s, {'sz': 5.7, 'c': MUT})], None, 'FFFFFF', B_PUR, 1.1, 0.04)
        poly(sl, [(2.72, 5.18), (3.08, ey + 0.14), (3.40, ey + 0.14)], PUR, 1.3)
    txt(sl, 0.36, 6.06, 6.02, 0.14, 'Monorepo：core 共享（业务 / 状态 / 契约）· 视图按端适配 —— 新端接入只写视图层',
        sz=6.0, b=True, c=PURD, align='c')

    # ================= ⑤ 独立交付与插件扩展 =================
    rect(sl, 6.66, 4.52, 6.45, 1.74, 'FFFFFF', BORDER, 1.2, 0.07, shadow=True)
    txt(sl, 6.80, 4.60, 6.0, 0.15, '⑤ 独立交付 · 灰度发布与插件扩展', sz=8.2, b=True, c=INK)
    rel = [('① 构建', 'CI 产物'), ('② CDN', '不可变版本'), ('③ manifest', 'PR 审批'),
           ('④ 灰度', '1→10→100%'), ('⑤ 回滚', '分钟级')]
    xx = 6.80
    for i, (t, s) in enumerate(rel):
        chip(sl, xx, 4.80, 1.12, 0.44, [(t, {'sz': 6.4, 'b': True, 'c': INK}), (s, {'sz': 5.5, 'c': MUT})],
             None, L_GRN, B_GRN, 1.1, 0.05)
        if i < 4:
            seg(sl, xx + 1.12, 5.02, xx + 1.21, 5.02, GRN, 1.3)
        xx += 1.21
    txt(sl, 6.80, 5.30, 6.18, 0.13, '模块独立流水线互不阻塞 · 回滚 = manifest 指回旧版本 · 质量门禁：Playwright 冒烟 ＋ Shell E2E',
        sz=6.0, c=SUB)
    rect(sl, 6.80, 5.50, 6.18, 0.66, 'FDF2F8', 'F9A8D4', 1.1, 0.05)
    txt(sl, 6.92, 5.50, 5.94, 0.66,
        [('插件 = 远程 MF ＋ 插件 API（经 APIM · HMAC 签名）＋ manifest 声明', {'sz': 6.2, 'b': True, 'c': 'BE185D'}),
         ('租户级 feature flag 启用 · 失败自动降级隔离 · 术后随访 / 竞品情报等场景即插即用', {'sz': 6.0, 'c': SUB})],
        anchor='m', leading=1.25)

    # ---- 图例与底部 ----
    leg = [('Shell（平台团队）', SLA), ('远程模块 / 插件', PUR), ('发布 / 灰度链路', GRN), ('契约 / 安全', AZ)]
    xx = 0.40
    for t, c in leg:
        oval(sl, xx, 6.44, 0.09, 0.09, c)
        txt(sl, xx + 0.14, 6.39, len(t) * 0.10 + 0.30, 0.16, t, sz=6.4, c=MUT)
        xx += 0.32 + len(t) * 0.10 + 0.50
    txt(sl, 0.22, 6.66, 12.89, 0.16,
        '治理防腐化：shared 白名单 ＋ manifest 兼容性 CI 门禁 · 契约测试 · 季度架构体检（依赖漂移 / 跨模块引用扫描）｜ 目标：团队自治 · 互不阻塞 · 统一体验',
        sz=6.4, c=MUT, align='c')


# ------------------------------------------------------------------- build --
def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = 'One-Abbott 共享应用平台 · 技术架构图集（复用与微前端）'
    prs.core_properties.author = 'One-Abbott Platform'
    slide_reuse(prs)
    slide_mfe(prs)
    import os as _os
    out = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), OUT_NAME)
    prs.save(out)
    with open('/tmp/arch_ops.json', 'w', encoding='utf-8') as f:
        import json
        json.dump(OPS, f, ensure_ascii=False)
    print('saved', OUT_NAME, '| shapes ops:', len(OPS))


if __name__ == '__main__':
    build()
