# -*- coding: utf-8 -*-
"""地算通 V2.8.1 —— 地质测量计算工具
布局重设计 + 经纬度度分秒批量转换（导入文本 / 自动识别经纬度 / 双向转换）
新增：批量坐标正反算 / 常用单位换算 / 计算历史 / 结果导出 Excel·CSV
"""
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import tkinter.font as tkfont
import math
import re
import csv
import os
import sys
import json
import time
import random
import threading
import webbrowser
import urllib.request
from datetime import datetime


# =========================================================================
# 软件信息（版本 / 作者）
# 升级版本或变更署名时只需改这里，界面与文件属性共用同一来源
# =========================================================================
APP_NAME = '地算通'
APP_SUBTITLE = '地质测量计算工具'
APP_VER = 'V2.8.1'         # 界面显示用（三段式：主.次.修订）
APP_VER_NUM = '2.8.1.0'    # 仅与 build/version_info.txt 的 ffi 四段号对照，运行时不引用
APP_AUTHOR = '杨志秋'
APP_CONTACT = '微信 YZQ09061219'

# -------------------------------------------------------------------------
# 版本更新通知（version.json 托管在 GitHub Pages / Gitee Pages，免费、不用买服务器）
#
#   ① 新建一个公开仓库，把 version.json 丢到根目录；
#   ② 仓库 Settings → Pages 开启：分支选 main、目录选 / (root)；
#   ③ 开启后拿到形如 https://<用户名>.github.io/<仓库名>/version.json 的公开链接，
#      把它填进下面的 VERSION_URL（Gitee Pages 同理：https://<用户名>.gitee.io/<仓库名>/version.json）；
#   ④ 以后每次发新版，只改这个 version.json —— 版本号、更新日志、下载地址。
#
# version.json 格式（字段顺序随意，notes 也可写成一段字符串）：
#   {
#     "version": "V2.9.0",
#     "date": "2026-12-01",
#     "notes": ["修复……", "新增……"],
#     "url": "https://github.com/<用户名>/<仓库>/releases/download/V2.9.0/地算通V2.9.0.exe"
#   }
#
# 留空则关闭自动检查（「帮助 → 检查更新」会提示如何配置）。
# -------------------------------------------------------------------------
VERSION_URL = 'https://yzq0906.github.io/DST/version.json'
# 备用地址：主地址读不到时按顺序再试（Pages 路径对仓库名大小写敏感；
# Pages 还没开启时，公开的仓库也能先用 raw 顶上）。主地址稳定后可清空此项。
VERSION_URL_FALLBACK = [
    'https://yzq0906.github.io/dst/version.json',
    'https://raw.githubusercontent.com/yzq0906/DST/main/version.json',
]
# 发布页：即使 version.json 里的下载直链写错或失效，用户也能从这里拿到安装包
APP_RELEASES_URL = 'https://github.com/yzq0906/DST/releases/latest'
APP_UPDATE_COOLDOWN = 6 * 3600  # 自动检查的最小间隔（秒），手动检查不受限制

# -------------------------------------------------------------------------
# 更新说明（启动时按版本号比对，首次运行新版本会自动弹出；也可从「帮助」菜单查看）
# 新版本只需在列表最前面加一条，界面与弹窗都从这里取，不会与程序版本脱节
# -------------------------------------------------------------------------
CHANGELOG = [
    ('V2.8.1', '2026-11-10', [
        '关于窗口、帮助文档、文件属性与下载页精简为软件名、版本号、作者与联系方式。',
        '同步更新线上更新源与下载页，已安装旧版的用户启动后会收到本次更新提示。',
    ]),
    ('V2.8.0', '2026-11-09', [
        '功能按类型重新分组：产状与厚度 / 钻孔计算 / 坐标与方位 / 地形与坡度 / '
        '储量与工程 / 单位与工具，编号连续，左侧导航与「功能」菜单同步调整。',
        '「平面坐标正反算」新增批量导入与导出：可一次读入多组坐标，正算、反算各自表格化，'
        '支持复制与导出 Excel / CSV。',
        '「经纬度 → 直角坐标」新增批量导入与导出，每行可附高程与中央子午线，'
        '椭球与分带沿用页面选项。',
        '启动时自动显示本页更新说明；之后可在「帮助 → 更新说明」随时调阅。',
        '新增版本更新通知：启动时在后台读取托管在网页上的 version.json，'
        '发现新版本会弹出更新内容与下载入口；「帮助 → 检查更新」可随时手动检查。'
        '以后发新版只需改那一份 version.json。',
        '默认窗口按屏幕自适应加宽，批量结果表在宽屏上可完整显示；'
        '窄屏时表格出现横向滑块，不会压缩列宽把数字截断。',
        '内部改为按稳定页名注册页面，以后调整功能顺序不会再出现页面错位。',
        '更新提示里的下载地址会自动核对 Release 实际资产，并增加「发布页」入口，'
        '避免直链变动时下载不到安装包。',
    ]),
    ('V2.7.0', '2026-11-09', [
        '新增「经纬度 → 直角坐标」：一次输出地心空间直角坐标 XYZ 与高斯-克吕格平面坐标 xy，'
        '支持 CGCS2000 / WGS84 / 西安80 / 北京54 四种椭球、3 度带与 6 度带。',
        '修复含带号通用坐标显示时混入千分位逗号的问题。',
        '功能菜单与左侧导航改为按目录长度自动生成，新增功能无需再逐处改循环范围。',
    ]),
    ('V2.6.0', '2026-11-04', [
        '左侧功能菜单、选项按钮与标题卡改为磨砂玻璃质感，可在「视图」菜单关闭。',
        '新增顶部菜单栏：文件 / 功能 / 编辑 / 视图 / 帮助，并配套常用快捷键。',
        '帮助文档重写，补充快捷键、坐标格式、常见问题与隐藏玩法线索。',
        '新增若干隐藏彩蛋：地层模式配色、猫挥地质锤、罗盘校准小游戏、野外记录卡等。',
        '写入完整版本资源，右键属性可查看版本号与作者信息。',
    ]),
    ('V2.5.0', '2026-11-04', [
        '任务栏按钮改用软件图标（猫与地质锤），不再显示 Python 默认图标。',
        '主窗口、帮助、关于、消息框、文件对话框的左上角图标全部统一。',
    ]),
    ('V2.4.0', '2026-11-04', [
        '重写「经纬度度分秒批量转换」的识别逻辑：全角逗号、顿号、中文方位、'
        '度分秒符号、序号与括注都能正确识别，不再漏掉纬度。',
        '清除各页面预填的示例数值，结果位统一显示「—」；每个功能页增加「清空输入」。',
        '修复「常用单位换算」页面未注册、点击无反应与两列错位的问题。',
    ]),
    ('V2.3 及更早', '', [
        '完成 28 项地质测量常用计算、单文件界面与可执行文件打包，'
        '包含产状、厚度、钻孔、坐标、坡度、储量等基础功能。',
    ]),
]


# =========================================================================
# 经纬度解析与转换（纯函数，便于单独测试，不依赖 Tk）
# =========================================================================

def _split_hemi(tok):
    """从坐标串中提取方位标记（N/S/E/W 或 中文方位），返回 (剩余串, 方位字母 or None)"""
    t = tok.strip()
    hemi = None
    # 中文方位优先
    cn_map = [('东经', 'E'), ('西经', 'W'), ('北纬', 'N'), ('南纬', 'S'),
              ('经度', ''), ('纬度', ''), ('东', 'E'), ('西', 'W'),
              ('北', 'N'), ('南', 'S'), ('经', ''), ('纬', '')]
    for cn, code in cn_map:
        if cn in t:
            if code:
                hemi = code
            t = t.replace(cn, '')
    # 字母方位（首尾）
    m = re.search(r'([EWNSewns])\s*$', t)
    if m and hemi is None:
        hemi = m.group(1).upper()
        t = t[:m.start()]
    m2 = re.search(r'^\s*([EWNSewns])', t)
    if m2 and hemi is None:
        hemi = m2.group(1).upper()
        t = t[m2.end():]
    return t.strip(), hemi


def _parse_number(low):
    """把可能含 度分秒/冒号/空格 的串解析为十进制度；无法解析返回 None"""
    if not low:
        return None
    tmp = re.sub(r'[°′″\'\":：度分秒]', ' ', low)
    parts = re.split(r'[,\s]+', tmp.strip())
    nums = []
    for p in parts:
        p = p.strip()
        if p == '':
            continue
        try:
            nums.append(float(p))
        except ValueError:
            return None
    if not nums:
        return None
    if len(nums) >= 3:
        return nums[0] + nums[1] / 60.0 + nums[2] / 3600.0
    if len(nums) == 2:
        return nums[0] + nums[1] / 60.0
    return nums[0]


def parse_token(tok):
    """解析单个坐标 token，返回 {'val':带符号十进制, 'hemi':方位} 或 None"""
    clean, hemi = _split_hemi(tok)
    val = _parse_number(clean)
    if val is None:
        return None
    if hemi in ('S', 'W'):
        val = -abs(val)
    elif hemi in ('N', 'E'):
        val = abs(val)
    return {'val': val, 'hemi': hemi}


def _kind_of(d):
    """判断一个解析值是经度还是纬度"""
    if d['hemi'] in ('E', 'W'):
        return 'long'
    if d['hemi'] in ('N', 'S'):
        return 'lat'
    return 'long' if abs(d['val']) > 90 else 'lat'


def _assign(a, b):
    """对两个解析值分配 经度/纬度（处理两个都在 [-90,90] 的歧义）"""
    ka, kb = _kind_of(a), _kind_of(b)
    if ka == kb:
        if abs(a['val']) >= abs(b['val']):
            ka, kb = 'long', 'lat'
        else:
            ka, kb = 'lat', 'long'
    return ka, kb


def fmt_dms(v, kind):
    """十进制度 -> 带方位字母的度分秒串"""
    if v is None:
        return ''
    neg = v < 0
    a = abs(v)
    d = int(a)
    mf = (a - d) * 60.0
    m = int(mf)
    s = (mf - m) * 60.0
    letter = ('W' if neg else 'E') if kind == 'long' else ('S' if neg else 'N')
    return f"{letter}{d}°{m}′{s:.2f}″"


# ---------------------------------------------------------------------------
# 经纬度 → 直角坐标：参考椭球、地心空间直角坐标、高斯-克吕格投影
# 全部为纯函数，不依赖 Tk，便于单独验算
# ---------------------------------------------------------------------------

# 常用参考椭球：名称 -> (长半轴 a / m, 扁率倒数 1/f)
ELLIPSOIDS = {
    'CGCS2000': (6378137.0, 298.257222101),
    'WGS84': (6378137.0, 298.257223563),
    '西安80': (6378140.0, 298.257),
    '北京54': (6378245.0, 298.3),
}


def _ellip(name):
    """取椭球参数，返回 (a, f, e², e′²)"""
    a, inv_f = ELLIPSOIDS.get(name, ELLIPSOIDS['CGCS2000'])
    f = 1.0 / inv_f
    e2 = 2 * f - f * f
    return a, f, e2, e2 / (1 - e2)


def blh_to_xyz(B, L, H, ellip='CGCS2000'):
    """大地坐标 (纬度 B, 经度 L, 大地高 H) → 地心空间直角坐标 (X, Y, Z)。
    B、L 为十进制度，H 为米。"""
    a, _f, e2, _ep2 = _ellip(ellip)
    b = math.radians(B)
    l = math.radians(L)
    N = a / math.sqrt(1 - e2 * math.sin(b) ** 2)
    X = (N + H) * math.cos(b) * math.cos(l)
    Y = (N + H) * math.cos(b) * math.sin(l)
    Z = (N * (1 - e2) + H) * math.sin(b)
    return X, Y, Z


def meridian_arc(B, ellip='CGCS2000'):
    """赤道到纬度 B 的子午线弧长 (m)，B 为十进制度。"""
    a, _f, e2, _ep2 = _ellip(ellip)
    b = math.radians(B)
    e4, e6, e8 = e2 ** 2, e2 ** 3, e2 ** 4
    A0 = 1 + 3 / 4 * e2 + 45 / 64 * e4 + 175 / 256 * e6 + 11025 / 16384 * e8
    A2 = 3 / 4 * e2 + 15 / 16 * e4 + 525 / 512 * e6 + 2205 / 2048 * e8
    A4 = 15 / 64 * e4 + 105 / 256 * e6 + 2205 / 4096 * e8
    A6 = 35 / 512 * e6 + 315 / 2048 * e8
    A8 = 315 / 16384 * e8
    return a * (1 - e2) * (A0 * b - A2 / 2 * math.sin(2 * b)
                           + A4 / 4 * math.sin(4 * b) - A6 / 6 * math.sin(6 * b)
                           + A8 / 8 * math.sin(8 * b))


def gauss_forward(B, L, L0, ellip='CGCS2000'):
    """高斯-克吕格投影正算：经纬度 → 平面直角坐标 (x 北向, y 东向)。
    y 已加 500 km 假东移，不含带号；L0 为中央子午线。"""
    a, _f, e2, ep2 = _ellip(ellip)
    b = math.radians(B)
    l = math.radians(L - L0)
    sb, cb = math.sin(b), math.cos(b)
    t = math.tan(b)
    eta2 = ep2 * cb ** 2
    N = a / math.sqrt(1 - e2 * sb ** 2)
    x = (meridian_arc(B, ellip)
         + N / 2 * sb * cb * l ** 2
         + N / 24 * sb * cb ** 3 * (5 - t ** 2 + 9 * eta2 + 4 * eta2 ** 2) * l ** 4
         + N / 720 * sb * cb ** 5 * (61 - 58 * t ** 2 + t ** 4) * l ** 6)
    y = (N * cb * l
         + N / 6 * cb ** 3 * (1 - t ** 2 + eta2) * l ** 3
         + N / 120 * cb ** 5 * (5 - 18 * t ** 2 + t ** 4 + 14 * eta2
                                - 58 * eta2 * t ** 2) * l ** 5)
    return x, y + 500000.0


def zone_of(L, width=3):
    """按经度推算投影带号与中央子午线。width=3（三度带）或 6（六度带）。"""
    if width == 6:
        n = int(math.floor((L + 180) / 6)) + 1
        n = min(n, 60)
        return n, n * 6 - 183
    n = int(math.floor(L / 3 + 0.5))
    n = max(1, min(n, 120))
    return n, n * 3


# ---------------------------------------------------------------------------
# 行级坐标解析：先归一化全角标点 / 中文方位 / 括注，再提取「数值簇」
# 约定：同一行「先经后纬」；若带明确方位或数值超出纬度范围则自动纠正顺序
# ---------------------------------------------------------------------------

_FULL2HALF = str.maketrans({
    '，': ',', '、': ',', '；': ';', '：': ':', '（': '(', '）': ')',
    '／': '/', '｜': '|', '　': ' ', '－': '-', '−': '-', '–': '-',
    '。': ' ', '＇': ':', '＂': '',   # 秒符号直接去掉，避免尾随冒号粘连两个坐标
})

_HEMI_WORDS = (('东经', 'E'), ('西经', 'W'), ('北纬', 'N'), ('南纬', 'S'))

# 数值簇：1 个十进制数，或由冒号连接的 度/分/秒；可带前缀或后缀方位字母
_CLUSTER_RE = re.compile(
    r'(?:(?P<h>[EWNSewns])\s*)?'
    r'(?P<nums>[-+]?\d+(?:\.\d+)?(?:\s*:\s*\d+(?:\.\d+)?)*)'
    r'(?:\s*:?\s*(?P<h2>[EWNSewns])(?![A-Za-z0-9]))?'
)
def _normalize_coord_line(line):
    """全角标点转半角、去括注、中文方位转字母、度分秒符号统一为冒号。
    只有「夹在两个数字之间」的冒号才是度分秒分隔符，其余（标签冒号、
    秒符号留下的尾冒号）一律当作分隔符清掉，避免把两个坐标粘成一个数值。"""
    s = line.translate(_FULL2HALF)
    s = re.sub(r'\([^)]*\)', ' ', s)          # 括号注释不参与取值
    for cn, code in _HEMI_WORDS:
        s = s.replace(cn, code)                # 东经/西经/北纬/南纬 -> E/W/N/S
    for cn in ('经度', '纬度', '坐标', '东', '西', '北', '南', '经', '纬'):
        s = s.replace(cn, ' ')
    for ch in ('°', '′', '″', '度', '分', '秒', "'", '"'):
        s = s.replace(ch, ':')
    s = re.sub(r':+', ':', s)
    # 前面不是数字（如「经度:」「 」）的冒号 -> 分隔符
    s = re.sub(r'(?<!\d):+', ' ', s)
    # 后面不是数字（如 27″ 留下的尾冒号）-> 分隔符
    s = re.sub(r':+(?!\d)', ' ', s)
    return s


def _cluster_to_deg(nums_str):
    """把 116 / 116:23 / 116:23:45 折算为十进制度"""
    parts = [p for p in re.split(r'\s*:\s*', nums_str.strip().strip(':')) if p]
    if not parts:
        return None
    try:
        vals = [float(p) for p in parts]
    except ValueError:
        return None
    neg = vals[0] < 0
    a = abs(vals[0])
    if len(vals) >= 3:
        a += abs(vals[1]) / 60.0 + abs(vals[2]) / 3600.0
    elif len(vals) == 2:
        a += abs(vals[1]) / 60.0
    return -a if neg else a


def _raw_display(raw):
    """归一化片段还原为易读原值：116:23:45 -> 116°23′45″"""
    m = re.fullmatch(
        r'([EWNSewns])?\s*(-?\d+(?:\.\d+)?)(?::(\d+(?:\.\d+)?))?(?::(\d+(?:\.\d+)?))?:?',
        raw.strip())
    if not m:
        return raw.strip().strip(':')
    h, d, mi, se = m.groups()
    out = (h or '') + d
    if mi is not None:
        out += '°' + mi
        out += ('′' + se + '″') if se is not None else '′'
    return out


def _parse_coord_line(line):
    """返回 [(十进制度, 方位字母 or None, 原值片段), ...]"""
    s = _normalize_coord_line(line)
    ms = list(_CLUSTER_RE.finditer(s))
    if not ms:
        return []
    # 空格书写的度分秒：116 23 45 39 54 23 -> 两组三元组
    if len(ms) == 6 and all(
            ':' not in x.group('nums') and '.' not in x.group('nums')
            and not (x.group('h') or x.group('h2')) for x in ms):
        def tri(group):
            n = [float(x.group('nums')) for x in group]
            neg = n[0] < 0
            v = abs(n[0]) + abs(n[1]) / 60.0 + abs(n[2]) / 3600.0
            return -v if neg else v
        rawa = ' '.join(x.group('nums') for x in ms[:3])
        rawb = ' '.join(x.group('nums') for x in ms[3:])
        return [(tri(ms[:3]), None, rawa), (tri(ms[3:]), None, rawb)]
    out = []
    for x in ms:
        v = _cluster_to_deg(x.group('nums'))
        if v is None:
            continue
        h = (x.group('h') or x.group('h2') or '').upper() or None
        out.append((v, h, x.group(0).strip()))
    return out


def convert_line(line):
    """解析一行文本，返回经/纬度的十进制度与度分秒；无法解析返回 None。
    同一行默认「先经后纬」，亦兼容明确方位与「先纬后经」。"""
    if not line or not line.strip() or not re.search(r'\d', line):
        return None
    items = _parse_coord_line(line)
    if not items:
        return None
    # 方位字母决定正负
    vals = []
    for v, h, raw in items:
        if h in ('W', 'S'):
            v = -abs(v)
        elif h in ('E', 'N'):
            v = abs(v)
        vals.append((v, h, raw))
    # 常见导出格式「序号, 经度, 纬度」：丢弃行首的小整数序号
    if len(vals) >= 3 and vals[0][1] is None and vals[0][0] == int(vals[0][0]) \
            and 0 <= vals[0][0] <= 999 and abs(vals[1][0]) > 90:
        vals = vals[1:]

    lng = lat = None
    lng_raw = lat_raw = ''
    if len(vals) >= 2:
        a, b = vals[0], vals[1]
        ha, hb = a[1], b[1]
        if ha in ('N', 'S') or hb in ('E', 'W'):
            lng, lat = b, a            # 明确标注为「先纬后经」
        elif abs(a[0]) <= 90 < abs(b[0]):
            lng, lat = b, a            # 第二个值超出纬度范围，只能是经度
        else:
            lng, lat = a, b            # 约定：先经后纬
        lng_raw, lat_raw = _raw_display(lng[2]), _raw_display(lat[2])
    else:
        d = vals[0]
        k = ('long' if d[1] in ('E', 'W') else
             'lat' if d[1] in ('N', 'S') else
             'long' if abs(d[0]) > 90 else 'lat')
        if k == 'long':
            lng, lng_raw = d, _raw_display(d[2])
        else:
            lat, lat_raw = d, _raw_display(d[2])

    return {
        'lng_dec': lng[0] if lng else None,
        'lat_dec': lat[0] if lat else None,
        'lng_dms': fmt_dms(lng[0], 'long') if lng else '',
        'lat_dms': fmt_dms(lat[0], 'lat') if lat else '',
        'lng_raw': lng_raw,
        'lat_raw': lat_raw,
    }


# =========================================================================
# 配色与字体（克制的工程测量风格，避免“AI 味”）
# =========================================================================
BG      = '#eef1f4'   # 应用背景（冷灰）
PANEL   = '#f7f9fb'   # 侧栏底色
CARD    = '#ffffff'   # 卡片底色
INK     = '#1b2733'   # 主文字（墨色）
MUTED   = '#64707b'   # 次要文字
ACCENT  = '#16568f'   # 主强调（测绘蓝）
ACCENT_DK = '#103f6b'
RESULT  = '#b5471a'   # 结果数值（赤陶色，仪器感）
LINE    = '#d8dde3'   # 细分割线
READOUT_BG = '#eef3f7'

# -------------------------------------------------------------------------
# 磨砂玻璃视觉层（纯 Canvas 自绘：柔和渐变 + 上沿高光 + 亮边 + 轻投影）
# 不依赖 Windows DWM，Win10 / Win11 / 高 DPI 下表现一致
# -------------------------------------------------------------------------
GLASS_PANEL   = '#e9f1f8'   # 侧栏玻璃底
GLASS_TOP     = '#ffffff'   # 玻璃上沿高光
GLASS_BOT     = '#e1edf8'   # 玻璃下沿
GLASS_HOVER   = '#f4faff'
GLASS_SEL_TOP = '#4a8bcb'
GLASS_SEL_BOT = '#16568f'
GLASS_EDGE    = '#ffffff'
GLASS_EDGE_LO = '#cddbe8'
GLASS_SHADOW  = '#c6d5e3'
# 「地层模式」彩蛋用色：暖褐岩层
STRATA_PANEL  = '#c8b08b'
STRATA_INK    = '#3b2a18'
STRATA_ACCENT = '#a45f2c'
STRATA_AGES   = ('  寒武纪 · 基础', '  侏罗纪 · 钻孔',
                 '  白垩纪 · 工程', '  第四纪 · 工具')


def _mix(c1, c2, t):
    """两个 #rrggbb 之间线性插值，用于绘制柔和渐变。"""
    t = max(0.0, min(1.0, t))
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return '#%02x%02x%02x' % tuple(int(a[k] + (b[k] - a[k]) * t) for k in range(3))


def _round_rect(c, x0, y0, x1, y1, r, fill, outline='', width=1):
    """Canvas 上绘制圆角矩形（两个矩形 + 四个扇形），可选亮色描边。"""
    if x1 <= x0 or y1 <= y0:
        return
    r = max(0, min(r, (x1 - x0) / 2.0, (y1 - y0) / 2.0))
    kw = dict(fill=fill, outline='')
    if r <= 0:
        c.create_rectangle(x0, y0, x1, y1, **kw)
        if outline:
            c.create_rectangle(x0, y0, x1, y1, outline=outline, width=width)
        return
    c.create_rectangle(x0 + r, y0, x1 - r, y1, **kw)
    c.create_rectangle(x0, y0 + r, x1, y1 - r, **kw)
    c.create_arc(x0, y0, x0 + 2 * r, y0 + 2 * r, start=90, extent=90, **kw)
    c.create_arc(x1 - 2 * r, y0, x1, y0 + 2 * r, start=0, extent=90, **kw)
    c.create_arc(x0, y1 - 2 * r, x0 + 2 * r, y1, start=180, extent=90, **kw)
    c.create_arc(x1 - 2 * r, y1 - 2 * r, x1, y1, start=270, extent=90, **kw)
    if outline:
        c.create_line(x0 + r, y0, x1 - r, y0, fill=outline, width=width)
        c.create_line(x0 + r, y1, x1 - r, y1, fill=outline, width=width)
        c.create_line(x0, y0 + r, x0, y1 - r, fill=outline, width=width)
        c.create_line(x1, y0 + r, x1, y1 - r, fill=outline, width=width)
        for ax, ay, bx, by, st in (
                (x0, y0, x0 + 2 * r, y0 + 2 * r, 90),
                (x1 - 2 * r, y0, x1, y0 + 2 * r, 0),
                (x0, y1 - 2 * r, x0 + 2 * r, y1, 180),
                (x1 - 2 * r, y1 - 2 * r, x1, y1, 270)):
            c.create_arc(ax, ay, bx, by, start=st, extent=90,
                         style='arc', outline=outline, width=width)


class GlassButton(tk.Canvas):
    """磨砂玻璃按钮：柔和竖向渐变 + 顶部高光 + 亮色描边 + 轻投影。

    kind='nav' 用于左侧功能菜单，kind='option' 用于页面内的选项按钮。
    """

    def __init__(self, master, text='', command=None, height=34,
                 kind='nav', glass=True, padx=13, **kw):
        bg = kw.pop('bg', GLASS_PANEL)
        # Canvas 默认请求宽度达 378px，会把选项组撑散；按文字实测宽度设定
        if 'width' in kw:
            req_w = kw.pop('width')
        elif kind == 'nav':
            req_w = 180
        else:
            try:
                fnt = tkfont.Font(family='微软雅黑', size=10)
                req_w = fnt.measure(text) + padx * 2 + 12
                del fnt
            except tk.TclError:
                req_w = max(96, len(text) * 14 + padx * 2)
        tk.Canvas.__init__(self, master, width=req_w, height=height, bg=bg,
                           highlightthickness=0, bd=0, **kw)
        self._text = text
        self._command = command
        self._kind = kind
        self._glass = glass
        self._padx = padx
        self._font = ('微软雅黑', 9 if kind == 'nav' else 10)
        self._hover = False
        self._active = False
        self._pressed = False
        self._strata = False
        self.bind('<Configure>', lambda e: self._paint())
        self.bind('<Enter>', lambda e: self.set_hover(True))
        self.bind('<Leave>', self._on_leave)
        self.bind('<ButtonPress-1>', self._on_press)
        self.bind('<ButtonRelease-1>', self._on_release)
        self.after_idle(self._paint)

    def cget(self, key):
        """兼容旧代码里对 text 的读取。"""
        if key in ('text', '-text'):
            return self._text
        return tk.Canvas.cget(self, key)

    def set_glass(self, on):
        self._glass = bool(on)
        self._paint()

    def set_hover(self, on):
        if self._hover != bool(on):
            self._hover = bool(on)
            self._paint()

    def set_active(self, on):
        self._active = bool(on)
        self._paint()

    def set_strata(self, on):
        """地层模式：玻璃仍在，只是换成暖褐岩层色。"""
        self._strata = bool(on)
        self._paint()

    def invoke(self):
        if self._command:
            self._command()

    def _on_leave(self, _e):
        self._hover = False
        self._pressed = False
        self._paint()

    def _on_press(self, _e):
        self._pressed = True
        self._paint()

    def _on_release(self, _e):
        was = self._pressed
        self._pressed = False
        self._paint()
        if was:
            self.invoke()

    def _ellipsize(self, text, maxw, font):
        """按当前画布可用宽度裁字，避免高 DPI / 窄侧栏下文字被硬切。"""
        if maxw <= 12 or not text:
            return text
        tmp = self.create_text(-9999, -9999, text=text, font=font, anchor='nw')
        try:
            bb = self.bbox(tmp)
        finally:
            self.delete(tmp)
        if not bb or (bb[2] - bb[0]) <= maxw:
            return text
        s = text
        while s:
            s = s[:-1]
            tmp = self.create_text(-9999, -9999, text=s + '…', font=font, anchor='nw')
            try:
                bb = self.bbox(tmp)
            finally:
                self.delete(tmp)
            if bb and (bb[2] - bb[0]) <= maxw:
                return s + '…'
        return ''

    def _paint(self):
        try:
            w, h = self.winfo_width(), self.winfo_height()
        except tk.TclError:
            return
        if w < 6 or h < 6:
            return
        self.delete('all')
        x0, y0, x1, y1 = 2, 2, w - 3, h - 3
        r = 8 if self._kind == 'nav' else 9

        if self._strata and self._glass:
            if self._active:
                top, bot = '#e2a45d', '#a45f2c'
                fg, edge = '#fff6e6', '#f0c48a'
            elif self._hover or self._pressed:
                top, bot = '#e8d2ac', '#c8a87c'
                fg, edge = STRATA_INK, '#f4e6cd'
            else:
                top, bot = '#dcc29b', '#ac8d66'
                fg, edge = STRATA_INK, '#efe0c6'
        elif self._active:
            top, bot = GLASS_SEL_TOP, GLASS_SEL_BOT
            fg, edge = '#ffffff', '#9cc8ec'
        elif self._hover or self._pressed:
            top, bot = '#ffffff', GLASS_HOVER
            fg, edge = INK, '#ffffff'
        else:
            top, bot = GLASS_TOP, GLASS_BOT
            fg, edge = INK, GLASS_EDGE
        if self._pressed and self._active and not self._strata:
            top, bot = GLASS_SEL_BOT, ACCENT_DK

        newbg = (STRATA_PANEL if self._strata else GLASS_PANEL) if self._glass else CARD
        if tk.Canvas.cget(self, 'bg') != newbg:
            self.configure(bg=newbg)

        label = self._ellipsize(self._text, (x1 - x0) - self._padx - 10, self._font)
        if not self._glass:
            flat = ACCENT if self._active else (GLASS_HOVER if self._hover else CARD)
            _round_rect(self, x0, y0, x1, y1, 2, flat, GLASS_EDGE_LO)
            self.create_text(x0 + self._padx, (y0 + y1) / 2.0, text=label,
                             anchor='w', fill='white' if self._active else INK,
                             font=self._font)
            return

        # 轻投影 → 主体玻璃 → 竖向渐变 → 上沿高光 → 文字
        _round_rect(self, x0 + 1, y0 + 2, x1 + 1, y1 + 2, r, GLASS_SHADOW)
        _round_rect(self, x0, y0, x1, y1, r, bot, edge)
        n = max(4, y1 - y0 - 2)
        for i in range(n):
            yy = y0 + 1 + i
            if yy >= y1 - 1:
                break
            self.create_line(x0 + r, yy, x1 - r, yy,
                             fill=_mix(top, bot, i / float(max(1, n - 1))))
        self.create_line(x0 + r + 1, y0 + 1, x1 - r - 1, y0 + 1,
                         fill='#ffffff' if not self._active else '#bcdcf5')
        self.create_text(x0 + self._padx, (y0 + y1) / 2.0, text=label,
                         anchor='w', fill=fg, font=self._font)


class GlassOptionGroup(tk.Frame):
    """分段式玻璃选项按钮组，共享一个 StringVar，可替代 ttk.Radiobutton。"""

    def __init__(self, master, variable, options, command=None, glass=True, **kw):
        tk.Frame.__init__(self, master, bg=kw.pop('bg', CARD), **kw)
        self.var = variable
        self._command = command
        self.buttons = []
        for label, value in options:
            b = GlassButton(self, text=label, height=30, kind='option',
                            glass=glass, padx=12,
                            command=lambda v=value: self._choose(v))
            b.pack(side='left', padx=(0, 8))
            self.buttons.append((b, value))
        self._sync()
        try:
            self.var.trace_add('write', lambda *a: self._sync())
        except (AttributeError, tk.TclError):
            pass

    def _choose(self, value):
        changed = self.var.get() != value
        self.var.set(value)
        self._sync()
        if changed and self._command:
            self._command()

    def _sync(self):
        cur = self.var.get()
        for b, v in self.buttons:
            b.set_active(v == cur)

    def set_glass(self, on):
        for b, _ in self.buttons:
            b.set_glass(on)

    def set_strata(self, on):
        for b, _ in self.buttons:
            b.set_strata(on)


class GlassBanner(tk.Canvas):
    """侧栏顶部的磨砂玻璃标题卡。"""

    def __init__(self, master, text, height=58, glass=True, **kw):
        bg = kw.pop('bg', GLASS_PANEL)
        tk.Canvas.__init__(self, master, height=height, bg=bg,
                           highlightthickness=0, bd=0, **kw)
        self._text = text
        self._glass = glass
        self._strata = False
        self.bind('<Configure>', lambda e: self._paint())
        self.after_idle(self._paint)

    def set_glass(self, on):
        self._glass = bool(on)
        self._paint()

    def set_strata(self, on):
        self._strata = bool(on)
        self._paint()

    def _paint(self):
        try:
            w, h = self.winfo_width(), self.winfo_height()
        except tk.TclError:
            return
        if w < 6 or h < 6:
            return
        self.delete('all')
        if not self._glass:
            self.create_rectangle(0, 0, w, h, fill=PANEL, outline='')
            self.create_text(16, h / 2.0, text=self._text, anchor='w', fill=INK,
                             font=('微软雅黑', 12, 'bold'))
            self.create_rectangle(16, h - 18, 42, h - 16, fill=ACCENT, outline='')
            return
        x0, y0, x1, y1 = 10, 8, w - 10, h - 8
        if self._strata:
            top, bot, edge, tcol, bar = ('#f0dcbb', '#b99b72', '#f6e8cf',
                                         STRATA_INK, STRATA_ACCENT)
            if tk.Canvas.cget(self, 'bg') != STRATA_PANEL:
                self.configure(bg=STRATA_PANEL)
        else:
            top, bot, edge, tcol, bar = (GLASS_TOP, GLASS_BOT, GLASS_EDGE,
                                         INK, ACCENT)
            if tk.Canvas.cget(self, 'bg') != GLASS_PANEL:
                self.configure(bg=GLASS_PANEL)
        _round_rect(self, x0 + 1, y0 + 2, x1 + 1, y1 + 2, 10,
                    GLASS_SHADOW if not self._strata else '#9a7f5c')
        _round_rect(self, x0, y0, x1, y1, 10, bot, edge)
        n = max(4, y1 - y0 - 2)
        for i in range(n):
            yy = y0 + 1 + i
            if yy >= y1 - 1:
                break
            self.create_line(x0 + 10, yy, x1 - 10, yy,
                             fill=_mix(top, bot, i / float(max(1, n - 1))))
        self.create_text(x0 + 14, (y0 + y1) / 2.0 - 3, text=self._text,
                         anchor='w', fill=tcol, font=('微软雅黑', 12, 'bold'))
        self.create_rectangle(x0 + 14, y1 - 10, x0 + 40, y1 - 8,
                              fill=bar, outline='')


# -------------------------------------------------------------------------
# 彩蛋文案（地质主题，低调克制，绝不干扰任何计算）
# -------------------------------------------------------------------------
EASTER_QUOTES = (
    '岩层不会说谎，只是说得比较慢。',
    '真倾角永远比视倾角诚实。',
    '所有等高线闭合的地方，都藏着一个故事。',
    '野外三件宝：地质锤、罗盘、一瓶还剩下的水。',
    '把标本装进背包的时候，别忘了把坐标记进本子。',
    '地层的新老关系，比任何日历都可靠。',
    '测量误差不会消失，它只是换了个地方等你。',
    '一条好剖面，胜过三千字描述。',
    '罗盘指北，经验指路，午饭指时间。',
)

EASTER_ROCKS = (
    ('花岗岩', '深成侵入岩，石英 + 长石 + 云母。抗风化，也抗工期一改再改。'),
    ('玄武岩', '喷出岩，常见气孔与杏仁构造。熔岩冷却得比验收还快。'),
    ('石灰岩', '碳酸盐岩，滴盐酸会起泡。喀斯特的幕后主角。'),
    ('砂岩', '碎屑岩，层理最会讲故事。交错层理里写着古水流方向。'),
    ('页岩', '泥质岩，页理发育。遇水易软化，边坡上的重点关照对象。'),
    ('大理岩', '石灰岩变质而来，方解石重结晶。既是石材，也是标本柜的门面。'),
    ('片麻岩', '高级变质岩，条带状构造。褶皱起来比任何地图都复杂。'),
    ('砾岩', '粗碎屑岩，砾石磨圆度记录搬运距离。野外一眼就能认出。'),
)

EASTER_TITLE_LINES = (
    '其实是个带罗盘的计算器',
    '地质人的小算盘',
    '今天也要记得带地质锤',
    '把坐标记进本子，别只记在脑子里',
    '岩层不会说谎，只是说得比较慢',
)

EASTER_SPOTS = (
    (116.391, 39.907, '北京 · 中国地质博物馆一带'),
    (86.9250, 27.9881, '珠穆朗玛峰 · 8848.86 m'),
    (0.0, 0.0, '几内亚湾 · 经纬度原点'),
)


# -------------------------------------------------------------------------
# 功能目录：按「类型」分组，列表顺序即左侧导航、顶部菜单与功能编号的顺序。
# 名称是稳定键（页名），对外的「13. xxx」编号由顺序自动生成，
# 因此以后调整分组或插入新功能，只需改这张表，页面注册不会错位。
# -------------------------------------------------------------------------
FUNC_GROUPS = [
    ('产状与厚度', [
        '视倾角↔真倾角互算', '走向倾向互换', '地层真/视厚度',
        '岩层水平厚度', '露头出露宽度', '实测剖面厚度',
    ]),
    ('钻孔计算', [
        '钻孔分层真厚度', '钻孔视厚度校正', '钻孔真倾角校正',
        '钻孔垂直埋深', '钻孔水平投影', '多段钻孔累计', '地表岩层埋深',
    ]),
    ('坐标与方位', [
        '经纬度度分秒批量转换', '经纬度→直角坐标', '平面坐标正反算',
        '批量坐标正反算', '方位角前后视', '坐标格网长度', '导线三维XYZ增量',
    ]),
    ('地形与坡度', [
        '坡度高差平距', '平距斜距换算', '等高线坡角', '坡度坡率互转',
    ]),
    ('储量与工程', [
        '块段矿石储量', '岩土稳定坡角',
    ]),
    ('单位与工具', [
        '面积单位换算', '常用单位换算', '计算历史',
    ]),
]


# -------------------------------------------------------------------------
# 批量文本切分：逗号（含全角）、分号、顿号、制表符、空格都视为列分隔
# -------------------------------------------------------------------------
_BATCH_SEP_RE = re.compile(r'[,;，；、\t 　]+')


def split_batch_fields(line):
    return [t for t in _BATCH_SEP_RE.split(line.strip()) if t]


# 平面坐标批量换算的两种表头（正算 / 反算列不同，切换时整表重建）
_COORD_BATCH_FWD = (
    ['序号', '起点X', '起点Y', '方位角(°)', '距离(m)', '终点X', '终点Y'],
    [44, 92, 92, 88, 84, 92, 92],
)
_COORD_BATCH_REV = (
    ['序号', 'X1', 'Y1', 'X2', 'Y2', '方位角(°)', '距离(m)'],
    [44, 92, 92, 92, 92, 88, 84],
)


class BatchPanel(tk.Frame):
    """通用「批量导入 → 逐行换算 → 表格 → 复制/导出」面板。

    交互顺序与「经纬度度分秒批量转换」保持一致，解析规则由调用方以 parse 回调注入，
    因此同一套界面可以挂在不同换算页上，不必把 UI 代码复制成好几份。
    """

    def __init__(self, master, app, headers, widths, parse, default_name,
                 hint='', text_height=5, tree_height=6):
        super().__init__(master, bg=CARD)
        self.app = app
        self.default_name = default_name
        self.parse = parse
        self.headers = list(headers)
        self.rows = []

        box = tk.LabelFrame(self, text=' 批量换算（每行一组，可导入 .txt / .csv） ',
                            bg=CARD, fg=INK, font=('微软雅黑', 9),
                            relief='groove', bd=1, highlightbackground=LINE)
        box.pack(fill='x', pady=(2, 3))
        self.text = scrolledtext.ScrolledText(box, height=text_height,
                                              font=('Consolas', 9), bg='#fbfcfd',
                                              relief='flat', bd=0)
        self.text.pack(fill='both', expand=True, padx=6, pady=5)

        if hint:
            tk.Label(self, text=hint, bg=CARD, fg=MUTED, font=('微软雅黑', 8),
                     wraplength=620, justify='left',
                     anchor='w').pack(fill='x', pady=(2, 0))

        btns = tk.Frame(self, bg=CARD)
        btns.pack(fill='x', pady=6)
        for txt, style, fn in (
                ('导入文本', 'Ghost.TButton', self.load_file),
                ('开始换算', 'Primary.TButton', self.run),
                ('清空', 'Ghost.TButton', self.clear),
                ('复制结果', 'Ghost.TButton', self.copy),
                ('导出结果', 'Ghost.TButton', self.export)):
            ttk.Button(btns, text=txt, style=style, command=fn).pack(side='left', padx=3)

        wrap = tk.Frame(self, bg=CARD)
        wrap.pack(fill='both', expand=True)
        self.tree = ttk.Treeview(wrap, columns=[], show='headings',
                                 style='Geo.Batch.Treeview', height=tree_height)
        vsb = ttk.Scrollbar(wrap, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        self.hsb = ttk.Scrollbar(self, orient='horizontal', command=self.tree.xview)
        self.tree.configure(xscrollcommand=self.hsb.set)
        self.hsb.pack(fill='x')
        self.tree.bind('<Configure>', self._sync_hsb, add='+')

        self.status = tk.StringVar(value='等待输入…')
        self._status_label = tk.Label(self, textvariable=self.status, bg=CARD,
                                      fg=MUTED, font=('微软雅黑', 9), anchor='w')
        self._status_label.pack(fill='x')

        self.set_columns(self.headers, widths, parse)

    def _sync_hsb(self, _e=None):
        """列宽合计不超过可视宽度时收起横向滑块，避免出现一条永远用不上的灰条。"""
        try:
            total = getattr(self, '_need_w', 0)
            view = self.tree.winfo_width()
        except tk.TclError:
            return
        need = total > view + 1
        if need and not self.hsb.winfo_ismapped():
            self.hsb.pack(fill='x', before=self._status_label)
        elif not need and self.hsb.winfo_ismapped():
            self.hsb.pack_forget()

    # ---- 列定义 ----
    def set_columns(self, headers, widths, parse):
        """切换换算口径（例如正算 / 反算）时重建表头，并清掉旧结果。"""
        self.parse = parse
        self.headers = list(headers)
        cols = ['c%d' % i for i in range(len(headers))]
        self.tree.configure(columns=cols)
        self._need_w = sum(widths)
        for col, head, w in zip(cols, headers, widths):
            self.tree.heading(col, text=head)
            # 列宽固定：宁可留白或交给横向滑块，也不压缩列宽导致数字被截
            self.tree.column(col, width=w, anchor='center', stretch=False)
        self._reset_rows()

    def _reset_rows(self):
        try:
            self.tree.delete(*self.tree.get_children())
        except tk.TclError:
            pass
        self.rows = []

    # ---- 动作 ----
    def load_file(self):
        path = filedialog.askopenfilename(
            parent=self.app.root, title='导入坐标文本',
            filetypes=[('文本文件', '*.txt'), ('CSV 文件', '*.csv'), ('所有文件', '*.*')])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as fh:
                content = fh.read()
        except UnicodeDecodeError:
            with open(path, 'r', encoding='gbk') as fh:
                content = fh.read()
        self.text.delete('1.0', tk.END)
        self.text.insert(tk.END, content)
        self.status.set('已导入：' + os.path.basename(path))
        self.run()

    def run(self):
        self._reset_rows()
        ok = skip = 0
        for line in self.text.get('1.0', tk.END).splitlines():
            if not line.strip():
                continue
            try:
                row = self.parse(line)
            except (ValueError, TypeError, ZeroDivisionError):
                row = None
            if not row:
                skip += 1
                continue
            ok += 1
            self.rows.append([str(len(self.rows) + 1)] + [str(v) for v in row])
            self.tree.insert('', tk.END, values=self.rows[-1])
        self.status.set('已换算 %d 条，跳过无法识别 %d 行。' % (ok, skip))
        return ok

    def clear(self):
        self.text.delete('1.0', tk.END)
        self._reset_rows()
        self.status.set('已清空。')

    def copy(self):
        if not self.rows:
            self.app._dialog('showinfo', '提示', '暂无可复制的结果，请先换算。')
            return
        lines = ['\t'.join(self.headers)]
        lines.extend('\t'.join(r) for r in self.rows)
        self.app.root.clipboard_clear()
        self.app.root.clipboard_append('\n'.join(lines))
        self.status.set('结果已复制到剪贴板。')

    def export(self):
        if not self.rows:
            self.app._dialog('showinfo', '提示', '暂无可导出的结果，请先换算。')
            return
        res = self.app.export_table(self.headers, self.rows, self.default_name)
        if res:
            self.status.set('已导出：' + os.path.basename(res[1]))


# =========================================================================
# 主程序
# =========================================================================
class GeoCalcApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f'{APP_NAME} {APP_VER}')
        # 默认尺寸随屏幕自适应：宽屏上给批量结果表留出足够列宽，小屏则退回 980
        _sw, _sh = root.winfo_screenwidth(), root.winfo_screenheight()
        _w = max(980, min(1200, _sw - 80))
        _h = max(680, min(760, _sh - 120))
        self.root.geometry('%dx%d' % (_w, _h))
        self.root.resizable(True, True)
        self.root.minsize(900, 664)
        self.root.configure(bg=BG)
        # 磨砂玻璃效果状态（可在「视图」菜单中切换）
        self.glass_var = tk.BooleanVar(value=True)
        self._glass_widgets = []     # GlassButton / GlassOptionGroup / GlassBanner
        self._glass_frames = []      # 需要随开关换底色的 Frame / Label
        self._glass_options = []     # GlassOptionGroup 单独留一份便于刷新
        self._group_labels = []      # 侧栏分组标题（地层模式会改文字）
        self._readout_labels = []    # 结果读数 Label（地层模式会换字体）
        # 软件图标（打包后从 MEIPASS 读取，未打包时读取源码同目录）
        self._set_icon()
        # 窗口真正映射后再补一次，确保任务栏按钮拿到图标
        self.root.after(150, self._set_taskbar_icon)
        self.root.after(700, self._set_taskbar_icon)

        self.style_app()
        self.build_shell()

        # 功能目录由模块级 FUNC_GROUPS 派生：
        #   func_bases —— 稳定页名，页面注册与查表都用它
        #   func_names —— 带编号的显示名，编号随顺序自动生成
        self.func_groups = []
        self.func_bases = []
        for _gname, _bases in FUNC_GROUPS:
            start = len(self.func_bases)
            self.func_bases.extend(_bases)
            self.func_groups.append((_gname, list(range(start, len(self.func_bases)))))
        self.func_names = ['%d. %s' % (i + 1, b) for i, b in enumerate(self.func_bases)]
        self.func_index = {b: i for i, b in enumerate(self.func_bases)}
        self.tab_frames = {}
        self.nav_buttons = []
        self.nav_inds = []      # 左侧导航激活指示条
        self.active_idx = -1
        self.build_nav()
        # 历史记录初始化（须在包装计算方法之前）
        self.history = []
        self._last_log = {}
        self._wrap_calc_methods()
        self.build_all_tabs()
        self.switch_to(0)
        # 菜单栏与快捷键放到最后，确保 29 个功能页都已就绪
        self.build_menubar()
        self._bind_shortcuts()
        self._init_easter()
        # 主窗口先显示完毕，再弹更新说明，避免启动瞬间两层窗口一起闪
        self.root.after(700, self._maybe_show_changelog)
        # 更新说明之后再做版本检查（后台线程），有新版才弹提示
        self.root.after(2600, lambda: self.check_update(False))

    # ---------------- 头部徽标（手绘测量罗盘，纯 Canvas 绘制，非通用色块） ----------------
    def _make_logo(self, parent):
        """绘制一枚“测量罗盘 / 十字瞄准”徽标，呼应地质测量主题，扁平克制、无渐变光晕。"""
        size = 30
        c = tk.Canvas(parent, width=size, height=size, bg='#ffffff',
                      highlightthickness=0, cursor='hand2')
        c.pack(side='left', padx=(18, 0), pady=15)
        self.logo_canvas = c
        cx = cy = size / 2.0
        r = size / 2.0 - 3
        self.logo_geom = (cx, cy, r)
        c.bind('<Button-1>', self._logo_click)
        # 外圈（测绘蓝，仪器感）
        c.create_oval(cx - r, cy - r, cx + r, cy + r,
                      outline=ACCENT, width=2)
        # 十字瞄准线（细灰）
        c.create_line(cx, cy - r, cx, cy + r, fill='#9fb3c7', width=1)
        c.create_line(cx - r, cy, cx + r, cy, fill='#9fb3c7', width=1)
        # 方位刻度（N/E/S/W 四个短标记，暗示方位）
        for ang in (0, 90, 180, 270):
            a = math.radians(ang - 90)
            x1 = cx + (r - 3) * math.cos(a)
            y1 = cy + (r - 3) * math.sin(a)
            x2 = cx + r * math.cos(a)
            y2 = cy + r * math.sin(a)
            c.create_line(x1, y1, x2, y2, fill=ACCENT, width=1.5)
        # 指针（赤陶色指向东北，墨灰尾指向西南）
        self.logo_needle_a = c.create_line(
            cx, cy, cx + r * 0.7, cy - r * 0.7,
            fill=RESULT, width=2.5, capstyle='round')
        self.logo_needle_b = c.create_line(
            cx, cy, cx - r * 0.5, cy + r * 0.5,
            fill=MUTED, width=1.5, capstyle='round')
        # 中心轴点
        c.create_oval(cx - 2, cy - 2, cx + 2, cy + 2,
                      fill=ACCENT, outline=ACCENT)

    # ---------------- 软件图标 ----------------
    def _set_icon(self):
        # 地算通.ico 含 16~256 多档尺寸，小图标下更清晰；旧文件作为兜底
        names = ('地算通.ico', '1_256x256.ico')
        if getattr(sys, 'frozen', False):
            # 打包后：优先 MEIPASS 内的资源，其次与 exe 同目录
            bases = [getattr(sys, '_MEIPASS', ''),
                     os.path.dirname(sys.executable)]
        else:
            bases = [os.path.dirname(os.path.abspath(__file__))]
        self.icon_path = None
        for ico in names:
            for base in bases:
                if not base:
                    continue
                p = os.path.join(base, ico)
                if os.path.exists(p):
                    self.icon_path = p
                    self._apply_icon(self.root)
                    self._set_taskbar_icon()
                    return

    def _apply_icon(self, win):
        """把软件图标应用到主窗口及每一个弹出窗口（标题栏左上角）。"""
        p = getattr(self, 'icon_path', None)
        if not p:
            return
        try:
            win.iconbitmap(p)          # Windows 标题栏图标
        except (tk.TclError, AttributeError):
            pass

    def _set_taskbar_icon(self):
        """让任务栏按钮也显示软件图标。

        Tk 的 iconbitmap 只保证标题栏正确；任务栏图标取自窗口的 ICON_BIG，
        且进程若没有独立 AppUserModelID，Windows 会把它并进 python.exe /
        引导进程的分组，沿用对方的图标。这里两件事都补上。
        """
        if sys.platform != 'win32':
            return
        p = getattr(self, 'icon_path', None)
        if not p or not os.path.exists(p):
            return
        try:
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.WinDLL('user32', use_last_error=True)
            shell32 = ctypes.WinDLL('shell32', use_last_error=True)
        except Exception:
            return

        # 1) 独立的 AppUserModelID：任务栏不再沿用宿主解释器的图标
        try:
            shell32.SetCurrentProcessExplicitAppUserModelID(
                ctypes.c_wchar_p('YangZhiqiu.DiSuanTong.GeoCalc'))
        except Exception:
            pass

        # 2) 用 ICON_BIG / ICON_SMALL 两个尺寸覆盖窗口图标
        try:
            user32.GetAncestor.restype = wintypes.HWND
            user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
            user32.LoadImageW.restype = wintypes.HANDLE
            user32.LoadImageW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR,
                                          wintypes.UINT, ctypes.c_int,
                                          ctypes.c_int, wintypes.UINT]
            user32.SendMessageW.restype = ctypes.c_ssize_t
            user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT,
                                            wintypes.WPARAM, wintypes.LPARAM]
            user32.GetSystemMetrics.restype = ctypes.c_int

            self.root.update_idletasks()
            hwnd = user32.GetAncestor(self.root.winfo_id(), 2)  # GA_ROOT
            if not hwnd:
                return
            IMAGE_ICON, LR_LOADFROMFILE = 1, 0x00000010
            WM_SETICON, ICON_SMALL, ICON_BIG = 0x0080, 0, 1
            big = user32.GetSystemMetrics(11) or 32   # SM_CXICON
            for size, which in ((big, ICON_BIG), (16, ICON_SMALL)):
                hicon = user32.LoadImageW(None, p, IMAGE_ICON,
                                          size, size, LR_LOADFROMFILE)
                if hicon:
                    user32.SendMessageW(hwnd, WM_SETICON, which, hicon)
        except Exception:
            pass

    def _dialog(self, kind, title, message):
        """统一的消息弹窗：指定 parent，让系统弹窗继承主窗口的软件图标。"""
        fn = getattr(messagebox, kind)
        try:
            fn(title, message, parent=self.root)
        except tk.TclError:
            fn(title, message)

    # ---------------- 主题 ----------------
    def style_app(self):
        s = ttk.Style()
        s.theme_use('clam')
        s.configure('Geo.TEntry', fieldbackground='#ffffff', background='#ffffff',
                    bordercolor=LINE, lightcolor=LINE, darkcolor=LINE,
                    foreground=INK, font=('微软雅黑', 10))
        s.map('Geo.TEntry', focus=[('!disabled', {'bordercolor': ACCENT})])
        s.configure('Geo.Treeview', background='#ffffff', fieldbackground='#ffffff',
                    foreground=INK, font=('微软雅黑', 10), rowheight=26, bordercolor=LINE)
        s.configure('Geo.Treeview.Heading', background='#f0f3f6', foreground=INK,
                    font=('微软雅黑', 10, 'bold'))
        # 批量结果表列多、数字长，字号收小一档才能在一张卡片里放全
        s.configure('Geo.Batch.Treeview', background='#ffffff',
                    fieldbackground='#ffffff', foreground=INK,
                    font=('微软雅黑', 9), rowheight=24, bordercolor=LINE)
        s.configure('Geo.Batch.Treeview.Heading', background='#f0f3f6',
                    foreground=INK, font=('微软雅黑', 9, 'bold'))
        s.configure('Geo.TRadiobutton', background=CARD, foreground=INK,
                    font=('微软雅黑', 10))
        # 按钮也走玻璃语言：浅色半透感底 + 亮色描边 + hover 提亮
        s.configure('Primary.TButton', background=ACCENT, foreground='white',
                    font=('微软雅黑', 10), borderwidth=1, relief='solid',
                    bordercolor=ACCENT_DK, lightcolor='#7fb4e4',
                    darkcolor=ACCENT_DK, padding=(12, 6))
        s.map('Primary.TButton',
              background=[('active', ACCENT_DK), ('pressed', '#0d3559')],
              relief=[('pressed', 'sunken')])
        s.configure('Ghost.TButton', background='#f2f8fd', foreground=INK,
                    font=('微软雅黑', 10), borderwidth=1, relief='solid',
                    bordercolor=GLASS_EDGE_LO, lightcolor='#ffffff',
                    darkcolor=GLASS_SHADOW, padding=(12, 6))
        s.map('Ghost.TButton',
              background=[('active', '#ffffff'), ('pressed', '#e2edf7')],
              relief=[('pressed', 'sunken')])
        # 细窄滚动条（克制、不抢眼）
        s.configure('Geo.Vertical.TScrollbar', troughcolor=GLASS_PANEL,
                    background='#b9c4cf', borderwidth=0, relief='flat',
                    activebackground=ACCENT)
        s.map('Geo.Vertical.TScrollbar',
              background=[('active', ACCENT), ('pressed', ACCENT_DK)])

    # ---------------- 外壳（头部 + 侧栏 + 内容区） ----------------
    def build_shell(self):
        # 头部
        header = tk.Frame(self.root, bg='#ffffff', height=60)
        header.pack(fill='x')
        header.pack_propagate(False)
        # 左侧：手绘测量罗盘徽标 + 标题 + 副标题
        self._make_logo(header)
        self.title_label = tk.Label(header, text='地算通', bg='#ffffff', fg=INK,
                                    font=('微软雅黑', 18, 'bold'), cursor='hand2')
        self.title_label.pack(side='left', padx=(10, 8), pady=8)
        self.title_label.bind('<Button-1>', self._title_click)
        self.subtitle_label = tk.Label(header, text='地质测量计算工具',
                                       bg='#ffffff', fg=MUTED,
                                       font=('微软雅黑', 10))
        self.subtitle_label.pack(side='left', pady=10)
        tk.Frame(header, bg=LINE, width=1, height=26).pack(
            side='left', padx=12, pady=17)
        # 右侧版本与署名
        right = tk.Frame(header, bg='#ffffff')
        right.pack(side='right', padx=18, pady=8)
        tk.Label(right, text=APP_VER, bg=ACCENT, fg='white',
                 font=('微软雅黑', 10, 'bold'), padx=8, pady=3).pack(side='right')
        tk.Label(right, text=APP_AUTHOR, bg='#ffffff', fg=MUTED,
                 font=('微软雅黑', 9)).pack(side='right', padx=(0, 12))
        # 强调色下沿
        tk.Frame(self.root, bg=ACCENT, height=3).pack(fill='x')

        body = tk.Frame(self.root, bg=BG)
        body.pack(fill='both', expand=True)

        # 侧栏（磨砂玻璃面板）
        self.nav_panel = tk.Frame(body, bg=GLASS_PANEL, width=216)
        self.nav_panel.pack(side='left', fill='y')
        self.nav_panel.pack_propagate(False)
        self._glass_frames.append(self.nav_panel)
        self.nav_banner = GlassBanner(self.nav_panel, '功能目录',
                                      height=58, glass=True, bg=GLASS_PANEL)
        self.nav_banner.pack(fill='x', padx=8, pady=(12, 4))
        self._glass_widgets.append(self.nav_banner)

        # 侧栏底部按钮（先放，避免被滚动区挤占）
        foot = tk.Frame(self.nav_panel, bg=GLASS_PANEL)
        foot.pack(side='bottom', fill='x', pady=10)
        self._glass_frames.append(foot)
        for label, cmd in (('帮助文档', self.show_doc), ('关于软件', self.show_about)):
            b = GlassButton(foot, text=label, command=cmd, height=32,
                            kind='option', glass=True, padx=14, bg=GLASS_PANEL)
            b.pack(fill='x', padx=12, pady=4)
            self._glass_widgets.append(b)

        # 侧栏可滚动导航列表
        nav_scroll = ttk.Scrollbar(self.nav_panel, orient='vertical',
                                   style='Geo.Vertical.TScrollbar')
        nav_canvas = tk.Canvas(self.nav_panel, bg=GLASS_PANEL, highlightthickness=0,
                               yscrollcommand=nav_scroll.set)
        nav_scroll.config(command=nav_canvas.yview)
        self.nav_canvas = nav_canvas
        self.nav_list_frame = tk.Frame(nav_canvas, bg=GLASS_PANEL)
        self._glass_frames.extend((nav_canvas, self.nav_list_frame))
        self.nav_window = nav_canvas.create_window((0, 0), window=self.nav_list_frame, anchor='nw')
        self.nav_list_frame.bind(
            '<Configure>',
            lambda e: nav_canvas.config(scrollregion=nav_canvas.bbox('all')))
        nav_canvas.bind('<Configure>',
                        lambda e: nav_canvas.itemconfig(self.nav_window, width=e.width))
        nav_scroll.pack(side='right', fill='y', pady=4)
        nav_canvas.pack(side='left', fill='both', expand=True, padx=(4, 0), pady=4)
        # 鼠标滚轮：同时绑定到画布与内部列表框，确保悬停在按钮上也能滚动
        nav_canvas.bind('<MouseWheel>',
                        lambda e: self._do_wheel(nav_canvas, e))
        self.nav_list_frame.bind('<MouseWheel>',
                                 lambda e: self._do_wheel(nav_canvas, e))

        # 内容滚动区（右侧固定面板）
        self.content = tk.Frame(body, bg=BG)
        self.content.pack(side='right', fill='both', expand=True)
        self.canvas = tk.Canvas(self.content, highlightthickness=0, bg=BG)
        self.scroll_bar = ttk.Scrollbar(self.content, orient='vertical',
                                        command=self.canvas.yview,
                                        style='Geo.Vertical.TScrollbar')
        self.scroll_box = tk.Frame(self.canvas, bg=BG)
        self.content_window = self.canvas.create_window((0, 0), window=self.scroll_box, anchor='nw')
        self.scroll_box.bind('<Configure>',
                             lambda e: self.canvas.config(scrollregion=self.canvas.bbox('all')))
        self.canvas.bind('<Configure>',
                         lambda e: self.canvas.itemconfig(self.content_window, width=e.width))
        self.canvas.config(yscrollcommand=self.scroll_bar.set)
        self.canvas.pack(side='left', fill='both', expand=True)
        self.scroll_bar.pack(side='right', fill='y')
        self.scroll_box.columnconfigure(0, weight=1)
        # 鼠标滚轮（悬停在卡片上也能滚动内容区）
        self.canvas.bind('<MouseWheel>',
                         lambda e: self._do_wheel(self.canvas, e))
        self.scroll_box.bind('<MouseWheel>',
                             lambda e: self._do_wheel(self.canvas, e))

    def _do_wheel(self, canvas, event):
        # delta 为负表示向下滚；每个刻度约移动 8px，手感细腻不跳页
        canvas.yview_scroll(int(-1 * event.delta / 15), 'units')

    # ---------------- 侧栏导航按钮（分组 + 激活指示条） ----------------
    def build_nav(self):
        for gname, idxs in self.func_groups:
            lab = tk.Label(self.nav_list_frame, text='  ' + gname,
                           bg=GLASS_PANEL, fg=ACCENT,
                           font=('微软雅黑', 9, 'bold'), anchor='w')
            lab.pack(anchor='w', padx=14, pady=(10, 4))
            self._glass_frames.append(lab)
            self._group_labels.append((lab, lab.cget('text')))
            for idx in idxs:
                name = self.func_names[idx]
                row = tk.Frame(self.nav_list_frame, bg=GLASS_PANEL)
                row.pack(fill='x', padx=6, pady=1)
                self._glass_frames.append(row)
                ind = tk.Frame(row, bg=GLASS_PANEL, width=3)
                ind.pack(side='left', fill='y')
                self._glass_frames.append(ind)
                b = GlassButton(row, text=name, height=32, kind='nav',
                                glass=True, padx=12, bg=GLASS_PANEL,
                                command=lambda i=idx: self.switch_to(i))
                b.pack(side='left', fill='both', expand=True)
                b.bind('<Enter>', lambda ev, btn=b: self._nav_hover(btn, True), add='+')
                b.bind('<Leave>', lambda ev, btn=b: self._nav_hover(btn, False), add='+')
                self.nav_buttons.append(b)
                self.nav_inds.append(ind)
                self._glass_widgets.append(b)
        # Tk 的 <MouseWheel> 只派发给光标下的最深控件，父框架不在其 bindtags 中；
        # 因此必须把滚轮处理逐个绑定到左侧菜单里的每一个控件（分组标题/条目/按钮），
        # 这样鼠标悬停在任意功能上都能滚动，而不是只有停在空白处才滚得动。
        handler = lambda e: self._do_wheel(self.nav_canvas, e)
        self._bind_wheel(self.nav_list_frame, handler)
        self.nav_canvas.bind('<MouseWheel>', handler)

    def _bind_wheel(self, widget, handler):
        try:
            widget.bind('<MouseWheel>', handler)
        except tk.TclError:
            pass
        for ch in widget.winfo_children():
            self._bind_wheel(ch, handler)

    def _nav_hover(self, btn, entering):
        idx = self.nav_buttons.index(btn)
        if idx == self.active_idx:
            return
        btn.set_hover(entering)

    def set_active(self, idx):
        self.active_idx = idx
        if getattr(self, 'strata_on', False):
            neutral, hot = STRATA_PANEL, STRATA_ACCENT
        else:
            neutral = GLASS_PANEL if self.glass_var.get() else PANEL
            hot = ACCENT
        for i, b in enumerate(self.nav_buttons):
            b.set_active(i == idx)
            self.nav_inds[i].config(bg=hot if i == idx else neutral)

    def _page(self, card, base):
        """按稳定页名注册功能页：编号随目录顺序变化，页名不变，重排不会错位。"""
        self.tab_frames[base] = card

    def switch_to(self, idx):
        for f in self.tab_frames.values():
            f.grid_forget()
        try:
            target = self.tab_frames[self.func_bases[idx]]
            target.grid(row=0, column=0, sticky='nsew', padx=12, pady=8)
            self.canvas.update_idletasks()
            self.canvas.config(scrollregion=self.canvas.bbox('all'))
            self.set_active(idx)
            self.canvas.yview_moveto(0)
        except (IndexError, KeyError):
            pass

    # ---------------- 顶部菜单栏 ----------------
    def _menu(self, parent):
        return tk.Menu(parent, tearoff=0, bg=CARD, fg=INK,
                       activebackground='#dfeefb', activeforeground=ACCENT,
                       relief='flat', bd=0, font=('微软雅黑', 9))

    def build_menubar(self):
        mb = self._menu(self.root)
        self.root.config(menu=mb)
        self.menubar = mb

        fm = self._menu(mb)
        mb.add_cascade(label='文件', menu=fm)
        fm.add_command(label='导入坐标文本…', accelerator='Ctrl+O',
                       command=self.menu_import)
        fm.add_command(label='导出当前结果…', accelerator='Ctrl+S',
                       command=self.menu_export)
        fm.add_separator()
        fm.add_command(label='退出', accelerator='Ctrl+Q',
                       command=self.menu_quit)

        jm = self._menu(mb)
        mb.add_cascade(label='功能', menu=jm)
        for gname, idxs in self.func_groups:
            sub = self._menu(jm)
            jm.add_cascade(label=gname, menu=sub)
            for i in idxs:
                sub.add_command(label=self.func_names[i],
                                command=lambda x=i: self.switch_to(x))

        em = self._menu(mb)
        mb.add_cascade(label='编辑', menu=em)
        em.add_command(label='清空当前输入', accelerator='Ctrl+W',
                       command=self.menu_clear)
        em.add_command(label='复制当前结果', accelerator='Ctrl+Shift+C',
                       command=self.menu_copy)
        em.add_separator()
        em.add_command(label='打开计算历史', accelerator='Ctrl+H',
                       command=self._goto('计算历史'))

        vm = self._menu(mb)
        mb.add_cascade(label='视图', menu=vm)
        vm.add_checkbutton(label='磨砂玻璃质感', variable=self.glass_var,
                           command=self.toggle_glass)
        vm.add_separator()
        vm.add_command(label='返回首页', command=lambda: self.switch_to(0))
        vm.add_command(label='经纬度度分秒转换', accelerator='Ctrl+D',
                       command=self._goto('经纬度度分秒批量转换'))
        vm.add_command(label='常用单位换算', accelerator='Ctrl+U',
                       command=self._goto('常用单位换算'))
        # 隐藏玩法的鼠标入口：中文输入法激活时组合键会被 IME 吃掉，菜单始终可用
        vm.add_separator()
        vm.add_command(label='地层模式（界面变装）', command=self.easter_strata)
        vm.add_command(label='罗盘校准（小游戏）', command=self.easter_game)
        vm.add_command(label='野外记录卡', command=self.easter_field_card)
        vm.add_command(label='今日岩层', command=self.easter_rock)

        hm = self._menu(mb)
        mb.add_cascade(label='帮助', menu=hm)
        hm.add_command(label='帮助文档', accelerator='F1', command=self.show_doc)
        hm.add_command(label='更新说明', command=lambda: self.show_changelog())
        hm.add_command(label='检查更新…', command=lambda: self.check_update(True))
        hm.add_command(label='关于软件', command=self.show_about)

    def _goto(self, base):
        """按稳定页名跳转，避免把编号写死在菜单与快捷键里。"""
        return lambda: self.switch_to(self.func_index[base])

    def _bind_shortcuts(self):
        table = (
            ('<F1>', self.show_doc),
            ('<Control-o>', self.menu_import), ('<Control-O>', self.menu_import),
            ('<Control-s>', self.menu_export), ('<Control-S>', self.menu_export),
            ('<Control-w>', self.menu_clear), ('<Control-W>', self.menu_clear),
            ('<Control-Shift-C>', self.menu_copy),
            ('<Control-q>', self.menu_quit), ('<Control-Q>', self.menu_quit),
            ('<Control-h>', self._goto('计算历史')),
            ('<Control-H>', self._goto('计算历史')),
            ('<Control-d>', self._goto('经纬度度分秒批量转换')),
            ('<Control-D>', self._goto('经纬度度分秒批量转换')),
            ('<Control-u>', self._goto('常用单位换算')),
            ('<Control-U>', self._goto('常用单位换算')),
        )
        for key, fn in table:
            self.root.bind(key, lambda e, f=fn: (f(), 'break')[1])

    def toggle_glass(self):
        """视图菜单：在磨砂玻璃与经典平面之间即时切换。"""
        if getattr(self, 'strata_on', False):
            self.easter_strata()      # 先退出地层模式，避免两套配色打架
        on = self.glass_var.get()
        for w in self._glass_widgets:
            try:
                w.set_glass(on)
            except (tk.TclError, AttributeError):
                pass
        for w in self._glass_frames:
            try:
                w.configure(bg=GLASS_PANEL if on else PANEL)
            except (tk.TclError, AttributeError):
                pass
        try:
            ttk.Style().configure('Geo.Vertical.TScrollbar',
                                  troughcolor=GLASS_PANEL if on else PANEL)
        except tk.TclError:
            pass
        if self.active_idx >= 0:
            self.set_active(self.active_idx)
        self.root.update_idletasks()

    # ---------------- 菜单命令 ----------------
    def _active_base(self):
        i = self.active_idx
        return self.func_bases[i] if 0 <= i < len(self.func_bases) else ''

    def menu_import(self):
        base = self._active_base()
        if base == '经纬度度分秒批量转换':
            self.dms_import()
        elif base == '批量坐标正反算':
            self.bcoord_import()
        elif base == '平面坐标正反算':
            self.coord_batch.load_file()
        elif base == '经纬度→直角坐标':
            self.geo_batch.load_file()
        else:
            self._dialog('showinfo', '提示',
                         '导入功能用于「经纬度度分秒批量转换」「经纬度→直角坐标」'
                         '「平面坐标正反算」「批量坐标正反算」。\n请先切换到相应页面。')

    def menu_export(self):
        base = self._active_base()
        if base == '经纬度度分秒批量转换':
            self.dms_export()
        elif base == '批量坐标正反算':
            self.bcoord_export()
        elif base == '平面坐标正反算':
            self.coord_batch.export()
        elif base == '经纬度→直角坐标':
            self.geo_batch.export()
        elif base == '计算历史':
            self.history_export()
        else:
            self._dialog('showinfo', '提示',
                         '可导出表格的页面：\n'
                         '经纬度度分秒批量转换\n'
                         '经纬度 → 直角坐标\n'
                         '平面坐标正反算\n'
                         '批量坐标正反算\n'
                         '计算历史')

    def menu_copy(self):
        base = self._active_base()
        if base == '经纬度度分秒批量转换':
            self.dms_copy()
            return
        if base == '批量坐标正反算':
            self.bcoord_copy()
            return
        if base == '平面坐标正反算':
            self.coord_batch.copy()
            return
        if base == '经纬度→直角坐标':
            self.geo_batch.copy()
            return
        if base == '计算历史':
            self.history_copy()
            return
        frame = self.tab_frames.get(base)
        vals = []

        def walk(w):
            for ch in w.winfo_children():
                if isinstance(ch, tk.Label):
                    try:
                        if ch.cget('background') == READOUT_BG:
                            tv = ch.cget('textvariable')
                            v = self.root.getvar(tv) if tv else ''
                            if v and v != '—':
                                vals.append(str(v))
                    except (tk.TclError, AttributeError):
                        pass
                walk(ch)

        if frame is not None:
            walk(frame)
        if not vals:
            self._dialog('showinfo', '提示', '当前页暂无可复制的结果。')
            return
        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(vals))
        self._dialog('showinfo', '提示', '结果已复制到剪贴板。')

    def menu_clear(self):
        i = self.active_idx
        base = self._active_base()
        if base == '经纬度度分秒批量转换':
            self.dms_clear()
            return
        if base == '批量坐标正反算':
            self.bcoord_clear()
            return
        if base == '计算历史':
            self.history_clear()
            return
        if 0 <= i < len(self.func_bases):
            frame = self.tab_frames.get(self.func_bases[i])
            if frame is not None:
                self._clear_tab(frame)

    def menu_quit(self):
        self.root.destroy()

    # ---------------- 卡片 / 标题 / 字段 / 结果 辅助 ----------------
    def card(self):
        shadow = tk.Frame(self.scroll_box, bg='#cfd7e0')   # 投影底（右侧/下侧露出）
        base = tk.Frame(shadow, bg=CARD)
        base.pack(fill='both', expand=True, padx=(0, 2), pady=(0, 2))
        inner = tk.Frame(base, bg=CARD)
        inner.pack(fill='x', padx=22, pady=18)
        return shadow, inner

    def title(self, parent, text):
        head = tk.Frame(parent, bg=CARD)
        head.pack(anchor='w', fill='x', pady=(0, 10))
        bar = tk.Frame(head, bg=ACCENT, width=4, height=18)
        bar.pack(side='left', padx=(0, 9))
        bar.pack_propagate(False)
        tk.Label(head, text=text, bg=CARD, fg=INK,
                 font=('微软雅黑', 13, 'bold'), anchor='w').pack(side='left')

    def field(self, parent, label, width=16):
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill='x', pady=4)
        tk.Label(row, text=label, bg=CARD, fg=INK, font=('微软雅黑', 10),
                 width=16, anchor='w').pack(side='left')
        e = ttk.Entry(row, width=width, style='Geo.TEntry')
        e.pack(side='left', padx=8)
        return e

    def readout(self, parent, var):
        box = tk.Frame(parent, bg=READOUT_BG, highlightbackground=LINE,
                       highlightthickness=1)
        box.pack(fill='x', pady=(10, 2))
        tk.Label(box, textvariable=var, bg=READOUT_BG, fg=RESULT,
                 font=('Consolas', 11, 'bold'), anchor='w',
                 padx=12, pady=8).pack(fill='x')
        return box

    # ---------------- 构建全部标签页 ----------------
    def build_all_tabs(self):
        # 1. 视倾角↔真倾角
        c, f1 = self.card()
        self.title(f1, '视倾角 ↔ 真倾角')
        self.mode_dip = tk.StringVar(value='true2view')
        rf = tk.Frame(f1, bg=CARD); rf.pack(fill='x', pady=6)
        grp = GlassOptionGroup(
            rf, self.mode_dip,
            [('真倾角 → 视倾角', 'true2view'), ('视倾角 → 真倾角', 'view2true')],
            command=lambda: self.calc_dip(None))
        grp.pack(anchor='w')
        self._glass_widgets.append(grp)
        self._glass_options.append(grp)
        self.e1_sec = self.field(f1, '剖面方向(°)')
        self.e1_az = self.field(f1, '岩层倾向(°)')
        self.e1_dip = self.field(f1, '倾角(°)')
        self.res1 = tk.StringVar(value='—')
        self.readout(f1, self.res1)
        for w in (self.e1_sec, self.e1_az, self.e1_dip):
            w.bind('<KeyRelease>', self.calc_dip)
        self._page(c, '视倾角↔真倾角互算')

        # 2. 地层真/视厚度
        c, f2 = self.card()
        self.title(f2, '地层真/视厚度')
        self.e2_L = self.field(f2, '实测长度(m)')
        self.e2_a = self.field(f2, '真倾角(°)')
        self.e2_t = self.field(f2, '与走向夹角(°)')
        self.res2 = tk.StringVar(value='—')
        self.readout(f2, self.res2)
        for w in (self.e2_L, self.e2_a, self.e2_t):
            w.bind('<KeyRelease>', self.calc_thick)
        self._page(c, '地层真/视厚度')

        # 3. 走向↔倾向
        c, f3 = self.card()
        self.title(f3, '走向 ↔ 倾向')
        self.e3_s = self.field(f3, '走向(°)')
        self.res3s = tk.StringVar(value='—')
        self.readout(f3, self.res3s)
        tk.Frame(f3, bg=LINE, height=1).pack(fill='x', pady=8)
        self.e3_d = self.field(f3, '倾向(°)')
        self.res3d = tk.StringVar(value='—')
        self.readout(f3, self.res3d)
        self.e3_s.bind('<KeyRelease>', self.calc_strike)
        self.e3_d.bind('<KeyRelease>', self.calc_strike)
        self._page(c, '走向倾向互换')

        # 4. 经纬度度分秒批量转换（全新功能）
        self.build_dms_tab()

        # 5. 坡度高差平距
        c, f5 = self.card()
        self.title(f5, '坡度 高差 平距')
        self.e5_ang = self.field(f5, '坡面坡角(°)')
        self.e5_len = self.field(f5, '斜坡斜距(m)')
        self.res5 = tk.StringVar(value='—')
        self.readout(f5, self.res5)
        for w in (self.e5_ang, self.e5_len):
            w.bind('<KeyRelease>', self.calc_slope)
        self._page(c, '坡度高差平距')

        # 6. 平面坐标正反算
        c, f6 = self.card()
        self.title(f6, '平面坐标正反算')
        self.e6x0 = self.field(f6, '起点X')
        self.e6y0 = self.field(f6, '起点Y')
        self.e6az = self.field(f6, '方位角(°)')
        self.e6dis = self.field(f6, '水平距离(m)')
        self.res6 = tk.StringVar(value='—')
        self.readout(f6, self.res6)
        tk.Frame(f6, bg=LINE, height=1).pack(fill='x', pady=8)
        tk.Label(f6, text='两点反算 X1 Y1 X2 Y2', bg=CARD, fg=MUTED,
                 font=('微软雅黑', 10)).pack(anchor='w', pady=2)
        box = tk.Frame(f6, bg=CARD); box.pack(fill='x', pady=4)
        self.e6x1 = ttk.Entry(box, width=8, style='Geo.TEntry'); self.e6x1.pack(side='left', padx=3)
        self.e6y1 = ttk.Entry(box, width=8, style='Geo.TEntry'); self.e6y1.pack(side='left', padx=3)
        self.e6x2 = ttk.Entry(box, width=8, style='Geo.TEntry'); self.e6x2.pack(side='left', padx=3)
        self.e6y2 = ttk.Entry(box, width=8, style='Geo.TEntry'); self.e6y2.pack(side='left', padx=3)
        self.res6rev = tk.StringVar(value='—')
        self.readout(f6, self.res6rev)
        for w in (self.e6x0, self.e6y0, self.e6az, self.e6dis, self.e6x1, self.e6y1, self.e6x2, self.e6y2):
            w.bind('<KeyRelease>', self.calc_coord)

        # 批量区：与「经纬度度分秒批量转换」一致的四步操作，正算 / 反算共用一张表
        tk.Frame(f6, bg=LINE, height=1).pack(fill='x', pady=10)
        tk.Label(f6, text='批量换算', bg=CARD, fg=INK,
                 font=('微软雅黑', 11, 'bold')).pack(anchor='w')
        bm = tk.Frame(f6, bg=CARD); bm.pack(fill='x', pady=(6, 0))
        tk.Label(bm, text='批量模式', bg=CARD, fg=INK, font=('微软雅黑', 10),
                 width=10, anchor='w').pack(side='left')
        self.coord_batch_mode = tk.StringVar(value='forward')
        gb = GlassOptionGroup(
            bm, self.coord_batch_mode,
            [('正算（X Y 方位角 距离）', 'forward'),
             ('反算（X1 Y1 X2 Y2）', 'reverse')],
            command=self._coord_batch_mode_changed)
        gb.pack(side='left')
        self._glass_widgets.append(gb)
        self._glass_options.append(gb)

        self.coord_batch = BatchPanel(
            f6, self, *_COORD_BATCH_FWD, self._coord_batch_parse,
            '平面坐标换算结果',
            hint='每行一组，空格 / 逗号 / 制表符分隔；正算为「起点X 起点Y 方位角 距离」，'
                 '反算为「X1 Y1 X2 Y2」。方位角自北顺时针起算。',
            text_height=5, tree_height=6)
        self.coord_batch.pack(fill='x', pady=(2, 2))
        self._page(c, '平面坐标正反算')

        # 7. 方位角前后视
        c, f7 = self.card()
        self.title(f7, '方位角前后视推算')
        self.e7b = self.field(f7, '后视方位角：')
        self.res7f = tk.StringVar(value='—')
        self.readout(f7, self.res7f)
        tk.Frame(f7, bg=LINE, height=1).pack(fill='x', pady=8)
        self.e7f = self.field(f7, '前视方位角：')
        self.res7b = tk.StringVar(value='—')
        self.readout(f7, self.res7b)
        self.e7b.bind('<KeyRelease>', self.calc_az)
        self.e7f.bind('<KeyRelease>', self.calc_az)
        self._page(c, '方位角前后视')

        # 8. 平距↔斜距
        c, f8 = self.card()
        self.title(f8, '平距 ↔ 斜距')
        self.e8a = self.field(f8, '坡角(°)')
        self.e8h = self.field(f8, '输入水平平距(m)')
        self.res8s = tk.StringVar(value='—')
        self.readout(f8, self.res8s)
        tk.Frame(f8, bg=LINE, height=1).pack(fill='x', pady=8)
        self.e8l = self.field(f8, '输入斜坡斜距(m)')
        self.res8h = tk.StringVar(value='—')
        self.readout(f8, self.res8h)
        for w in (self.e8a, self.e8h, self.e8l):
            w.bind('<KeyRelease>', self.calc_hs)
        self._page(c, '平距斜距换算')

        # 9. 等高线坡角
        c, f9 = self.card()
        self.title(f9, '等高线坡角计算')
        self.e9h = self.field(f9, '等高距(m)')
        self.e9d = self.field(f9, '图上平距(m)')
        self.res9 = tk.StringVar(value='—')
        self.readout(f9, self.res9)
        for w in (self.e9h, self.e9d):
            w.bind('<KeyRelease>', self.calc_contour)
        self._page(c, '等高线坡角')

        # 10. 坡度↔坡率
        c, f10 = self.card()
        self.title(f10, '坡度 ↔ 坡率')
        self.e10a = self.field(f10, '坡角(°) → 坡率1:x')
        self.res10r = tk.StringVar(value='—')
        self.readout(f10, self.res10r)
        tk.Frame(f10, bg=LINE, height=1).pack(fill='x', pady=8)
        self.e10r = self.field(f10, '坡率1:x → 坡角')
        self.res10a = tk.StringVar(value='—')
        self.readout(f10, self.res10a)
        self.e10a.bind('<KeyRelease>', self.calc_rate)
        self.e10r.bind('<KeyRelease>', self.calc_rate)
        self._page(c, '坡度坡率互转')

        # 11. 钻孔分层真厚度
        c, f11 = self.card()
        self.title(f11, '钻孔分层真厚度')
        self.e11da = self.field(f11, '钻孔倾角(°)')
        self.e11l = self.field(f11, '岩芯进尺(m)')
        self.e11ra = self.field(f11, '岩层倾角(°)')
        self.e11ad = self.field(f11, '钻孔-倾向夹角(°)')
        self.res11 = tk.StringVar(value='—')
        self.readout(f11, self.res11)
        for w in (self.e11da, self.e11l, self.e11ra, self.e11ad):
            w.bind('<KeyRelease>', self.calc_dthick)
        self._page(c, '钻孔分层真厚度')

        # 12. 钻孔垂直埋深
        c, f12 = self.card()
        self.title(f12, '钻孔垂直埋深')
        self.e12l = self.field(f12, '钻孔斜段长(m)')
        self.e12a = self.field(f12, '钻孔倾角(°)')
        self.res12 = tk.StringVar(value='—')
        self.readout(f12, self.res12)
        for w in (self.e12l, self.e12a):
            w.bind('<KeyRelease>', self.calc_ddepth)
        self._page(c, '钻孔垂直埋深')

        # 13. 钻孔水平投影
        c, f13 = self.card()
        self.title(f13, '钻孔水平投影长度')
        self.e13l = self.field(f13, '斜长(m)')
        self.e13a = self.field(f13, '钻孔倾角(°)')
        self.res13 = tk.StringVar(value='—')
        self.readout(f13, self.res13)
        for w in (self.e13l, self.e13a):
            w.bind('<KeyRelease>', self.calc_dhor)
        self._page(c, '钻孔水平投影')

        # 14. 多段钻孔累计
        c, f14 = self.card()
        self.title(f14, '多段钻孔累计')
        for lbl, (le, ae) in [
            ('段1', ('e14l1', 'e14a1')), ('段2', ('e14l2', 'e14a2'))]:
            row = tk.Frame(f14, bg=CARD); row.pack(fill='x', pady=4)
            tk.Label(row, text=lbl + '长', bg=CARD, fg=INK,
                     font=('微软雅黑', 10), width=16, anchor='w').pack(side='left')
            ent_l = ttk.Entry(row, width=10, style='Geo.TEntry'); ent_l.pack(side='left', padx=8)
            tk.Label(row, text='倾角', bg=CARD, fg=INK, font=('微软雅黑', 10)).pack(side='left', padx=(6, 2))
            ent_a = ttk.Entry(row, width=10, style='Geo.TEntry'); ent_a.pack(side='left', padx=8)
            setattr(self, le, ent_l); setattr(self, ae, ent_a)
        self.res14 = tk.StringVar(value='—')
        self.readout(f14, self.res14)
        for w in (self.e14l1, self.e14a1, self.e14l2, self.e14a2):
            w.bind('<KeyRelease>', self.calc_dsum)
        self._page(c, '多段钻孔累计')

        # 15. 钻孔视厚度校正
        c, f15 = self.card()
        self.title(f15, '钻孔岩层视厚度校正')
        self.e15c = self.field(f15, '岩芯实测长(m)')
        self.e15d = self.field(f15, '钻孔倾角(°)')
        self.e15r = self.field(f15, '岩层倾角(°)')
        self.res15 = tk.StringVar(value='—')
        self.readout(f15, self.res15)
        for w in (self.e15c, self.e15d, self.e15r):
            w.bind('<KeyRelease>', self.calc_dview)
        self._page(c, '钻孔视厚度校正')

        # 16. 钻孔真倾角校正
        c, f16 = self.card()
        self.title(f16, '钻孔产真倾角校正')
        self.e16v = self.field(f16, '孔内视倾角(°)')
        self.e16t = self.field(f16, '钻孔与走向夹角(°)')
        self.res16 = tk.StringVar(value='—')
        self.readout(f16, self.res16)
        for w in (self.e16v, self.e16t):
            w.bind('<KeyRelease>', self.calc_dtrue)
        self._page(c, '钻孔真倾角校正')

        # 17. 地表岩层埋深
        c, f17 = self.card()
        self.title(f17, '地表岩层埋深')
        self.e17d = self.field(f17, '露头水平距(m)')
        self.e17a = self.field(f17, '岩层倾角(°)')
        self.res17 = tk.StringVar(value='—')
        self.readout(f17, self.res17)
        for w in (self.e17d, self.e17a):
            w.bind('<KeyRelease>', self.calc_buried)
        self._page(c, '地表岩层埋深')

        # 18. 露头出露宽度
        c, f18 = self.card()
        self.title(f18, '露头出露宽度')
        self.e18t = self.field(f18, '真厚度(m)')
        self.e18a = self.field(f18, '岩层倾角(°)')
        self.e18t2 = self.field(f18, '剖面夹角(°)')
        self.res18 = tk.StringVar(value='—')
        self.readout(f18, self.res18)
        for w in (self.e18t, self.e18a, self.e18t2):
            w.bind('<KeyRelease>', self.calc_outcrop)
        self._page(c, '露头出露宽度')

        # 19. 岩层水平厚度
        c, f19 = self.card()
        self.title(f19, '岩层水平厚度')
        self.e19t = self.field(f19, '地层真厚度(m)')
        self.e19a = self.field(f19, '岩层倾角(°)')
        self.res19 = tk.StringVar(value='—')
        self.readout(f19, self.res19)
        for w in (self.e19t, self.e19a):
            w.bind('<KeyRelease>', self.calc_hthick)
        self._page(c, '岩层水平厚度')

        # 20. 导线三维XYZ增量
        c, f20 = self.card()
        self.title(f20, '导线三维XYZ增量')
        self.e20az = self.field(f20, '方位角(°)')
        self.e20sl = self.field(f20, '坡角(°)仰正俯负')
        self.e20l = self.field(f20, '斜距(m)')
        self.res20 = tk.StringVar(value='—')
        self.readout(f20, self.res20)
        for w in (self.e20az, self.e20sl, self.e20l):
            w.bind('<KeyRelease>', self.calc_xyz)
        self._page(c, '导线三维XYZ增量')

        # 21. 面积单位换算
        c, f21 = self.card()
        self.title(f21, '面积单位换算')
        self.e21 = self.field(f21, '平方米㎡')
        self.res21 = tk.StringVar(value='—')
        self.readout(f21, self.res21)
        self.e21.bind('<KeyRelease>', self.calc_area)
        self._page(c, '面积单位换算')

        # 22. 块段矿石储量
        c, f22 = self.card()
        self.title(f22, '块段矿石储量')
        row = tk.Frame(f22, bg=CARD); row.pack(fill='x', pady=4)
        for lbl, attr in [('长', 'e22l'), ('宽', 'e22w'), ('厚', 'e22t'), ('体重', 'e22d')]:
            tk.Label(row, text=lbl, bg=CARD, fg=INK, font=('微软雅黑', 10)).pack(side='left', padx=(0, 4))
            ent = ttk.Entry(row, width=8, style='Geo.TEntry'); ent.pack(side='left', padx=(0, 10))
            setattr(self, attr, ent)
        self.res22 = tk.StringVar(value='—')
        self.readout(f22, self.res22)
        for w in (self.e22l, self.e22w, self.e22t, self.e22d):
            w.bind('<KeyRelease>', self.calc_reserve)
        self._page(c, '块段矿石储量')

        # 23. 岩土稳定坡角
        c, f23 = self.card()
        self.title(f23, '岩土稳定坡角')
        self.e23 = self.field(f23, '土内摩擦角φ(°)')
        self.res23 = tk.StringVar(value='—')
        self.readout(f23, self.res23)
        self.e23.bind('<KeyRelease>', self.calc_stable)
        self._page(c, '岩土稳定坡角')

        # 24. 坐标格网长度
        c, f24 = self.card()
        self.title(f24, '坐标格网长度')
        self.e24g = self.field(f24, '实地格网边长(m)')
        self.e24s = self.field(f24, '比例尺分母')
        self.res24 = tk.StringVar(value='—')
        self.readout(f24, self.res24)
        for w in (self.e24g, self.e24s):
            w.bind('<KeyRelease>', self.calc_grid)
        self._page(c, '坐标格网长度')

        # 25. 实测剖面厚度
        c, f25 = self.card()
        self.title(f25, '实测剖面厚度')
        self.e25_az = self.field(f25, '剖面导线方位(°)')
        self.e25_slope = self.field(f25, '导线坡度角(°)仰正俯负')
        self.e25_L = self.field(f25, '导线斜距(m)')
        self.e25_tend = self.field(f25, '岩层倾向(°)')
        self.e25_dip = self.field(f25, '岩层倾角(°)')
        self.res25 = tk.StringVar(value='—')
        self.readout(f25, self.res25)
        for w in (self.e25_az, self.e25_slope, self.e25_L, self.e25_tend, self.e25_dip):
            w.bind('<KeyRelease>', self.calc_profile_new)
        self._page(c, '实测剖面厚度')

        # 26. 批量坐标正反算
        self.build_batchcoord_tab()
        # 27. 常用单位换算
        self.build_unit_tab()
        # 28. 计算历史
        self.build_history_tab()
        # 29. 经纬度 → 直角坐标
        self.build_geo2xyz_tab()
        # 给 25 个单功能页统一加「清空输入」
        self._add_clear_buttons()

    def _add_clear_buttons(self):
        """单功能页右下角加一个低调的「清空输入」，清掉本页输入框与结果读数。

        批量页（自带「清空」）与本页已自备按钮的页面跳过，避免出现两个清空键。
        """
        skip = {'经纬度度分秒批量转换', '批量坐标正反算', '计算历史',
                '经纬度→直角坐标', '常用单位换算'}
        for base in self.func_bases:
            if base in skip:
                continue
            frame = self.tab_frames.get(base)
            if frame is None:
                continue
            try:
                inner = frame.winfo_children()[0].winfo_children()[0]
            except (IndexError, tk.TclError):
                continue
            bar = tk.Frame(inner, bg=CARD)
            bar.pack(fill='x', pady=(8, 2))
            ttk.Button(bar, text='清空输入', style='Ghost.TButton',
                       command=lambda fr=inner: self._clear_tab(fr)).pack(side='right')

    def _clear_tab(self, frame):
        """递归清空某个功能页里的所有输入框，并把结果读数复位为中性占位。"""
        def walk(w):
            for ch in w.winfo_children():
                if isinstance(ch, BatchPanel):
                    ch.clear()
                    continue
                if isinstance(ch, (tk.Entry, ttk.Entry)):
                    try:
                        ch.delete(0, tk.END)
                    except tk.TclError:
                        pass
                elif isinstance(ch, tk.Text):
                    try:
                        ch.delete('1.0', tk.END)
                    except tk.TclError:
                        pass
                if isinstance(ch, tk.Label):
                    try:
                        if ch.cget('background') == READOUT_BG and ch.cget('textvariable'):
                            self.root.setvar(ch.cget('textvariable'), '—')
                    except (tk.TclError, AttributeError):
                        pass
                walk(ch)
        walk(frame)

    # ---------------- 第4功能：经纬度度分秒批量转换 ----------------
    def build_dms_tab(self):
        c, f = self.card()
        self.title(f, '经纬度 · 度分秒批量转换')

        tk.Label(f, text='支持批量处理：粘贴或导入文本，软件自动识别经度与纬度，并双向转换为「十进制度」与「度分秒」。',
                 bg=CARD, fg=MUTED, font=('微软雅黑', 9), wraplength=620, justify='left').pack(anchor='w', pady=(0, 6))

        # 输入区
        in_frame = tk.LabelFrame(f, text=' 坐标文本（每行一条，可导入 .txt/.csv） ',
                                 bg=CARD, fg=INK, font=('微软雅黑', 10),
                                 relief='groove', bd=1, highlightbackground=LINE)
        in_frame.pack(fill='x', pady=4)
        self.dms_text = scrolledtext.ScrolledText(in_frame, height=8,
                                                  font=('Consolas', 10), bg='#fbfcfd',
                                                  relief='flat', bd=0)
        self.dms_text.pack(fill='both', expand=True, padx=6, pady=6)
        # 输入框默认留空，仅给出格式说明，不预填任何示例数值
        tk.Label(f, text='每行一组：经度在前、纬度在后；支持十进制度 / 度分秒 / 中文或字母方位，逗号、空格、制表符分隔。',
                 bg=CARD, fg=MUTED, font=('微软雅黑', 9), wraplength=640,
                 justify='left').pack(anchor='w', pady=(4, 0))

        # 按钮行
        btn_row = tk.Frame(f, bg=CARD); btn_row.pack(fill='x', pady=8)
        ttk.Button(btn_row, text='导入文本', style='Primary.TButton',
                   command=self.dms_import).pack(side='left', padx=4)
        ttk.Button(btn_row, text='开始转换', style='Primary.TButton',
                   command=self.dms_run).pack(side='left', padx=4)
        ttk.Button(btn_row, text='清空', style='Ghost.TButton',
                   command=self.dms_clear).pack(side='left', padx=4)
        ttk.Button(btn_row, text='复制结果', style='Ghost.TButton',
                   command=self.dms_copy).pack(side='left', padx=4)
        ttk.Button(btn_row, text='导出结果', style='Ghost.TButton',
                   command=self.dms_export).pack(side='left', padx=4)

        # 结果表
        cols = ('idx', 'lng_raw', 'lat_raw', 'lng_dec', 'lat_dec', 'lng_dms', 'lat_dms')
        self.dms_tree = ttk.Treeview(f, columns=cols, show='headings',
                                     style='Geo.Treeview', height=10)
        headers = {
            'idx': '序号', 'lng_raw': '经度(原值)', 'lat_raw': '纬度(原值)',
            'lng_dec': '经度(十进制度)', 'lat_dec': '纬度(十进制度)',
            'lng_dms': '经度(度分秒)', 'lat_dms': '纬度(度分秒)',
        }
        widths = {'idx': 50, 'lng_raw': 130, 'lat_raw': 130,
                  'lng_dec': 120, 'lat_dec': 120, 'lng_dms': 150, 'lat_dms': 150}
        for col in cols:
            self.dms_tree.heading(col, text=headers[col])
            self.dms_tree.column(col, width=widths[col], anchor='center')
        vsb = ttk.Scrollbar(f, orient='vertical', command=self.dms_tree.yview)
        self.dms_tree.configure(yscrollcommand=vsb.set)
        self.dms_tree.pack(side='left', fill='both', expand=True, pady=6)
        vsb.pack(side='right', fill='y', pady=6)

        self.dms_status = tk.StringVar(value='等待输入…')
        tk.Label(f, textvariable=self.dms_status, bg=CARD, fg=MUTED,
                 font=('微软雅黑', 9)).pack(anchor='w', pady=(2, 0))

        self.dms_results = []  # 保存最近一次结果，供复制/导出
        self._page(c, '经纬度度分秒批量转换')

    def dms_import(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            title='导入坐标文本',
            filetypes=[('文本文件', '*.txt'), ('CSV 文件', '*.csv'), ('所有文件', '*.*')])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as fh:
                content = fh.read()
        except UnicodeDecodeError:
            with open(path, 'r', encoding='gbk') as fh:
                content = fh.read()
        self.dms_text.delete('1.0', tk.END)
        self.dms_text.insert(tk.END, content)
        self.dms_status.set('已导入：' + path.split('/')[-1])
        self.dms_run()

    def dms_run(self):
        text = self.dms_text.get('1.0', tk.END)
        lines = text.splitlines()
        self.dms_tree.delete(*self.dms_tree.get_children())
        self.dms_results = []
        ok = 0
        skip = 0
        for ln in lines:
            if not ln.strip():
                continue
            row = convert_line(ln)
            if row is None:
                skip += 1
                continue
            ok += 1
            idx = len(self.dms_results) + 1
            self.dms_results.append(row)
            self.dms_tree.insert('', tk.END, values=(
                idx,
                row['lng_raw'], row['lat_raw'],
                f"{row['lng_dec']:.6f}" if row['lng_dec'] is not None else '',
                f"{row['lat_dec']:.6f}" if row['lat_dec'] is not None else '',
                row['lng_dms'], row['lat_dms'],
            ))
        self.dms_status.set(f'已转换 {ok} 条，跳过无法识别 {skip} 行。')
        if ok:
            self.log_history('经纬度批量转换', f'转换 {ok} 条')
        self._coord_easter()

    def dms_clear(self):
        self.dms_text.delete('1.0', tk.END)
        self.dms_tree.delete(*self.dms_tree.get_children())
        self.dms_results = []
        self.dms_status.set('已清空。')

    def dms_copy(self):
        if not self.dms_results:
            self._dialog('showinfo', '提示', '暂无可复制的结果，请先转换。')
            return
        lines = ['序号\t经度(原值)\t纬度(原值)\t经度(十进制度)\t纬度(十进制度)\t经度(度分秒)\t纬度(度分秒)']
        for i, r in enumerate(self.dms_results, 1):
            lines.append('\t'.join([
                str(i), r['lng_raw'], r['lat_raw'],
                f"{r['lng_dec']:.6f}" if r['lng_dec'] is not None else '',
                f"{r['lat_dec']:.6f}" if r['lat_dec'] is not None else '',
                r['lng_dms'], r['lat_dms'],
            ]))
        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(lines))
        self.dms_status.set('结果已复制到剪贴板。')

    def dms_export(self):
        if not self.dms_results:
            self._dialog('showinfo', '提示', '暂无可导出的结果，请先转换。')
            return
        heads = ['序号', '经度(原值)', '纬度(原值)', '经度(十进制度)', '纬度(十进制度)', '经度(度分秒)', '纬度(度分秒)']
        rows = []
        for i, r in enumerate(self.dms_results, 1):
            rows.append([i, r['lng_raw'], r['lat_raw'],
                f"{r['lng_dec']:.6f}" if r['lng_dec'] is not None else '',
                f"{r['lat_dec']:.6f}" if r['lat_dec'] is not None else '',
                r['lng_dms'], r['lat_dms']])
        res = self.export_table(heads, rows, '经纬度转换结果')
        if res:
            self.dms_status.set('已导出：' + res[1].split('/')[-1])

    # ---------------- 第26功能：批量坐标正反算 ----------------
    def build_batchcoord_tab(self):
        c, f = self.card()
        self.title(f, '批量坐标正反算')

        tk.Label(f, text='批量处理多组坐标：粘贴或导入文本，每行一组。软件按行批量正算或反算，结果表格化，可复制/导出。',
                 bg=CARD, fg=MUTED, font=('微软雅黑', 9), wraplength=640, justify='left').pack(anchor='w', pady=(0, 6))

        self.bcoord_mode = tk.StringVar(value='forward')
        rf = tk.Frame(f, bg=CARD); rf.pack(fill='x', pady=6)
        grp = GlassOptionGroup(
            rf, self.bcoord_mode,
            [('正算（起点+方位+距离 → 终点）', 'forward'),
             ('反算（两点 → 方位+距离）', 'reverse')],
            command=self.bcoord_set_columns)
        grp.pack(anchor='w')
        self._glass_widgets.append(grp)
        self._glass_options.append(grp)

        in_frame = tk.LabelFrame(f, text=' 坐标文本（每行一组，空格/逗号/制表符分隔） ',
                                 bg=CARD, fg=INK, font=('微软雅黑', 10),
                                 relief='groove', bd=1, highlightbackground=LINE)
        in_frame.pack(fill='x', pady=4)
        self.bcoord_text = scrolledtext.ScrolledText(in_frame, height=8,
                                                     font=('Consolas', 10), bg='#fbfcfd',
                                                     relief='flat', bd=0)
        self.bcoord_text.pack(fill='both', expand=True, padx=6, pady=6)
        # 输入框默认留空，仅给出格式说明，不预填任何示例数值
        tk.Label(f, text='每行一组：正算填「起点X 起点Y 方位角 距离」，反算填「X1 Y1 X2 Y2」；空格、逗号或制表符分隔。',
                 bg=CARD, fg=MUTED, font=('微软雅黑', 9), wraplength=640,
                 justify='left').pack(anchor='w', pady=(4, 0))

        btn_row = tk.Frame(f, bg=CARD); btn_row.pack(fill='x', pady=8)
        ttk.Button(btn_row, text='导入文本', style='Primary.TButton',
                   command=self.bcoord_import).pack(side='left', padx=4)
        ttk.Button(btn_row, text='开始计算', style='Primary.TButton',
                   command=self.bcoord_run).pack(side='left', padx=4)
        ttk.Button(btn_row, text='清空', style='Ghost.TButton',
                   command=self.bcoord_clear).pack(side='left', padx=4)
        ttk.Button(btn_row, text='复制结果', style='Ghost.TButton',
                   command=self.bcoord_copy).pack(side='left', padx=4)
        ttk.Button(btn_row, text='导出结果', style='Ghost.TButton',
                   command=self.bcoord_export).pack(side='left', padx=4)

        self.bcoord_tree = ttk.Treeview(f, show='headings', style='Geo.Treeview', height=10)
        self.bcoord_set_columns()
        vsb = ttk.Scrollbar(f, orient='vertical', command=self.bcoord_tree.yview)
        self.bcoord_tree.configure(yscrollcommand=vsb.set)
        self.bcoord_tree.pack(side='left', fill='both', expand=True, pady=6)
        vsb.pack(side='right', fill='y', pady=6)

        self.bcoord_status = tk.StringVar(value='等待输入…')
        tk.Label(f, textvariable=self.bcoord_status, bg=CARD, fg=MUTED,
                 font=('微软雅黑', 9)).pack(anchor='w', pady=(2, 0))
        self.bcoord_results = []
        self._page(c, '批量坐标正反算')

    def bcoord_set_columns(self):
        mode = self.bcoord_mode.get()
        if mode == 'forward':
            cols = ('idx', 'x0', 'y0', 'az', 'dis', 'x1', 'y1')
            heads = {'idx': '序号', 'x0': '起点X', 'y0': '起点Y', 'az': '方位角(°)',
                     'dis': '距离(m)', 'x1': '终点X', 'y1': '终点Y'}
            widths = {'idx': 50, 'x0': 110, 'y0': 110, 'az': 100, 'dis': 100, 'x1': 120, 'y1': 120}
        else:
            cols = ('idx', 'x1', 'y1', 'x2', 'y2', 'az', 'dis')
            heads = {'idx': '序号', 'x1': '点1 X', 'y1': '点1 Y', 'x2': '点2 X',
                     'y2': '点2 Y', 'az': '方位角(°)', 'dis': '距离(m)'}
            widths = {'idx': 50, 'x1': 110, 'y1': 110, 'x2': 110, 'y2': 110, 'az': 100, 'dis': 100}
        self.bcoord_tree['columns'] = cols
        for col in cols:
            self.bcoord_tree.heading(col, text=heads[col])
            self.bcoord_tree.column(col, width=widths[col], anchor='center')
        self.bcoord_tree.delete(*self.bcoord_tree.get_children())
        self.bcoord_results = []

    def bcoord_import(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            title='导入坐标文本',
            filetypes=[('文本文件', '*.txt'), ('CSV 文件', '*.csv'), ('所有文件', '*.*')])
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as fh:
                content = fh.read()
        except UnicodeDecodeError:
            with open(path, 'r', encoding='gbk') as fh:
                content = fh.read()
        self.bcoord_text.delete('1.0', tk.END)
        self.bcoord_text.insert(tk.END, content)
        self.bcoord_status.set('已导入：' + path.split('/')[-1])
        self.bcoord_run()

    def bcoord_run(self):
        mode = self.bcoord_mode.get()
        text = self.bcoord_text.get('1.0', tk.END)
        self.bcoord_tree.delete(*self.bcoord_tree.get_children())
        self.bcoord_results = []
        ok = 0
        skip = 0
        for ln in text.splitlines():
            if not ln.strip():
                continue
            toks = re.split(r'[,;\t]+|\s+', ln.strip())
            toks = [t for t in toks if t != '']
            nums = []
            for t in toks:
                try:
                    nums.append(float(t))
                except ValueError:
                    pass
            if len(nums) < 4:
                skip += 1
                continue
            if mode == 'forward':
                x0, y0, az, dis = nums[0], nums[1], nums[2], nums[3]
                rad = math.radians(az)
                x1 = x0 + dis * math.sin(rad)
                y1 = y0 + dis * math.cos(rad)
                row = {'x0': x0, 'y0': y0, 'az': az, 'dis': dis, 'x1': x1, 'y1': y1}
                vals = (len(self.bcoord_results) + 1, f'{x0:.2f}', f'{y0:.2f}',
                        f'{az:.4f}', f'{dis:.2f}', f'{x1:.2f}', f'{y1:.2f}')
            else:
                x1, y1, x2, y2 = nums[0], nums[1], nums[2], nums[3]
                dx = x2 - x1
                dy = y2 - y1
                dis = math.hypot(dx, dy)
                az = math.degrees(math.atan2(dx, dy))
                az = az if az >= 0 else az + 360
                row = {'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2, 'az': az, 'dis': dis}
                vals = (len(self.bcoord_results) + 1, f'{x1:.2f}', f'{y1:.2f}',
                        f'{x2:.2f}', f'{y2:.2f}', f'{az:.4f}', f'{dis:.2f}')
            ok += 1
            self.bcoord_results.append(row)
            self.bcoord_tree.insert('', tk.END, values=vals)
        self.bcoord_status.set(f'已计算 {ok} 组，跳过格式不符 {skip} 行。')
        if ok:
            self.log_history('批量坐标' + ('正算' if mode == 'forward' else '反算'), f'计算 {ok} 组')

    def bcoord_clear(self):
        self.bcoord_text.delete('1.0', tk.END)
        self.bcoord_tree.delete(*self.bcoord_tree.get_children())
        self.bcoord_results = []
        self.bcoord_status.set('已清空。')

    def bcoord_copy(self):
        if not self.bcoord_results:
            self._dialog('showinfo', '提示', '暂无可复制的结果，请先计算。')
            return
        mode = self.bcoord_mode.get()
        if mode == 'forward':
            heads = ['序号', '起点X', '起点Y', '方位角', '距离', '终点X', '终点Y']
        else:
            heads = ['序号', '点1X', '点1Y', '点2X', '点2Y', '方位角', '距离']
        lines = ['\t'.join(heads)]
        for i, r in enumerate(self.bcoord_results, 1):
            if mode == 'forward':
                lines.append('\t'.join([str(i), f"{r['x0']:.6f}", f"{r['y0']:.6f}",
                                        f"{r['az']:.6f}", f"{r['dis']:.6f}",
                                        f"{r['x1']:.6f}", f"{r['y1']:.6f}"]))
            else:
                lines.append('\t'.join([str(i), f"{r['x1']:.6f}", f"{r['y1']:.6f}",
                                        f"{r['x2']:.6f}", f"{r['y2']:.6f}",
                                        f"{r['az']:.6f}", f"{r['dis']:.6f}"]))
        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(lines))
        self.bcoord_status.set('结果已复制到剪贴板。')

    def bcoord_export(self):
        if not self.bcoord_results:
            self._dialog('showinfo', '提示', '暂无可导出的结果，请先计算。')
            return
        mode = self.bcoord_mode.get()
        if mode == 'forward':
            heads = ['序号', '起点X', '起点Y', '方位角(°)', '距离(m)', '终点X', '终点Y']
            rows = [[i, f"{r['x0']:.6f}", f"{r['y0']:.6f}", f"{r['az']:.6f}",
                     f"{r['dis']:.6f}", f"{r['x1']:.6f}", f"{r['y1']:.6f}"]
                    for i, r in enumerate(self.bcoord_results, 1)]
        else:
            heads = ['序号', '点1X', '点1Y', '点2X', '点2Y', '方位角(°)', '距离(m)']
            rows = [[i, f"{r['x1']:.6f}", f"{r['y1']:.6f}", f"{r['x2']:.6f}",
                     f"{r['y2']:.6f}", f"{r['az']:.6f}", f"{r['dis']:.6f}"]
                    for i, r in enumerate(self.bcoord_results, 1)]
        res = self.export_table(heads, rows, '批量坐标结果')
        if res:
            self.bcoord_status.set('已导出：' + res[1].split('/')[-1])

    # ---------------- 第27功能：常用单位换算 ----------------
    def build_unit_tab(self):
        c, f = self.card()
        self.title(f, '常用单位换算')
        tk.Label(f, text='输入任一数值即时换算，结果实时显示。', bg=CARD, fg=MUTED,
                 font=('微软雅黑', 9)).pack(anchor='w', pady=(0, 6))

        grid = tk.Frame(f, bg=CARD)
        grid.pack(fill='x', pady=4)
        grid.columnconfigure(0, weight=1)
        convs = [
            ('角度(°)→弧度', lambda s: f'{math.radians(float(s)):.6f} rad'),
            ('弧度→角度(°)', lambda s: f'{math.degrees(float(s)):.4f} °'),
            ('米(m)→英尺(ft)', lambda s: f'{float(s) * 3.2808399:.4f} ft'),
            ('英尺(ft)→米(m)', lambda s: f'{float(s) / 3.2808399:.4f} m'),
            ('公里(km)→海里(nm)', lambda s: f'{float(s) * 0.5399568:.4f} nm'),
            ('海里(nm)→公里(km)', lambda s: f'{float(s) / 0.5399568:.4f} km'),
            ('公顷(ha)→亩', lambda s: f'{float(s) * 15:.4f} 亩'),
            ('亩→公顷(ha)', lambda s: f'{float(s) / 15:.4f} ha'),
            ('平方米(㎡)→亩', lambda s: f'{float(s) / 666.6667:.4f} 亩'),
            ('十进制度→度分秒', lambda s: self._dec2dms(float(s))),
            ('度分秒→十进制度', lambda s: self._dms2dec(s)),
        ]
        for i, (name, fn) in enumerate(convs):
            cell = tk.Frame(grid, bg=CARD)
            # 单列排布：卡片宽度有限，两列会把右侧结果框挤出可视区
            cell.grid(row=i, column=0, sticky='ew', pady=3)
            self._unit_row(cell, name, fn)
        tk.Label(f, text='注：度分秒支持 116°23′45″ / 116:23:45 / 116.391 等写法。',
                 bg=CARD, fg=MUTED, font=('微软雅黑', 9)).pack(anchor='w', pady=(6, 0))
        bar = tk.Frame(f, bg=CARD)
        bar.pack(fill='x', pady=(8, 2))
        ttk.Button(bar, text='清空输入', style='Ghost.TButton',
                   command=lambda: self._clear_tab(f)).pack(side='right')
        self._page(c, '常用单位换算')

    def _unit_row(self, parent, name, compute):
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill='x', pady=3)
        tk.Label(row, text=name, bg=CARD, fg=INK, font=('微软雅黑', 10),
                 width=18, anchor='w').pack(side='left')
        e = ttk.Entry(row, width=14, style='Geo.TEntry')
        e.pack(side='left', padx=8)
        var = tk.StringVar(value='—')
        box = tk.Frame(row, bg=READOUT_BG, highlightbackground=LINE, highlightthickness=1)
        box.pack(side='left', fill='x', expand=True)
        tk.Label(box, textvariable=var, bg=READOUT_BG, fg=RESULT,
                 font=('Consolas', 10), anchor='w', padx=8, pady=4).pack(fill='x')

        def on_key(ev=None):
            raw = e.get().strip()
            if not raw:
                var.set('—')
                return
            try:
                out = compute(raw)
                var.set(out)
                self.log_history('单位换算·' + name, f'{raw} → {out}')
            except (ValueError, TypeError):
                var.set('—')
        e.bind('<KeyRelease>', on_key)
        e.bind('<Return>', on_key)
        e.bind('<FocusOut>', on_key)
        e._refresh = on_key        # 便于程序化取值/测试
        e._result_var = var

    def _dec2dms(self, v):
        a = abs(v)
        d = int(a)
        mf = (a - d) * 60
        m = int(mf)
        s = (mf - m) * 60
        return f"{d}°{m}′{s:.2f}″"

    def _dms2dec(self, raw):
        p = parse_token(raw)
        if p is None:
            raise ValueError
        return f"{p['val']:.6f}°"

    # ---------------- 第28功能：计算历史 ----------------
    # ---------------- 第29功能：经纬度 → 直角坐标 ----------------
    def build_geo2xyz_tab(self):
        c, f = self.card()
        self.title(f, '经纬度 → 直角坐标')
        tk.Label(f, text='大地坐标(B, L, H) 换算地心空间直角坐标(X, Y, Z)，'
                         '同时给出高斯-克吕格投影平面直角坐标(x, y)。',
                 bg=CARD, fg=MUTED, font=('微软雅黑', 9),
                 wraplength=620, justify='left').pack(anchor='w', pady=(0, 8))

        row = tk.Frame(f, bg=CARD)
        row.pack(fill='x', pady=(0, 4))
        tk.Label(row, text='参考椭球', bg=CARD, fg=INK, font=('微软雅黑', 10),
                 width=10, anchor='w').pack(side='left')
        self.ellip29 = tk.StringVar(value='CGCS2000')
        g1 = GlassOptionGroup(row, self.ellip29,
                              [(k, k) for k in ELLIPSOIDS],
                              command=lambda: self.calc_geo2xyz(None))
        g1.pack(side='left')
        self._glass_widgets.append(g1)
        self._glass_options.append(g1)

        row2 = tk.Frame(f, bg=CARD)
        row2.pack(fill='x', pady=(0, 6))
        tk.Label(row2, text='分带方式', bg=CARD, fg=INK, font=('微软雅黑', 10),
                 width=10, anchor='w').pack(side='left')
        self.zone29 = tk.StringVar(value='3')
        g2 = GlassOptionGroup(row2, self.zone29,
                              [('3 度带', '3'), ('6 度带', '6')],
                              command=lambda: self.calc_geo2xyz(None))
        g2.pack(side='left')
        self._glass_widgets.append(g2)
        self._glass_options.append(g2)

        self.e29_b = self.field(f, '纬度 B(°)')
        self.e29_l = self.field(f, '经度 L(°)')
        self.e29_h = self.field(f, '大地高 H(m)')
        self.e29_l0 = self.field(f, '中央子午线 L0(°)')
        tk.Label(f, text='纬度/经度支持 39.907、39°54′25″、北纬39.907 等写法；'
                         '大地高留空按 0 m，中央子午线留空按所选分带自动推算。',
                 bg=CARD, fg=MUTED, font=('微软雅黑', 9),
                 wraplength=620, justify='left').pack(anchor='w', pady=(2, 0))

        self.res29xyz = tk.StringVar(value='—')
        self.readout(f, self.res29xyz)
        self.res29gauss = tk.StringVar(value='—')
        self.readout(f, self.res29gauss)
        self.res29info = tk.StringVar(value='—')
        self.readout(f, self.res29info)
        self.res29note = tk.StringVar(value='')
        tk.Label(f, textvariable=self.res29note, bg=CARD, fg=MUTED,
                 font=('微软雅黑', 9), wraplength=640, justify='left',
                 anchor='w').pack(fill='x', pady=(6, 0))

        for w in (self.e29_b, self.e29_l, self.e29_h, self.e29_l0):
            w.bind('<KeyRelease>', self.calc_geo2xyz)
            w.bind('<Return>', self.calc_geo2xyz)

        bar = tk.Frame(f, bg=CARD)
        bar.pack(fill='x', pady=(8, 2))
        ttk.Button(bar, text='清空输入', style='Ghost.TButton',
                   command=lambda: self._clear_geo2xyz(f)).pack(side='right')

        # 批量区：单点算完可以接着批量，椭球与分带沿用上面两项选择
        tk.Frame(f, bg=LINE, height=1).pack(fill='x', pady=10)
        tk.Label(f, text='批量换算', bg=CARD, fg=INK,
                 font=('微软雅黑', 11, 'bold')).pack(anchor='w')
        self.geo_batch = BatchPanel(
            f, self,
            ['序号', '纬度B(°)', '经度L(°)', '大地高(m)',
             'X(m)', 'Y(m)', 'Z(m)', 'x(m)', 'y(m)'],
            [50, 100, 110, 82, 110, 110, 110, 106, 106],
            self._geo_batch_parse, '经纬度直角坐标结果',
            hint='每行一组：纬度 经度 [大地高] [中央子午线]；椭球与分带方式沿用上方选项，'
                 '高程留空按 0 m、中央子午线留空按分带自动推算。列较多，'
                 '可拖动表格下方滑块查看右侧数值。',
            text_height=5, tree_height=6)
        self.geo_batch.pack(fill='x', pady=(4, 2))
        self._page(c, '经纬度→直角坐标')

    def _clear_geo2xyz(self, frame):
        self._clear_tab(frame)
        self.res29note.set('')

    def _geo_batch_parse(self, line):
        parts = split_batch_fields(line)
        if len(parts) < 2:
            return None
        pb, pl = parse_token(parts[0]), parse_token(parts[1])
        if pb is None or pl is None:
            return None
        B, L = pb['val'], pl['val']
        if not (-90.0 <= B <= 90.0) or not (-180.0 <= L <= 180.0):
            return None
        H = float(parts[2]) if len(parts) > 2 and parts[2] else 0.0
        ell = self.ellip29.get()
        width = int(self.zone29.get())
        if len(parts) > 3 and parts[3]:
            L0 = float(parts[3])
        else:
            _zone, L0 = zone_of(L, width)
        X, Y, Z = blh_to_xyz(B, L, H, ell)
        x, y = gauss_forward(B, L, L0, ell)
        return ['%.7f' % B, '%.7f' % L, '%.2f' % H,
                '%.3f' % X, '%.3f' % Y, '%.3f' % Z, '%.3f' % x, '%.3f' % y]

    def calc_geo2xyz(self, _e=None):
        try:
            bt = self.e29_b.get().strip()
            lt = self.e29_l.get().strip()
            if not bt or not lt:
                raise ValueError
            pb, pl = parse_token(bt), parse_token(lt)
            if pb is None or pl is None:
                raise ValueError
            B, L = pb['val'], pl['val']
            if not (-90.0 <= B <= 90.0):
                self.res29xyz.set('—')
                self.res29gauss.set('—')
                self.res29info.set('—')
                self.res29note.set('纬度需在 -90° ~ 90° 之间，请检查输入。')
                return
            if not (-180.0 <= L <= 180.0):
                self.res29xyz.set('—')
                self.res29gauss.set('—')
                self.res29info.set('—')
                self.res29note.set('经度需在 -180° ~ 180° 之间，请检查输入。')
                return

            ht = self.e29_h.get().strip()
            H = float(ht) if ht else 0.0
            ell = self.ellip29.get()
            width = int(self.zone29.get())

            l0t = self.e29_l0.get().strip()
            if l0t:
                L0, zone = float(l0t), None
            else:
                zone, L0 = zone_of(L, width)

            X, Y, Z = blh_to_xyz(B, L, H, ell)
            x, y = gauss_forward(B, L, L0, ell)

            self.res29xyz.set(f'X = {X:,.3f} m     Y = {Y:,.3f} m     Z = {Z:,.3f} m')
            self.res29gauss.set(f'x = {x:,.3f} m     y = {y:,.3f} m')
            if zone:
                self.res29info.set(f'第 {zone} 带 · L0 = {L0:.3f}° · 通用 y = {zone}{y:.2f}')
            else:
                self.res29info.set(f'手动 L0 = {L0:.3f}° · 未分带')

            dl = L - L0
            notes = []
            if abs(dl) > 3.5:
                notes.append(f'距中央子午线 {dl:+.2f}°，已超出投影分带范围，平面坐标仅作参考。')
            if B < 0:
                notes.append('南半球，x 为负值。')
            self.res29note.set(' '.join(notes))
        except (ValueError, TypeError, ZeroDivisionError, AttributeError):
            self.res29xyz.set('—')
            self.res29gauss.set('—')
            self.res29info.set('—')
            self.res29note.set('')

    def build_history_tab(self):
        c, f = self.card()
        self.title(f, '计算历史')
        tk.Label(f, text='自动记录本次会话所有计算结果（按功能去重），可复制或一键导出为 Excel / CSV。',
                 bg=CARD, fg=MUTED, font=('微软雅黑', 9), wraplength=640, justify='left').pack(anchor='w', pady=(0, 6))

        cols = ('time', 'cat', 'text')
        self.history_tree = ttk.Treeview(f, columns=cols, show='headings',
                                         style='Geo.Treeview', height=14)
        heads = {'time': '时间', 'cat': '功能', 'text': '结果'}
        widths = {'time': 90, 'cat': 150, 'text': 440}
        for col in cols:
            self.history_tree.heading(col, text=heads[col])
            self.history_tree.column(col, width=widths[col], anchor='w')
        vsb = ttk.Scrollbar(f, orient='vertical', command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=vsb.set)
        self.history_tree.pack(side='left', fill='both', expand=True, pady=6)
        vsb.pack(side='right', fill='y', pady=6)

        btn_row = tk.Frame(f, bg=CARD)
        btn_row.pack(fill='x', pady=8)
        ttk.Button(btn_row, text='复制全部', style='Ghost.TButton',
                   command=self.history_copy).pack(side='left', padx=4)
        ttk.Button(btn_row, text='导出', style='Ghost.TButton',
                   command=self.history_export).pack(side='left', padx=4)
        ttk.Button(btn_row, text='清空', style='Ghost.TButton',
                   command=self.history_clear).pack(side='left', padx=4)

        self.history_status = tk.StringVar(value='暂无记录。')
        tk.Label(f, textvariable=self.history_status, bg=CARD, fg=MUTED,
                 font=('微软雅黑', 9)).pack(anchor='w', pady=(2, 0))

        for h in self.history:
            self.history_tree.insert('', 0, values=(h['time'], h['cat'], h['text']))
        self._history_refresh_status()
        self._page(c, '计算历史')

    def _history_append_row(self, t, cat, text):
        self.history_tree.insert('', 0, values=(t, cat, text))
        self._history_refresh_status()

    def _history_refresh_status(self):
        n = len(self.history)
        self.history_status.set(f'共 {n} 条记录。' if n else '暂无记录。')

    def history_copy(self):
        if not self.history:
            self._dialog('showinfo', '提示', '暂无可复制的记录。')
            return
        lines = ['\t'.join(['时间', '功能', '结果'])]
        for h in self.history:
            lines.append('\t'.join([h['time'], h['cat'], h['text']]))
        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(lines))
        self.history_status.set('已复制到剪贴板。')

    def history_export(self):
        if not self.history:
            self._dialog('showinfo', '提示', '暂无可导出的记录。')
            return
        heads = ['时间', '功能', '结果']
        rows = [[h['time'], h['cat'], h['text']] for h in self.history]
        res = self.export_table(heads, rows, '计算历史')
        if res:
            self.history_status.set('已导出：' + res[1].split('/')[-1])

    def history_clear(self):
        self.history = []
        self._last_log = {}
        self.history_tree.delete(*self.history_tree.get_children())
        self.history_status.set('已清空。')

    # ---------------- 历史记录机制 ----------------
    def _wrap_calc_methods(self):
        calc_map = {
            'calc_dip': ('视倾角↔真倾角', ['res1']),
            'calc_thick': ('地层真/视厚度', ['res2']),
            'calc_strike': ('走向↔倾向', ['res3s', 'res3d']),
            'calc_slope': ('坡度高差平距', ['res5']),
            'calc_coord': ('平面坐标正反算', ['res6', 'res6rev']),
            'calc_az': ('方位角前后视', ['res7f', 'res7b']),
            'calc_hs': ('平距斜距换算', ['res8s', 'res8h']),
            'calc_contour': ('等高线坡角', ['res9']),
            'calc_rate': ('坡度坡率互转', ['res10r', 'res10a']),
            'calc_dthick': ('钻孔分层真厚度', ['res11']),
            'calc_ddepth': ('钻孔垂直埋深', ['res12']),
            'calc_dhor': ('钻孔水平投影', ['res13']),
            'calc_dsum': ('多段钻孔累计', ['res14']),
            'calc_dview': ('钻孔视厚度校正', ['res15']),
            'calc_dtrue': ('钻孔真倾角校正', ['res16']),
            'calc_buried': ('地表岩层埋深', ['res17']),
            'calc_outcrop': ('露头出露宽度', ['res18']),
            'calc_hthick': ('岩层水平厚度', ['res19']),
            'calc_xyz': ('导线三维XYZ增量', ['res20']),
            'calc_area': ('面积单位换算', ['res21']),
            'calc_reserve': ('块段矿石储量', ['res22']),
            'calc_stable': ('岩土稳定坡角', ['res23']),
            'calc_grid': ('坐标格网长度', ['res24']),
            'calc_profile_new': ('实测剖面厚度', ['res25']),
            'calc_geo2xyz': ('经纬度→直角坐标', ['res29xyz', 'res29gauss']),
        }
        for mname, (label, vns) in calc_map.items():
            setattr(self, mname, self._make_wrapper(mname, label, vns))

    def _make_wrapper(self, method_name, label, varnames):
        orig = getattr(self, method_name)

        def wrapper(e):
            orig(e)
            texts = []
            for vn in varnames:
                v = getattr(self, vn, None)
                if v is not None:
                    texts.append(v.get())
            self.log_history(label, '  '.join(texts))
        return wrapper

    def log_history(self, category, text):
        if self.is_blank_result(text):
            return
        if self._last_log.get(category) == text:
            return
        self._last_log[category] = text
        t = datetime.now().strftime('%H:%M:%S')
        self.history.append({'time': t, 'cat': category, 'text': text})
        if len(self.history) > 2000:
            self.history.pop(0)
        if hasattr(self, 'history_tree'):
            self._history_append_row(t, category, text)

    def is_blank_result(self, text):
        nums = re.findall(r'-?\d+\.?\d*', text)
        if not nums:
            return True
        try:
            return all(abs(float(n)) < 1e-9 for n in nums)
        except ValueError:
            return True

    def export_table(self, headers, rows, default_name):
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title='导出结果', defaultextension='.xlsx',
            filetypes=[('Excel 工作簿', '*.xlsx'), ('CSV 文件', '*.csv'), ('文本文件', '*.txt')])
        if not path:
            return None
        if path.lower().endswith('.xlsx'):
            try:
                from openpyxl import Workbook
                wb = Workbook()
                ws = wb.active
                ws.title = default_name[:31]
                ws.append(list(headers))
                for r in rows:
                    ws.append(list(r))
                wb.save(path)
                return ('xlsx', path)
            except Exception as ex:  # noqa
                self._dialog('showerror', '导出失败', str(ex))
                return None
        else:
            with open(path, 'w', encoding='utf-8-sig', newline='') as fh:
                w = csv.writer(fh)
                w.writerow(headers)
                for r in rows:
                    w.writerow(r)
            return ('file', path)

    # =====================================================================
    # 原有 24 个计算函数（计算逻辑保持不变）
    # =====================================================================
    def calc_dip(self, e):
        try:
            s = float(self.e1_sec.get())
            az = float(self.e1_az.get())
            d = float(self.e1_dip.get())
            delta = abs(az - s)
            delta = 360 - delta if delta > 180 else delta
            c = math.cos(math.radians(delta))
            if self.mode_dip.get() == 'true2view':
                res = math.degrees(math.atan(math.tan(math.radians(d)) * c))
                self.res1.set(f'视倾角：{res:.2f}°')
                return
            res = math.degrees(math.atan(math.tan(math.radians(d)) / c))
            self.res1.set(f'真倾角：{res:.2f}°')
            return
        except:
            self.res1.set('—')

    def calc_thick(self, e):
        try:
            L = float(self.e2_L.get())
            a = math.radians(float(self.e2_a.get()))
            t = math.radians(float(self.e2_t.get()))
            hv = L * math.sin(math.atan(math.tan(a) * math.cos(t)))
            ht = L * math.sin(a) * math.sin(t)
            self.res2.set(f'视厚度：{hv:.2f}m  真厚度：{ht:.2f}m')
            return
        except:
            self.res2.set('—')

    def calc_strike(self, e):
        try:
            s = float(self.e3_s.get())
            d1 = (s + 90) % 360
            d2 = (s - 90) % 360
            self.res3s.set(f'倾向：{d1:.1f} / {d2:.1f}')
        except:
            self.res3s.set('—')
        try:
            d = float(self.e3_d.get())
            s1 = (d - 90) % 360
            s2 = (d + 90) % 360
            self.res3d.set(f'走向：{s1:.1f} / {s2:.1f}')
        except:
            self.res3d.set('—')

    def calc_slope(self, e):
        try:
            ang = math.radians(float(self.e5_ang.get()))
            l = float(self.e5_len.get())
            h = l * math.sin(ang)
            p = l * math.cos(ang)
            self.res5.set(f'高差：{h:.2f}m  水平平距：{p:.2f}m')
            return
        except:
            self.res5.set('—')

    def calc_coord(self, e):
        try:
            x0 = float(self.e6x0.get())
            y0 = float(self.e6y0.get())
            az = math.radians(float(self.e6az.get()))
            dis = float(self.e6dis.get())
            dx = dis * math.sin(az)
            dy = dis * math.cos(az)
            self.res6.set(f'终点X={x0 + dx:.2f} Y={y0 + dy:.2f}')
        except:
            self.res6.set('—')
        try:
            x1 = float(self.e6x1.get())
            y1 = float(self.e6y1.get())
            x2 = float(self.e6x2.get())
            y2 = float(self.e6y2.get())
            dx = x2 - x1
            dy = y2 - y1
            dis = math.hypot(dx, dy)
            az = math.degrees(math.atan2(dx, dy))
            az = az if az >= 0 else az + 360
            self.res6rev.set(f'方位：{az:.2f}°  距离：{dis:.2f}m')
        except:
            self.res6rev.set('—')

    # ---- 平面坐标批量换算 ----
    def _coord_batch_mode_changed(self):
        """切换批量口径时重建表头，避免正算的列名套在反算结果上。"""
        head, wid = (_COORD_BATCH_FWD if self.coord_batch_mode.get() == 'forward'
                     else _COORD_BATCH_REV)
        self.coord_batch.set_columns(head, wid, self._coord_batch_parse)

    def _coord_batch_parse(self, line):
        parts = split_batch_fields(line)
        if len(parts) < 4:
            return None
        x_a, y_a, v3, v4 = [float(t) for t in parts[:4]]
        if self.coord_batch_mode.get() == 'forward':
            r = math.radians(v3)
            return ['%.3f' % x_a, '%.3f' % y_a, '%.4f' % v3, '%.3f' % v4,
                    '%.3f' % (x_a + v4 * math.sin(r)),
                    '%.3f' % (y_a + v4 * math.cos(r))]
        x2, y2 = v3, v4
        dx, dy = x2 - x_a, y2 - y_a
        dist = math.hypot(dx, dy)
        az = math.degrees(math.atan2(dx, dy)) % 360.0
        return ['%.3f' % x_a, '%.3f' % y_a, '%.3f' % x2, '%.3f' % y2,
                '%.4f' % az, '%.3f' % dist]

    def calc_az(self, e):
        try:
            b_az = float(self.e7b.get())
            f_az = (b_az + 180) % 360
            self.res7f.set(f'前视方位：{f_az:.2f}°')
        except:
            self.res7f.set('—')
        try:
            f_az = float(self.e7f.get())
            b_az = (f_az - 180) % 360
            self.res7b.set(f'后视方位：{b_az:.2f}°')
        except:
            self.res7b.set('—')

    def calc_hs(self, e):
        try:
            a = math.radians(float(self.e8a.get()))
            h = float(self.e8h.get())
            s = h / math.cos(a)
            self.res8s.set(f'对应斜距：{s:.2f}m')
        except:
            self.res8s.set('—')
        try:
            a = math.radians(float(self.e8a.get()))
            l = float(self.e8l.get())
            h = l * math.cos(a)
            self.res8h.set(f'对应平距：{h:.2f}m')
        except:
            self.res8h.set('—')

    def calc_contour(self, e):
        try:
            h = float(self.e9h.get())
            d = float(self.e9d.get())
            if d == 0:
                ang = 0
            else:
                ang = math.degrees(math.atan(h / d))
            self.res9.set(f'地形坡角：{ang:.2f}°')
            return
        except:
            self.res9.set('—')

    def calc_rate(self, e):
        try:
            a = math.radians(float(self.e10a.get()))
            rate = 1 / math.tan(a)
            self.res10r.set(f'坡率 1:{rate:.2f}')
        except:
            self.res10r.set('—')
        try:
            r = float(self.e10r.get())
            ang = math.degrees(math.atan(1 / r))
            self.res10a.set(f'坡角：{ang:.2f}°')
        except:
            self.res10a.set('—')

    def calc_dthick(self, e):
        try:
            da = math.radians(float(self.e11da.get()))
            l = float(self.e11l.get())
            ra = math.radians(float(self.e11ra.get()))
            ad = math.radians(float(self.e11ad.get()))
            thick = l * math.sin(ra) * math.sin(ad) / math.sin(da)
            self.res11.set(f'分层真厚度：{thick:.2f}m')
            return
        except:
            self.res11.set('—')

    def calc_ddepth(self, e):
        try:
            l = float(self.e12l.get())
            a = math.radians(float(self.e12a.get()))
            depth = l * math.sin(a)
            self.res12.set(f'垂直埋深：{depth:.2f}m')
            return
        except:
            self.res12.set('—')

    def calc_dhor(self, e):
        try:
            l = float(self.e13l.get())
            a = math.radians(float(self.e13a.get()))
            hor = l * math.cos(a)
            self.res13.set(f'水平投影：{hor:.2f}m')
            return
        except:
            self.res13.set('—')

    def calc_dsum(self, e):
        try:
            l1 = float(self.e14l1.get()) if self.e14l1.get() else 0
            a1 = math.radians(float(self.e14a1.get())) if self.e14a1.get() else 0
            l2 = float(self.e14l2.get()) if self.e14l2.get() else 0
            a2 = math.radians(float(self.e14a2.get())) if self.e14a2.get() else 0
            total_len = l1 + l2
            total_depth = l1 * math.sin(a1) + l2 * math.sin(a2)
            self.res14.set(f'总斜长：{total_len:.2f}m 总垂深：{total_depth:.2f}m')
            return
        except:
            self.res14.set('—')

    def calc_dview(self, e):
        try:
            c = float(self.e15c.get())
            d = math.radians(float(self.e15d.get()))
            r = math.radians(float(self.e15r.get()))
            view = c * math.sin(r) / math.sin(d)
            self.res15.set(f'校正后视厚度：{view:.2f}m')
            return
        except:
            self.res15.set('—')

    def calc_dtrue(self, e):
        try:
            v = math.radians(float(self.e16v.get()))
            t = math.radians(float(self.e16t.get()))
            true_ang = math.degrees(math.atan(math.tan(v) / math.sin(t)))
            self.res16.set(f'岩层真倾角：{true_ang:.2f}°')
            return
        except:
            self.res16.set('—')

    def calc_buried(self, e):
        try:
            d = float(self.e17d.get())
            a = math.radians(float(self.e17a.get()))
            buried = d * math.tan(a)
            self.res17.set(f'埋深：{buried:.2f}m')
            return
        except:
            self.res17.set('—')

    def calc_outcrop(self, e):
        try:
            t = float(self.e18t.get())
            a = math.radians(float(self.e18a.get()))
            t2 = math.radians(float(self.e18t2.get()))
            width = t / (math.sin(a) * math.sin(t2))
            self.res18.set(f'露头宽度：{width:.2f}m')
            return
        except:
            self.res18.set('—')

    def calc_hthick(self, e):
        try:
            t = float(self.e19t.get())
            a = math.radians(float(self.e19a.get()))
            h_thick = t / math.tan(a)
            self.res19.set(f'水平厚度：{h_thick:.2f}m')
            return
        except:
            self.res19.set('—')

    def calc_xyz(self, e):
        try:
            az = math.radians(float(self.e20az.get()))
            sl = math.radians(float(self.e20sl.get()))
            l = float(self.e20l.get())
            hor_dis = l * math.cos(sl)
            x = hor_dis * math.sin(az)
            y = hor_dis * math.cos(az)
            z = l * math.sin(sl)
            self.res20.set(f'X={x:.2f} Y={y:.2f} Z={z:.2f}')
            return
        except:
            self.res20.set('—')

    def calc_area(self, e):
        try:
            m2 = float(self.e21.get())
            hectare = m2 / 10000
            mu = m2 / 666.6667
            self.res21.set(f'公顷：{hectare:.4f}  亩：{mu:.4f}')
            return
        except:
            self.res21.set('—')

    def calc_reserve(self, e):
        try:
            l = float(self.e22l.get()) if self.e22l.get() else 0
            w = float(self.e22w.get()) if self.e22w.get() else 0
            t = float(self.e22t.get()) if self.e22t.get() else 0
            d = float(self.e22d.get()) if self.e22d.get() else 0
            volume = l * w * t
            reserve = volume * d
            self.res22.set(f'储量：{reserve:.2f} t')
            return
        except:
            self.res22.set('—')

    def calc_stable(self, e):
        try:
            phi = math.radians(float(self.e23.get()))
            stable_ang = math.degrees(math.atan(math.tan(phi)))
            self.res23.set(f'极限稳定坡角：{stable_ang:.2f}°')
            return
        except:
            self.res23.set('—')

    def calc_grid(self, e):
        try:
            g = float(self.e24g.get())
            s = float(self.e24s.get())
            mm = g * 1000 / s
            self.res24.set(f'图纸格网mm：{mm:.2f}')
            return
        except:
            self.res24.set('—')

    def calc_profile_new(self, e):
        try:
            az = math.radians(float(self.e25_az.get()))
            slope = math.radians(float(self.e25_slope.get()))
            L = float(self.e25_L.get())
            tend = math.radians(float(self.e25_tend.get()))
            dip = math.radians(float(self.e25_dip.get()))
            hor_dis = L * math.cos(slope)
            delta_az = abs(az - tend)
            view_thick = hor_dis * math.sin(dip) * math.sin(delta_az)
            true_thick = view_thick / math.sin(dip)
            vert_thick = hor_dis * math.sin(dip)
            hor_thick = hor_dis * math.cos(dip)
            self.res25.set(f'平距={hor_dis:.2f} 视厚={view_thick:.2f} 真厚={true_thick:.2f} 铅直厚={vert_thick:.2f} 水平厚={hor_thick:.2f}')
            return
        except:
            self.res25.set('—')

    # ---------------- 彩蛋（界面整活 / 猫 / 小游戏 / 实用工具） ----------------
    def _init_easter(self):
        self.strata_on = False
        self._game_open = False
        self._key_seq = []
        self._easter_seq = list('UUDDLRLRA')   # ↑↑↓↓←→←→ + 回车
        self._title_clicks = 0
        self._about_clicks = 0
        self._logo_clicks = 0
        self._cat_phrases = (
            '别戳了，去看剖面。', '地质锤不是这么用的。',
            '今天也要记得带安全帽。', '再点我就罢工了。',
            '……好吧，再来一下。',
        )
        self.root.bind_all('<Key>', self._easter_key, add='+')
        self.root.bind_all('<Control-Alt-g>',
                           lambda e: (self.easter_rock(), 'break')[1], add='+')
        self.root.bind_all('<Control-Alt-p>',
                           lambda e: (self.easter_game(), 'break')[1], add='+')
        self.root.bind_all('<Control-Alt-n>',
                           lambda e: (self.easter_field_card(), 'break')[1], add='+')
        self.root.bind_all('<Control-Alt-r>',
                           lambda e: (self.easter_strata() if self.strata_on
                                      else None, 'break')[1], add='+')

    def _easter_key(self, ev):
        if getattr(self, '_game_open', False):
            return
        kmap = {'Up': 'U', 'Down': 'D', 'Left': 'L', 'Right': 'R', 'Return': 'A'}
        k = ev.keysym
        if k in kmap:
            self._key_seq.append(kmap[k])
            self._key_seq = self._key_seq[-len(self._easter_seq):]
            if self._key_seq == self._easter_seq:
                self._key_seq = []
                self.easter_strata()
        elif k not in ('Shift_L', 'Shift_R', 'Control_L', 'Control_R',
                       'Alt_L', 'Alt_R', 'Caps_Lock', 'Tab',
                       'ISO_Left_Tab', 'KP_Enter'):
            self._key_seq = []

    def _toast(self, msg, ms=4200):
        """底部一闪而过的轻量提示，不打断操作、不弹窗。"""
        try:
            t = tk.Label(self.root, text=msg, bg=ACCENT, fg='#ffffff',
                         font=('微软雅黑', 9), padx=14, pady=6,
                         wraplength=560, justify='left')
            t.place(relx=0.5, rely=0.985, anchor='s')
            self.root.after(ms, t.destroy)
        except tk.TclError:
            pass

    # ---- 1. 界面整活：地层模式配色变装 ----
    def easter_strata(self):
        on = not self.strata_on
        self.strata_on = on
        neutral = (STRATA_PANEL if on else
                   (GLASS_PANEL if self.glass_var.get() else PANEL))
        for w in self._glass_widgets:
            for meth in ('set_strata',):
                try:
                    getattr(w, meth)(on)
                except (tk.TclError, AttributeError):
                    pass
        for w in self._glass_frames:
            try:
                w.configure(bg=neutral)
            except (tk.TclError, AttributeError):
                pass
        for i, (lab, orig) in enumerate(self._group_labels):
            try:
                lab.config(text=STRATA_AGES[i] if on else orig,
                           fg=STRATA_ACCENT if on else ACCENT,
                           bg=neutral)
            except tk.TclError:
                pass
        if not self._readout_labels:
            def collect(w):
                for ch in w.winfo_children():
                    if isinstance(ch, tk.Label):
                        try:
                            if (ch.cget('background') == READOUT_BG and
                                    ch.cget('textvariable')):
                                self._readout_labels.append(ch)
                        except tk.TclError:
                            pass
                    collect(ch)
            for fr in self.tab_frames.values():
                collect(fr)
        for lab in self._readout_labels:
            try:
                lab.config(font=('楷体', 11, 'bold') if on
                           else ('Consolas', 11, 'bold'),
                           fg=STRATA_ACCENT if on else RESULT)
            except tk.TclError:
                pass
        self.subtitle_label.config(
            text='第四纪 · 地层模式' if on else '地质测量计算工具')
        if self.active_idx >= 0:
            self.set_active(self.active_idx)
        self._toast('◆ 地层模式已开启 · 再敲一次或按 Ctrl+Alt+R 还原'
                    if on else '◆ 已回到常规视图')

    # ---- 2. 让猫动起来 ----
    def _cat_photo(self):
        if getattr(self, '_cat_img', 'miss') != 'miss':
            return self._cat_img
        base = (getattr(sys, '_MEIPASS', '')
                if getattr(sys, 'frozen', False) else
                os.path.dirname(os.path.abspath(__file__)))
        try:
            self._cat_img = tk.PhotoImage(file=os.path.join(base, 'cat_64.png'))
        except (tk.TclError, OSError):
            self._cat_img = None
        return self._cat_img

    def _about_cat(self, parent):
        box = tk.Canvas(parent, width=104, height=74, bg='#ffffff',
                        highlightthickness=0, cursor='hand2')
        box.pack(pady=(12, 0))
        img = self._cat_photo()
        cat = (box.create_image(40, 38, image=img) if img else
               box.create_oval(20, 18, 60, 58, fill=ACCENT, outline=''))
        handle = box.create_line(72, 62, 84, 34, fill=MUTED, width=3,
                                 capstyle='round')
        head = box.create_rectangle(76, 28, 94, 36, fill=ACCENT, outline='')
        say = tk.Label(parent, text='（点我一下试试）', bg='#ffffff',
                       fg=MUTED, font=('微软雅黑', 9))
        say.pack()
        box.bind('<Button-1>',
                 lambda e: self._cat_animate(box, cat, handle, head, say))
        self._about_cat_refs = (box, cat, handle, head, say)
        return box

    def _cat_animate(self, canvas, cat, handle, head, say):
        phrases = self._cat_phrases

        def step(k=0):
            if k >= 10:
                try:
                    canvas.coords(cat, 40, 38)
                    canvas.coords(handle, 72, 62, 84, 34)
                    canvas.coords(head, 76, 28, 94, 36)
                    say.config(text='（点我一下试试）')
                except tk.TclError:
                    pass
                return
            a = math.radians(-70 + (k % 4) * 24)
            x1 = 72 + 18 * math.cos(a)
            y1 = 62 + 18 * math.sin(a)
            try:
                canvas.coords(handle, 72, 62, x1, y1)
                canvas.coords(head, x1 - 9, y1 - 4, x1 + 9, y1 + 4)
                canvas.coords(cat, 40, 38 + (3 if k % 2 else 0))
                say.config(text=phrases[min(k, len(phrases) - 1)])
            except tk.TclError:
                return
            self.root.after(110, lambda: step(k + 1))
        step()

    def _about_click(self, _e):
        self._about_clicks += 1
        if self._about_clicks >= 5:
            self._about_clicks = 0
            refs = getattr(self, '_about_cat_refs', None)
            if refs:
                try:
                    self._cat_animate(*refs)
                    refs[4].config(text='隐藏表情已解锁：猫在瞪你。')
                except tk.TclError:
                    pass

    # ---- 3. 罗盘 / 标题的小互动 ----
    def _title_click(self, _e):
        self._title_clicks += 1
        if self._title_clicks >= 7:
            self._title_clicks = 0
            line = random.choice(EASTER_TITLE_LINES)
            self.subtitle_label.config(text=line)
            self.root.after(4500, lambda: self.subtitle_label.config(
                text='第四纪 · 地层模式' if self.strata_on
                else '地质测量计算工具'))
            self._toast('◆ ' + line)

    def _logo_click(self, _e):
        self._logo_clicks += 1
        if self._logo_clicks >= 5:
            self._logo_clicks = 0
            self._spin_compass()

    def _spin_compass(self):
        c = self.logo_canvas
        cx, cy, r = self.logo_geom
        old = self.subtitle_label.cget('text')
        self.subtitle_label.config(text='罗盘：别转了，我晕')

        def frame(k=0):
            if k > 40:
                self.subtitle_label.config(text=old)
                return
            a = math.radians(-45 + k * 18)
            try:
                c.coords(self.logo_needle_a, cx, cy,
                         cx + r * 0.7 * math.cos(a),
                         cy + r * 0.7 * math.sin(a))
                c.coords(self.logo_needle_b, cx, cy,
                         cx - r * 0.5 * math.cos(a),
                         cy - r * 0.5 * math.sin(a))
            except tk.TclError:
                return
            self.root.after(24, lambda: frame(k + 1))
        frame()

    # ---- 4. 实用向：今日岩层 / 野外记录卡 ----
    def easter_rock(self):
        name, desc = random.choice(EASTER_ROCKS)
        self._toast('◆ 今日岩层 · %s：%s' % (name, desc), ms=5200)

    def _collect_current(self):
        idx = self.active_idx
        name = (self.func_names[idx]
                if 0 <= idx < len(self.func_names) else '—')
        fr = self.tab_frames.get(self.func_bases[idx]) if 0 <= idx < len(self.func_bases) else None
        fields, results = [], []

        def walk(w):
            for ch in w.winfo_children():
                if isinstance(ch, (tk.Entry, ttk.Entry)):
                    label = ''
                    for sib in w.winfo_children():
                        if isinstance(sib, tk.Label):
                            try:
                                label = sib.cget('text')
                            except tk.TclError:
                                pass
                            break
                    try:
                        val = ch.get()
                    except tk.TclError:
                        val = ''
                    if val:
                        fields.append((label or '输入', val))
                if isinstance(ch, tk.Label):
                    try:
                        if (ch.cget('background') == READOUT_BG and
                                ch.cget('textvariable')):
                            v = self.root.getvar(ch.cget('textvariable'))
                            if v and v != '—':
                                results.append(str(v))
                    except (tk.TclError, AttributeError):
                        pass
                walk(ch)

        if fr is not None:
            walk(fr)
        return name, fields, results

    def easter_field_card(self):
        name, fields, results = self._collect_current()
        lines = [
            '野 外 记 录 卡',
            '=' * 34,
            '时间：' + datetime.now().strftime('%Y-%m-%d %H:%M'),
            '功能：' + name,
            '-' * 34,
            '【输入】',
        ]
        lines += ['  %s：%s' % (k, v) for k, v in fields] or ['  （暂无输入）']
        lines += ['', '【计算结果】']
        lines += ['  ' + v for v in results] or ['  （暂无结果）']
        lines += [
            '', '-' * 34,
            '点位：____________   天气：____________',
            '岩性：____________   产状：____________',
            '记录人：____________',
            '', '—— ' + random.choice(EASTER_QUOTES),
        ]
        content = '\n'.join(lines)

        w = tk.Toplevel(self.root)
        w.title('野外记录卡')
        w.geometry('520x460')
        w.configure(bg=BG)
        self._apply_icon(w)
        w.transient(self.root)
        txt = scrolledtext.ScrolledText(w, font=('Consolas', 10),
                                        bg=CARD, fg=INK, relief='flat',
                                        borderwidth=0, padx=12, pady=10)
        txt.pack(fill='both', expand=True, padx=12, pady=12)
        txt.insert('1.0', content)
        bar = tk.Frame(w, bg=BG)
        bar.pack(fill='x', padx=12, pady=(0, 12))
        ttk.Button(bar, text='复制到剪贴板', style='Ghost.TButton',
                   command=lambda: (self.root.clipboard_clear(),
                                    self.root.clipboard_append(txt.get('1.0', 'end')),
                                    self._toast('◆ 记录卡已复制'))).pack(
            side='left', padx=(0, 8))

        def save():
            path = filedialog.asksaveasfilename(
                parent=w, title='保存记录卡', defaultextension='.txt',
                initialfile='野外记录卡.txt',
                filetypes=[('文本文件', '*.txt'), ('所有文件', '*.*')])
            if path:
                with open(path, 'w', encoding='utf-8') as fh:
                    fh.write(txt.get('1.0', 'end'))
                self._toast('◆ 已保存：' + path)
        ttk.Button(bar, text='另存为 txt', style='Ghost.TButton',
                   command=save).pack(side='left')
        ttk.Button(bar, text='关闭', style='Primary.TButton',
                   command=w.destroy).pack(side='right')

    # ---- 5. 小游戏：罗盘校准 ----
    def easter_game(self):
        if getattr(self, '_game_open', False):
            return
        self._game_open = True
        w = tk.Toplevel(self.root)
        w.title('彩蛋 · 罗盘校准')
        w.geometry('400x460')
        w.resizable(False, False)
        w.configure(bg=BG)
        self._apply_icon(w)
        w.transient(self.root)
        w.protocol('WM_DELETE_WINDOW',
                   lambda: (setattr(self, '_game_open', False), w.destroy()))

        tk.Label(w, text='把红指针转到目标方位，空格锁定', bg=BG, fg=INK,
                 font=('微软雅黑', 11, 'bold')).pack(pady=(12, 4))
        info = tk.Label(w, text='', bg=BG, fg=ACCENT,
                        font=('微软雅黑', 10))
        info.pack()
        c = tk.Canvas(w, width=260, height=260, bg=CARD,
                      highlightthickness=0)
        c.pack(pady=8)
        sc = tk.IntVar(value=0)
        tk.Scale(w, from_=0, to=359, orient='horizontal', length=300,
                 variable=sc, bg=BG, fg=INK, highlightthickness=0,
                 troughcolor=READOUT_BG).pack()
        msg = tk.Label(w, text='', bg=BG, fg=RESULT,
                       font=('微软雅黑', 10, 'bold'), wraplength=360)
        msg.pack(pady=6)
        btns = tk.Frame(w, bg=BG)
        btns.pack(pady=(0, 12))
        state = {'target': 0, 'round': 1, 'score': 0, 'over': False}
        R = 100

        def draw():
            c.delete('all')
            cx = cy = 130
            c.create_oval(cx - R, cy - R, cx + R, cy + R,
                          outline=ACCENT, width=2)
            for deg in range(0, 360, 15):
                a = math.radians(deg - 90)
                r0 = R - (10 if deg % 45 == 0 else 5)
                c.create_line(cx + r0 * math.cos(a), cy + r0 * math.sin(a),
                              cx + R * math.cos(a), cy + R * math.sin(a),
                              fill='#9fb3c7', width=1)
            for deg, ch in ((0, 'N'), (90, 'E'), (180, 'S'), (270, 'W')):
                a = math.radians(deg - 90)
                c.create_text(cx + (R - 20) * math.cos(a),
                              cy + (R - 20) * math.sin(a),
                              text=ch, fill=INK,
                              font=('微软雅黑', 9, 'bold'))
            a = math.radians(state['target'] - 90)
            c.create_polygon(
                [(cx + (R - 14) * math.cos(a), cy + (R - 14) * math.sin(a)),
                 (cx + (R - 2) * math.cos(a - 0.09),
                  cy + (R - 2) * math.sin(a - 0.09)),
                 (cx + (R - 2) * math.cos(a + 0.09),
                  cy + (R - 2) * math.sin(a + 0.09))],
                fill=ACCENT, outline='')
            a2 = math.radians(sc.get() - 90)
            c.create_line(cx, cy, cx + (R - 20) * math.cos(a2),
                          cy + (R - 20) * math.sin(a2),
                          fill=RESULT, width=3, arrow=tk.LAST)
            c.create_oval(cx - 4, cy - 4, cx + 4, cy + 4,
                          fill=ACCENT, outline='')
            info.config(text='第 %d / 5 轮   目标 %03d°   当前 %03d°   总分 %d' %
                             (state['round'], state['target'], sc.get(),
                              state['score']))

        def new_round():
            state['target'] = random.randint(0, 359)
            sc.set(random.randint(0, 359))
            state['over'] = False
            msg.config(text='← → 微调，也可以直接拖滑杆')
            draw()

        def lock(_e=None):
            if state['over']:
                return
            err = abs(((state['target'] - sc.get() + 180) % 360) - 180)
            gain = max(0, 100 - err * 2)
            state['score'] += gain
            if state['round'] >= 5:
                state['over'] = True
                total = state['score']
                if total >= 450:
                    grade = '老测绘员，罗盘认得你'
                elif total >= 350:
                    grade = '手感不错，能独立跑剖面了'
                elif total >= 200:
                    grade = '还得多跑几条剖面'
                else:
                    grade = '罗盘表示想回仪器箱'
                msg.config(text='本轮误差 %d°，+%d 分\n总分 %d · %s' %
                                (err, gain, total, grade))
                info.config(text='结束 · 总分 %d' % total)
            else:
                state['round'] += 1
                msg.config(text='误差 %d°，+%d 分 · 下一轮' % (err, gain))
                self.root.after(700, new_round)

        def nudge(d):
            if not state['over']:
                sc.set((sc.get() + d) % 360)
                draw()

        sc.trace_add('write', lambda *a: draw() if not state['over'] else None)
        w.bind('<Left>', lambda e: nudge(-1))
        w.bind('<Right>', lambda e: nudge(1))
        w.bind('<Up>', lambda e: nudge(-10))
        w.bind('<Down>', lambda e: nudge(10))
        w.bind('<space>', lock)
        w.bind('<Return>', lock)
        ttk.Button(btns, text='锁定方位（空格）', style='Primary.TButton',
                   command=lock).pack(side='left', padx=(0, 8))
        ttk.Button(btns, text='再来一局', style='Ghost.TButton',
                   command=lambda: (state.update(round=1, score=0),
                                    new_round())).pack(side='left')
        w.focus_set()
        new_round()

    # ---- 6. 功能 4 的著名点位注记（不弹窗，只写状态栏） ----
    def _coord_easter(self):
        for r in getattr(self, 'dms_results', []):
            lng, lat = r.get('lng_dec'), r.get('lat_dec')
            if lng is None or lat is None:
                continue
            for slng, slat, label in EASTER_SPOTS:
                if abs(lng - slng) < 0.02 and abs(lat - slat) < 0.02:
                    self.dms_status.set(self.dms_status.get() + '  ◆ ' + label)
                    self._toast('◆ 点位注记：' + label)
                    return

    # ---------------- 帮助 / 关于 ----------------
    def show_doc(self):
        doc_win = tk.Toplevel(self.root)
        doc_win.title(f'帮助文档 · {APP_NAME} {APP_VER}')
        doc_win.geometry('720x580')
        doc_win.minsize(560, 420)
        doc_win.resizable(True, True)
        doc_win.configure(bg=BG)
        self._apply_icon(doc_win)
        doc_win.transient(self.root)

        top = tk.Frame(doc_win, bg=BG)
        top.pack(fill='x', padx=14, pady=(14, 0))
        tk.Label(top, text='使用帮助', bg=BG, fg=INK,
                 font=('微软雅黑', 14, 'bold')).pack(side='left')
        tk.Label(top, text=APP_VER, bg=ACCENT, fg='white',
                 font=('微软雅黑', 9, 'bold'), padx=8, pady=2).pack(side='right')

        box = tk.Frame(doc_win, bg=CARD, highlightbackground=LINE,
                       highlightthickness=1)
        box.pack(fill='both', expand=True, padx=14, pady=10)
        text = scrolledtext.ScrolledText(box, wrap='char', width=52, height=14,
                                         bg=CARD, fg=INK,
                                         font=('微软雅黑', 10), relief='flat',
                                         borderwidth=0, padx=14, pady=10)
        text.pack(fill='both', expand=True, padx=1, pady=1)
        text.tag_config('h1', font=('微软雅黑', 12, 'bold'), foreground=ACCENT,
                        spacing1=12, spacing3=5)
        text.tag_config('h2', font=('微软雅黑', 10, 'bold'), foreground=ACCENT_DK,
                        spacing1=8, spacing3=3)
        text.tag_config('p', font=('微软雅黑', 10), foreground=INK,
                        lmargin1=8, lmargin2=8, spacing2=2)
        text.tag_config('li', font=('微软雅黑', 10), foreground=INK,
                        lmargin1=20, lmargin2=32, spacing2=2)
        text.tag_config('mono', font=('Consolas', 10), foreground='#2c5d8a',
                        lmargin1=28, spacing2=3)
        text.tag_config('note', font=('微软雅黑', 9), foreground=MUTED,
                        lmargin1=8, spacing2=2)

        def add(line, tag='p'):
            text.insert(tk.END, line + '\n', tag)

        add(f'{APP_NAME} {APP_VER} · {APP_SUBTITLE}', 'h1')
        add(f'版本 {APP_VER}', 'note')
        add('本软件面向地质、测绘野外与室内计算，内置 29 项常用计算，'
            '所有结果随输入即时刷新。', 'p')

        add('一、快速开始', 'h1')
        add('1. 在左侧「功能目录」点选需要的计算项，也可从顶部「功能」菜单按分组跳转。', 'li')
        add('2. 填入已知量；带选项按钮的页面（如「视倾角↔真倾角」「常用单位换算」）'
            '先切换方向再输入。', 'li')
        add('3. 结果即时显示。点右下角「清空输入」或按 Ctrl+W 可一键重来。', 'li')

        add('二、界面说明', 'h1')
        add('· 顶部菜单栏：文件（导入 / 导出 / 退出）、功能（按 6 组分类快速跳转）、'
            '编辑（清空 / 复制 / 历史）、视图、帮助（帮助文档 / 更新说明 / 检查更新 / 关于）。', 'li')
        add('· 左侧功能目录：磨砂玻璃圆角按钮，蓝色高亮为当前功能；'
            '列表可滚动，鼠标停在任意条目上都能用滚轮滚动。', 'li')
        add('· 内容卡片：白底卡片内为输入项与结果读数，结果为「—」表示'
            '尚未计算或输入不完整。', 'li')
        add('· 视图 → 磨砂玻璃质感：可随时在玻璃外观与经典平面外观之间切换。', 'li')
        add('界面效果由软件自绘，不依赖特定 Windows 版本，Win10 / Win11 与高 DPI 缩放均可用。', 'note')

        add('三、功能一览（共 29 项，按类型分组）', 'h1')
        add('产状与厚度（1–6）', 'h2')
        add('1 视倾角↔真倾角   2 走向倾向互换   3 地层真/视厚度', 'li')
        add('4 岩层水平厚度   5 露头出露宽度   6 实测剖面厚度', 'li')
        add('钻孔计算（7–13）', 'h2')
        add('7 分层真厚度   8 视厚度校正   9 真倾角校正   10 垂直埋深', 'li')
        add('11 水平投影   12 多段累计   13 地表岩层埋深', 'li')
        add('坐标与方位（14–20）', 'h2')
        add('14 经纬度度分秒批量转换   15 经纬度→直角坐标   16 平面坐标正反算', 'li')
        add('17 批量坐标正反算   18 方位角前后视   19 坐标格网长度', 'li')
        add('20 导线三维 XYZ 增量', 'li')
        add('地形与坡度（21–24）', 'h2')
        add('21 坡度高差平距   22 平距斜距换算   23 等高线坡角   24 坡度坡率互转', 'li')
        add('储量与工程（25–26）', 'h2')
        add('25 块段矿石储量   26 岩土稳定坡角', 'li')
        add('单位与工具（27–29）', 'h2')
        add('27 面积单位换算   28 常用单位换算   29 计算历史', 'li')
        add('15 经纬度→直角坐标支持 CGCS2000 / WGS84 / 西安80 / 北京54 四种椭球，'
            '同时输出地心直角坐标 XYZ 与高斯投影平面坐标；'
            '28 常用单位换算提供角度↔弧度、米↔英尺、公里↔海里、公顷↔亩、'
            '平方米↔亩、十进制度↔度分秒等 11 组即时换算。', 'note')

        add('四、经纬度文本格式（功能 14）', 'h1')
        add('数值为「先经后纬、一行一组」时直接粘贴即可。分隔符支持半角 / 全角逗号、'
            '顿号、分号、斜杠、空格与制表符：', 'p')
        add('116.391, 39.907', 'mono')
        add('116.391，39.907              ← 全角逗号', 'mono')
        add('经度：116.391，纬度：39.907   ← 中文标签', 'mono')
        add('E116°23′45″, N39°54′23″      ← 度分秒 + 方位', 'mono')
        add('东经116度23分45秒 北纬39度54分23秒', 'mono')
        add('120:15:30, 30:20:10         ← 冒号格式', 'mono')
        add('1, 116.391, 39.907          ← 带序号会自动跳过', 'mono')
        add('· 方位支持中文（东经 / 西经 / 北纬 / 南纬）与 E/W/N/S，西经、南纬自动取负。', 'li')
        add('· 括号注释、标签文字、序号前缀会自动忽略；若写成「先纬后经」，'
            '软件按数值范围自动纠正。', 'li')
        add('· 两个数值默认按先经后纬处理；只填一个数值时，|值|>90 判为经度，否则判为纬度。', 'li')

        add('五、导入、导出与计算历史', 'h1')
        add('支持表格化结果的页面：14 经纬度度分秒批量转换、15 经纬度→直角坐标、'
            '16 平面坐标正反算、17 批量坐标正反算、29 计算历史。'
            '这几页都能「导入文本 → 开始换算 → 复制 / 导出」。', 'p')
        add('· 导入：读取 .txt / .csv，UTF-8 与 GBK 编码都能识别。', 'li')
        add('· 导出：可存为 Excel(.xlsx) 或 CSV(.csv)，带表头。', 'li')
        add('· 文件菜单的「导入坐标文本 / 导出当前结果」会作用于当前所在页面。', 'li')
        add('计算历史自动记录本次会话的计算结果（按功能去重），'
            '可随时查看、复制、导出或清空；退出软件后不保留。', 'p')

        add('六、快捷键', 'h1')
        add('F1              帮助文档', 'mono')
        add('Ctrl+O          导入坐标文本', 'mono')
        add('Ctrl+S          导出当前结果', 'mono')
        add('Ctrl+W          清空当前输入', 'mono')
        add('Ctrl+Shift+C    复制当前结果', 'mono')
        add('Ctrl+D          经纬度度分秒转换', 'mono')
        add('Ctrl+U          常用单位换算', 'mono')
        add('Ctrl+H          计算历史', 'mono')
        add('Ctrl+Q          退出', 'mono')

        add('七、常见问题', 'h1')
        add('Q：输入留空或填了文字，结果是什么？', 'li')
        add('A：显示「—」，不会给出 0 值结果，避免误读为有效计算值。', 'p')
        add('Q：任务栏图标还是 Python 图标？', 'li')
        add('A：请运行打包后的 exe。直接用 python 运行 .py 时，'
            '任务栏按钮归属于解释器 python.exe，其图标无法更改。', 'p')
        add('Q：切换选项按钮后结果没变？', 'li')
        add('A：带选项的页面（视倾角互算、单位换算）切换方向后会立即重算；'
            '其余页面请重新输入或点「计算」。', 'p')
        add('Q：结果在哪里复制？', 'li')
        add('A：编辑 → 复制当前结果（Ctrl+Shift+C），或使用页面自带的「复制结果」按钮。', 'p')

        add('八、输入约定与软件信息', 'h1')
        add('· 所有角度输入支持十进制（如 30.5°）；长度单位默认为米(m)。', 'li')
        add('· 度分秒支持 116°23′45″ / 116:23:45 / 116.391 等写法。', 'li')
        add(f'软件名称：{APP_NAME} {APP_VER}（{APP_SUBTITLE}）', 'p')
        add(f'版本号：{APP_VER}        作者：{APP_AUTHOR}', 'p')
        add(f'联系方式：{APP_CONTACT}', 'p')
        add('在资源管理器中右键 exe → 属性 → 详细信息，'
            '同样可以查看版本号、产品名称与作者信息。', 'note')

        add('九、版本更新通知', 'h1')
        add('软件启动后会在后台向更新源查询一次版本号，不影响正常使用；'
            '发现新版本时会弹出提示，写明更新内容与下载地址。', 'p')
        add('更新源是一份放在公共网页上的 version.json，'
            '用 GitHub Pages 或 Gitee Pages 托管即可，免费、不需要服务器。', 'li')
        add('· 自动检查最多每 6 小时一次；「帮助 → 检查更新」可随时手动检查。', 'li')
        add('· 连不上更新源时自动跳过，离线使用不受任何影响。', 'li')
        add('· 关掉提示后 6 小时内不再重复提醒，之后仍会提示，直到升级到新版为止。', 'li')
        add('· 主地址读不通会自动尝试备用地址（仓库名大小写变体、仓库直链），全部失败则静默跳过。', 'li')
        add('· 提示窗口的「前往下载」会按 Release 实际资产地址自动校正；'
            '另有「发布页」按钮，可直接打开全部历史版本。', 'li')
        add('· 更新源也支持本地文件与局域网共享盘路径，方便单位内部使用。', 'li')

        add('十、隐藏的小趣味', 'h1')
        add('软件里藏了几处与地质有关的小彩蛋：界面会变装、罗盘会转、'
            '关于窗口里的猫会挥锤子，还有一个能玩的小游戏和一张'
            '可直接填写的野外记录卡。', 'p')
        add('提示：老测绘员总爱在键盘上敲一串方向键；'
            'Ctrl+Alt 加上 G / P / N 也各有点名堂。'
            '若输入法把组合键吃掉了，「视图」菜单里同样能打开它们。', 'note')

        text.config(state=tk.DISABLED)
        text.yview_moveto(0)      # 确保打开时首屏显示标题，而不是停在中间
        text.focus_set()

    # ---------------- 更新说明 ----------------
    def _state_file(self):
        """记录「已看过的版本」，放在 %APPDATA%\\地算通 下，升级后首次启动才会再弹。"""
        base = os.environ.get('APPDATA') or os.path.expanduser('~')
        folder = os.path.join(base, APP_NAME)
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError:
            folder = os.path.expanduser('~')
        return os.path.join(folder, 'state.json')

    def _load_state(self):
        try:
            with open(self._state_file(), 'r', encoding='utf-8') as fh:
                data = json.load(fh)
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    def _save_state(self, data):
        try:
            with open(self._state_file(), 'w', encoding='utf-8') as fh:
                json.dump(data, fh, ensure_ascii=False, indent=2)
        except OSError:
            pass

    def _maybe_show_changelog(self):
        """版本变化后的第一次启动自动弹出更新说明；同一版本只弹一次。"""
        state = self._load_state()
        if state.get('seen_version') == APP_VER:
            return
        state['seen_version'] = APP_VER
        self._save_state(state)
        self.show_changelog(auto=True)

    def show_changelog(self, auto=False):
        win = tk.Toplevel(self.root)
        win.title(f'更新说明 · {APP_NAME} {APP_VER}')
        win.geometry('620x520')
        win.minsize(480, 360)
        win.configure(bg=BG)
        self._apply_icon(win)
        win.transient(self.root)

        top = tk.Frame(win, bg=BG)
        top.pack(fill='x', padx=14, pady=(14, 0))
        tk.Label(top, text='更新说明', bg=BG, fg=INK,
                 font=('微软雅黑', 14, 'bold')).pack(side='left')
        tk.Label(top, text=APP_VER, bg=ACCENT, fg='white',
                 font=('微软雅黑', 9, 'bold'), padx=8, pady=2).pack(side='right')
        tk.Label(win, text=('本次更新内容如下，祝你野外顺利。' if auto
                            else '按版本倒序排列，最新版本在最上方。'),
                 bg=BG, fg=MUTED, font=('微软雅黑', 9)).pack(
                     anchor='w', padx=16, pady=(4, 0))

        box = tk.Frame(win, bg=CARD, highlightbackground=LINE,
                       highlightthickness=1)
        box.pack(fill='both', expand=True, padx=14, pady=10)
        text = scrolledtext.ScrolledText(box, wrap='char', width=52, height=14,
                                         bg=CARD, fg=INK,
                                         font=('微软雅黑', 10), relief='flat',
                                         borderwidth=0, padx=14, pady=10)
        text.pack(fill='both', expand=True, padx=1, pady=1)
        text.tag_config('v', font=('微软雅黑', 12, 'bold'), foreground=ACCENT,
                        spacing1=14, spacing3=4)
        text.tag_config('cur', font=('微软雅黑', 8, 'bold'), foreground='#ffffff',
                        background=ACCENT, spacing3=4, lmargin1=8, lmargin2=8)
        text.tag_config('date', font=('微软雅黑', 9), foreground=MUTED, spacing3=3)
        text.tag_config('li', font=('微软雅黑', 10), foreground=INK,
                        lmargin1=20, lmargin2=34, spacing2=3)

        for ver, date, items in CHANGELOG:
            text.insert(tk.END, ver, 'v')
            if ver == APP_VER:
                text.insert(tk.END, '   ')
                text.insert(tk.END, ' 当前版本 ', 'cur')
            text.insert(tk.END, '\n')
            if date:
                text.insert(tk.END, date + '\n', 'date')
            for it in items:
                text.insert(tk.END, '· ' + it + '\n', 'li')
        text.config(state=tk.DISABLED)
        text.yview_moveto(0)
        text.focus_set()

        bar = tk.Frame(win, bg=BG)
        bar.pack(fill='x', padx=14, pady=(0, 12))
        ttk.Button(bar, text='知道了', style='Primary.TButton',
                   command=win.destroy).pack(side='right')

    # ---------------- 版本更新通知 ----------------
    def _ver_tuple(self, text):
        """把 V2.10 这类版本号转成可比较的数字元组，避免字符串比较把 2.10 判成小于 2.9。"""
        nums = re.findall(r'\d+', str(text or ''))
        return tuple(int(x) for x in nums[:3]) if nums else ()

    def _feed_sources(self):
        """主地址 + 备用地址；备用项可写成单个字符串或列表。"""
        fb = VERSION_URL_FALLBACK or []
        if isinstance(fb, str):
            fb = [fb]
        urls = [VERSION_URL] + list(fb)
        return [u.strip() for u in urls if isinstance(u, str) and u.strip()]

    def _feed_source(self):
        """主地址，用于手动检查失败时的提示文案。"""
        urls = self._feed_sources()
        return urls[0] if urls else ''

    @staticmethod
    def _cache_bust(url):
        """Pages 走 CDN，会缓存十分钟左右；加一个时间戳参数，保证读到刚改过的 version.json。"""
        if not url.lower().startswith(('http://', 'https://')):
            return url
        sep = '&' if '?' in url else '?'
        return '%s%st=%d' % (url, sep, int(time.time()))

    def _fetch_one(self, src):
        """读单个地址；支持 http/https、本地文件与共享盘路径，失败一律返回 None。"""
        try:
            if src.lower().startswith(('http://', 'https://')):
                req = urllib.request.Request(
                    self._cache_bust(src),
                    headers={'User-Agent': 'Dicalc/%s' % APP_VER,
                             'Cache-Control': 'no-cache'})
                with urllib.request.urlopen(req, timeout=6) as resp:
                    raw = resp.read().decode('utf-8-sig', 'replace')
            else:
                with open(src, 'r', encoding='utf-8-sig') as fh:
                    raw = fh.read()
            data = json.loads(raw)
            return data if isinstance(data, dict) else None
        except (OSError, ValueError):
            return None

    def _read_feed(self):
        """按主 → 备顺序尝试，任一地址读通即返回；全部失败返回 None（静默跳过）。"""
        for src in self._feed_sources():
            info = self._fetch_one(src)
            if info is not None:
                return info
        return None

    def _resolve_download(self, url):
        """GitHub Release 直链里的中文文件名常被改动，导致 version.json 写的地址 404。

        这里用公开的 GitHub API 反查该 tag 的真实资产地址（无需 token）；
        查不到就原样返回，再由弹窗里的「发布页」按钮兜底。全程在后台线程执行。
        """
        url = str(url or '').strip()
        m = re.match(r'https://github\.com/([^/]+)/([^/]+)/releases/download/([^/]+)/',
                     url)
        if not m:
            return url
        api = ('https://api.github.com/repos/%s/%s/releases/tags/%s'
               % (m.group(1), m.group(2), m.group(3)))
        try:
            req = urllib.request.Request(
                api, headers={'User-Agent': 'Dicalc/%s' % APP_VER,
                              'Accept': 'application/vnd.github+json'})
            with urllib.request.urlopen(req, timeout=6) as resp:
                data = json.loads(resp.read().decode('utf-8', 'replace'))
            assets = data.get('assets') if isinstance(data, dict) else None
            if assets:
                real = str(assets[0].get('browser_download_url') or '').strip()
                if real:
                    return real
        except (OSError, ValueError):
            pass
        return url

    def check_update(self, manual=False):
        """检查更新。网络读取放在后台线程，离线或超时都不会卡住界面。"""
        if not self._feed_source():
            if manual:
                self._dialog(
                    'showinfo', '检查更新',
                    '尚未配置更新源，自动检查已关闭。\n\n'
                    '做法：把 version.json 放到 GitHub / Gitee 仓库根目录，'
                    '在仓库 Settings → Pages 里开启页面服务，\n'
                    '拿到 https://<用户名>.github.io/<仓库名>/version.json 这样的公开链接，'
                    '填进源码顶部的 VERSION_URL 后重新打包即可。\n'
                    '以后发新版只需修改这一个 version.json。')
            return
        state = self._load_state()
        if not manual:
            try:
                if time.time() - float(state.get('last_update', 0)) < APP_UPDATE_COOLDOWN:
                    return
            except (TypeError, ValueError):
                pass
        threading.Thread(target=self._update_worker, args=(manual,),
                         daemon=True).start()

    def _update_worker(self, manual):
        info = self._read_feed()
        # 只有确实发现新版本时，才顺带校正下载直链，避免每次启动都多打一次 API
        if isinstance(info, dict):
            remote = str(info.get('version', '')).strip()
            if remote and self._ver_tuple(remote) > self._ver_tuple(APP_VER):
                info = dict(info)
                info['url'] = self._resolve_download(info.get('url', ''))
        try:
            self.root.after(0, lambda: self._update_done(info, manual))
        except tk.TclError:
            pass

    def _update_done(self, info, manual):
        state = self._load_state()
        state['last_update'] = time.time()
        self._save_state(state)
        if info is None:
            if manual:
                self._dialog('showwarning', '检查更新',
                             '无法读取更新源，请检查网络连接，或确认以下地址是否可访问：\n\n'
                             + '\n'.join(self._feed_sources()))
            return
        remote = str(info.get('version', '')).strip()
        if not remote:
            if manual:
                self._dialog('showwarning', '检查更新', '更新源内容缺少 version 字段。')
            return
        if self._ver_tuple(remote) <= self._ver_tuple(APP_VER):
            if manual:
                self._dialog('showinfo', '检查更新',
                             '当前已是最新版本 %s。' % APP_VER)
            return
        self.show_update_notice(info)

    def show_update_notice(self, info):
        remote = str(info.get('version', '')).strip()
        url = str(info.get('url', '')).strip()
        notes = info.get('notes') or []
        if isinstance(notes, str):
            notes = [notes]

        win = tk.Toplevel(self.root)
        win.title('发现新版本 · %s %s' % (APP_NAME, remote))
        win.geometry('620x520')
        win.minsize(480, 380)
        win.configure(bg=BG)
        self._apply_icon(win)
        win.transient(self.root)

        top = tk.Frame(win, bg=BG)
        top.pack(fill='x', padx=14, pady=(14, 0))
        tk.Label(top, text='发现新版本', bg=BG, fg=INK,
                 font=('微软雅黑', 14, 'bold')).pack(side='left')
        tk.Label(top, text=remote, bg=ACCENT, fg='white',
                 font=('微软雅黑', 9, 'bold'), padx=8, pady=2).pack(side='right')
        tk.Label(win, text='当前版本 %s，可更新到 %s。' % (APP_VER, remote),
                 bg=BG, fg=MUTED, font=('微软雅黑', 9)).pack(
                     anchor='w', padx=16, pady=(4, 0))

        box = tk.Frame(win, bg=CARD, highlightbackground=LINE,
                       highlightthickness=1)
        box.pack(fill='both', expand=True, padx=14, pady=10)
        text = scrolledtext.ScrolledText(box, wrap='char', width=52, height=14,
                                         bg=CARD, fg=INK,
                                         font=('微软雅黑', 10), relief='flat',
                                         borderwidth=0, padx=14, pady=10)
        text.pack(fill='both', expand=True, padx=1, pady=1)
        text.tag_config('h', font=('微软雅黑', 11, 'bold'), foreground=ACCENT,
                        spacing1=4, spacing3=6)
        text.tag_config('li', font=('微软雅黑', 10), foreground=INK,
                        lmargin1=18, lmargin2=32, spacing2=3)
        text.tag_config('n', font=('微软雅黑', 9), foreground=MUTED,
                        lmargin1=8, spacing1=8)

        if info.get('date'):
            text.insert(tk.END, '%s 发布\n' % info['date'], 'n')
        text.insert(tk.END, '本次更新内容\n', 'h')
        for item in notes:
            text.insert(tk.END, '· %s\n' % item, 'li')
        if not notes:
            text.insert(tk.END, '（更新源未提供详细说明）\n', 'n')
        if url:
            text.insert(tk.END, '\n下载地址\n%s\n' % url, 'n')
        text.config(state=tk.DISABLED)
        text.yview_moveto(0)

        bar = tk.Frame(win, bg=BG)
        bar.pack(fill='x', padx=14, pady=(0, 12))
        ttk.Button(bar, text='关闭', style='Ghost.TButton',
                   command=win.destroy).pack(side='right', padx=(6, 0))
        if url:
            ttk.Button(bar, text='前往下载', style='Primary.TButton',
                       command=lambda: webbrowser.open(url)).pack(side='right')
        if APP_RELEASES_URL:
            ttk.Button(bar, text='发布页', style='Ghost.TButton',
                       command=lambda: webbrowser.open(APP_RELEASES_URL)).pack(
                           side='right', padx=(0, 6))
        if not url:
            tk.Label(bar, text='更新源未提供下载地址，可点「发布页」获取',
                     bg=BG, fg=MUTED, font=('微软雅黑', 9)).pack(side='left')

    def show_about(self):
        about_win = tk.Toplevel(self.root)
        about_win.title(f'关于 {APP_NAME} {APP_VER}')
        about_win.geometry('380x310')
        about_win.resizable(False, False)
        about_win.configure(bg='#ffffff')
        self._apply_icon(about_win)
        about_win.transient(self.root)
        about_win.grab_set()
        self._about_cat(about_win)
        ver = tk.Label(about_win, text=f'{APP_NAME} {APP_VER}', bg='#ffffff', fg=INK,
                       font=('微软雅黑', 16, 'bold'), cursor='hand2')
        ver.pack(pady=(18, 4))
        ver.bind('<Button-1>', self._about_click)
        tk.Label(about_win, text=APP_SUBTITLE, bg='#ffffff', fg=MUTED,
                 font=('微软雅黑', 10)).pack()
        # 软件信息区：版本号 / 作者 / 联系方式
        tk.Frame(about_win, bg='#e7edf4', height=1,
                 width=260).pack(pady=(12, 10))
        tk.Label(about_win, text=f'版本 {APP_VER}', bg='#ffffff', fg=MUTED,
                 font=('微软雅黑', 9)).pack()
        tk.Label(about_win, text=f'作者：{APP_AUTHOR}    联系方式：{APP_CONTACT}',
                 bg='#ffffff', fg=MUTED, font=('微软雅黑', 9)).pack(pady=(6, 0))


if __name__ == '__main__':
    main_root = tk.Tk()
    app = GeoCalcApp(main_root)
    main_root.mainloop()
