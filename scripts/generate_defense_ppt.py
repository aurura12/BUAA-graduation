from __future__ import annotations

import copy
import re
import shutil
import struct
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "周佳琪-王静远-PPT.pptx"
OUTPUT = ROOT / "周佳琪-王静远-本科毕设答辩-v2.pptx"
FIG = ROOT / "figure"

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

for prefix, uri in NS.items():
    if prefix not in {"rel", "ct"}:
        ET.register_namespace(prefix, uri)
ET.register_namespace("", NS["rel"])

EMU = 914400
SLIDE_W = int(13.333333 * EMU)
SLIDE_H = int(7.5 * EMU)

BLUE = "1F4E79"
DEEP = "17365D"
LIGHT = "EAF2F8"
PALE = "F7FAFC"
TEXT = "24364B"
GRAY = "6B7280"
WHITE = "FFFFFF"
ORANGE = "E36C0A"

shape_id = 9000


def qn(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def emu(v: float) -> int:
    return int(v * EMU)


def next_id() -> int:
    global shape_id
    shape_id += 1
    return shape_id


def sub(parent: ET.Element, prefix: str, tag: str, attrs: dict | None = None) -> ET.Element:
    child = ET.SubElement(parent, qn(prefix, tag), attrs or {})
    return child


def make_solid_fill(parent: ET.Element, color: str) -> None:
    solid = sub(parent, "a", "solidFill")
    sub(solid, "a", "srgbClr", {"val": color})


def make_run(parent: ET.Element, text: str, size: int, color: str, bold: bool) -> None:
    run = sub(parent, "a", "r")
    rpr = sub(run, "a", "rPr", {"lang": "zh-CN", "sz": str(size * 100)})
    if bold:
        rpr.set("b", "1")
    make_solid_fill(rpr, color)
    latin = sub(rpr, "a", "latin", {"typeface": "Microsoft YaHei"})
    latin.set("typeface", "Microsoft YaHei")
    sub(rpr, "a", "ea", {"typeface": "Microsoft YaHei"})
    sub(run, "a", "t").text = text


def add_rect(
    tree: ET.Element,
    x: float,
    y: float,
    w: float,
    h: float,
    fill: str,
    line: str | None = None,
    name: str = "Rect",
) -> ET.Element:
    sp = sub(tree, "p", "sp")
    nv = sub(sp, "p", "nvSpPr")
    sub(nv, "p", "cNvPr", {"id": str(next_id()), "name": name})
    sub(nv, "p", "cNvSpPr")
    sub(nv, "p", "nvPr")
    sppr = sub(sp, "p", "spPr")
    xfrm = sub(sppr, "a", "xfrm")
    sub(xfrm, "a", "off", {"x": str(emu(x)), "y": str(emu(y))})
    sub(xfrm, "a", "ext", {"cx": str(emu(w)), "cy": str(emu(h))})
    prst = sub(sppr, "a", "prstGeom", {"prst": "rect"})
    sub(prst, "a", "avLst")
    make_solid_fill(sppr, fill)
    if line:
        ln = sub(sppr, "a", "ln", {"w": "12700"})
        make_solid_fill(ln, line)
    else:
        ln = sub(sppr, "a", "ln")
        sub(ln, "a", "noFill")
    return sp


def add_text(
    tree: ET.Element,
    x: float,
    y: float,
    w: float,
    h: float,
    text: str,
    size: int = 24,
    color: str = TEXT,
    bold: bool = False,
    align: str = "l",
    name: str = "Text",
    margin: float = 0.04,
) -> ET.Element:
    sp = sub(tree, "p", "sp")
    nv = sub(sp, "p", "nvSpPr")
    sub(nv, "p", "cNvPr", {"id": str(next_id()), "name": name})
    sub(nv, "p", "cNvSpPr", {"txBox": "1"})
    sub(nv, "p", "nvPr")
    sppr = sub(sp, "p", "spPr")
    xfrm = sub(sppr, "a", "xfrm")
    sub(xfrm, "a", "off", {"x": str(emu(x)), "y": str(emu(y))})
    sub(xfrm, "a", "ext", {"cx": str(emu(w)), "cy": str(emu(h))})
    prst = sub(sppr, "a", "prstGeom", {"prst": "rect"})
    sub(prst, "a", "avLst")
    ln = sub(sppr, "a", "ln")
    sub(ln, "a", "noFill")
    tx = sub(sp, "p", "txBody")
    body = sub(
        tx,
        "a",
        "bodyPr",
        {
            "wrap": "square",
            "lIns": str(emu(margin)),
            "rIns": str(emu(margin)),
            "tIns": str(emu(margin)),
            "bIns": str(emu(margin)),
        },
    )
    body.set("anchor", "t")
    sub(tx, "a", "lstStyle")
    for line in text.split("\n"):
        p = sub(tx, "a", "p")
        sub(p, "a", "pPr", {"algn": align})
        make_run(p, line, size, color, bold)
    return sp


def clear_slide(root: ET.Element) -> ET.Element:
    sp_tree = root.find(".//p:spTree", NS)
    if sp_tree is None:
        raise RuntimeError("spTree not found")
    keep = list(sp_tree)[:2]
    sp_tree.clear()
    for item in keep:
        sp_tree.append(item)
    return sp_tree


def add_footer(tree: ET.Element, num: int) -> None:
    add_rect(tree, 0, 0, 0.24, 7.5, BLUE)
    add_text(tree, 0.03, 0.22, 0.22, 0.35, f"{num:02d}", size=14, color=WHITE, bold=True, align="c", margin=0)
    add_text(tree, 10.62, 7.12, 2.1, 0.22, "北京航空航天大学", size=10, color=GRAY, align="r", margin=0)


def add_title(tree: ET.Element, num: int, title: str, subtitle: str | None = None) -> None:
    add_footer(tree, num)
    add_text(tree, 1.02, 0.35, 5.3, 0.45, title, size=25, color=DEEP, bold=True)
    add_rect(tree, 1.03, 0.88, 1.02, 0.045, ORANGE)
    if subtitle:
        add_text(tree, 6.8, 0.43, 4.2, 0.28, subtitle, size=11, color=GRAY, align="r", margin=0)


def add_bullets(tree: ET.Element, x: float, y: float, items: list[str], w: float = 5.5, size: int = 19) -> None:
    for idx, item in enumerate(items):
        yy = y + idx * 0.58
        add_rect(tree, x, yy + 0.14, 0.13, 0.13, ORANGE)
        add_text(tree, x + 0.28, yy, w, 0.42, item, size=size, color=TEXT)


def add_metric_card(tree: ET.Element, x: float, y: float, title: str, value: str, note: str, color: str = BLUE) -> None:
    add_rect(tree, x, y, 2.6, 1.25, LIGHT, line="D6E5F0")
    add_text(tree, x + 0.15, y + 0.12, 2.3, 0.25, title, size=12, color=GRAY, bold=True)
    add_text(tree, x + 0.15, y + 0.42, 2.3, 0.38, value, size=25, color=color, bold=True)
    add_text(tree, x + 0.15, y + 0.88, 2.3, 0.22, note, size=9, color=GRAY)


def png_size(path: Path) -> tuple[int, int]:
    with path.open("rb") as f:
        sig = f.read(24)
    if sig[:8] != b"\x89PNG\r\n\x1a\n":
        return (1200, 800)
    return struct.unpack(">II", sig[16:24])


def image_box(path: Path, x: float, y: float, w: float, h: float) -> tuple[float, float, float, float]:
    iw, ih = png_size(path)
    scale = min(w / iw, h / ih)
    nw, nh = iw * scale, ih * scale
    return x + (w - nw) / 2, y + (h - nh) / 2, nw, nh


def add_image_rel(rels_root: ET.Element, target: str) -> str:
    used = []
    for rel in rels_root.findall("rel:Relationship", NS):
        rid = rel.get("Id", "")
        m = re.match(r"rId(\d+)$", rid)
        if m:
            used.append(int(m.group(1)))
    rid = f"rId{(max(used) + 1) if used else 1}"
    ET.SubElement(
        rels_root,
        qn("rel", "Relationship"),
        {
            "Id": rid,
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            "Target": target,
        },
    )
    return rid


def add_picture(tree: ET.Element, rid: str, x: float, y: float, w: float, h: float, name: str = "Picture") -> None:
    pic = sub(tree, "p", "pic")
    nv = sub(pic, "p", "nvPicPr")
    sub(nv, "p", "cNvPr", {"id": str(next_id()), "name": name})
    sub(nv, "p", "cNvPicPr")
    sub(nv, "p", "nvPr")
    blip_fill = sub(pic, "p", "blipFill")
    blip = sub(blip_fill, "a", "blip")
    blip.set(qn("r", "embed"), rid)
    stretch = sub(blip_fill, "a", "stretch")
    sub(stretch, "a", "fillRect")
    sppr = sub(pic, "p", "spPr")
    xfrm = sub(sppr, "a", "xfrm")
    sub(xfrm, "a", "off", {"x": str(emu(x)), "y": str(emu(y))})
    sub(xfrm, "a", "ext", {"cx": str(emu(w)), "cy": str(emu(h))})
    prst = sub(sppr, "a", "prstGeom", {"prst": "rect"})
    sub(prst, "a", "avLst")


def clean_slide_rels(root: ET.Element) -> None:
    for rel in list(root.findall("rel:Relationship", NS)):
        rel_type = rel.get("Type", "")
        if rel_type.endswith("/notesSlide"):
            root.remove(rel)


SLIDES = [
    {
        "kind": "cover",
        "title": "基于大模型的交通智能体设计与实现",
        "subtitle": "本科毕业设计答辩",
    },
    {
        "kind": "contents",
        "items": ["研究背景与目标", "系统设计与实现", "实验结果与分析", "总结与展望"],
    },
    {
        "title": "研究背景",
        "subtitle": "为什么要做这个题",
        "bullets": [
            "交通系统每天都会产生大量传感器数据。",
            "这些数据有流量、占有率和车速，但是用户想问的是路况。",
            "传统报表需要先选点位、再选时间，使用门槛比较高。",
            "我希望把一句自然语言问题，直接转成一次交通分析任务。",
        ],
    },
    {
        "title": "研究目标",
        "subtitle": "我要让系统做到什么",
        "bullets": [
            "用户可以直接说“查一下 1 号传感器上午十点的流量”。",
            "智能体把这句话转成 entity_id 和 time 两个参数。",
            "系统根据问题自动选择查询、分析或预测工具。",
            "最后结果不是一串 JSON，而是一段能读懂的中文报告。",
        ],
        "cards": [
            ("查询", "simple_query", "单点交通状态"),
            ("分析", "deep_analysis", "历史趋势研判"),
            ("预测", "predict_traffic", "未来流量外推"),
        ],
    },
    {
        "title": "研究内容",
        "subtitle": "这篇论文实际做了什么",
        "bullets": [
            "第一步，我把 PeMS08 数据整理成系统能读的格式。",
            "第二步，我用 Qwen3.5-9B 做交通任务的理解和调度。",
            "第三步，我把 PDFormer 预测模型接成智能体工具。",
            "第四步，我用 37 条用例检查系统能不能稳定工作。",
        ],
    },
    {
        "title": "总体技术路线",
        "subtitle": "五层架构",
        "flow": ["感知层", "认知层", "决策层", "执行层", "应用层"],
        "bullets": [
            "感知层把交通数据整理成统一输入。",
            "认知层让大模型理解用户到底想查什么。",
            "决策层负责判断下一步该调用哪个工具。",
            "执行层把查询、历史分析和预测模型封装成服务。",
        ],
    },
    {
        "title": "数据构建",
        "subtitle": "PeMS08 数据处理",
        "bullets": [
            "论文实验使用 PeMS08 交通传感器数据。",
            ".dyna 文件保存每个传感器在各时间点的动态数值。",
            ".rel 文件保存传感器之间的空间关系。",
            "处理后的数据给查询工具和预测模型共同使用。",
        ],
        "image": FIG / "image3.png",
    },
    {
        "title": "系统架构",
        "subtitle": "慢思考 + 快反应",
        "bullets": [
            "慢思考层由 Qwen3.5-9B 负责，它先判断用户意图。",
            "快反应层由三个工具负责，它们只做具体的数据处理。",
            "LangGraph 把“思考”和“执行”接成 ReAct 循环。",
            "这种设计让模型负责判断，让工具负责给真实数据。",
        ],
        "image": FIG / "image14.png",
    },
    {
        "title": "大模型部署",
        "subtitle": "本地推理环境",
        "bullets": [
            "系统核心模型是本地部署的 Qwen3.5-9B。",
            "我用 vLLM 提供 OpenAI 兼容接口。",
            "模型服务运行在本地 GPU 服务器上，所以数据不用传到外部网络。",
            "部署时最大上下文长度设置为 20480 Token。",
        ],
        "cards": [
            ("模型", "Qwen3.5-9B", "负责语义理解"),
            ("框架", "vLLM", "负责本地推理"),
            ("窗口", "20K Token", "支持长上下文"),
        ],
    },
    {
        "title": "状态图推理",
        "subtitle": "LangGraph 调度流程",
        "bullets": [
            "Agent Node 先读用户问题，然后生成工具调用。",
            "Tool Execution Node 接到调用后并行执行工具。",
            "如果工具有返回结果，结果会再交给大模型整理。",
            "如果模型判断不需要工具，流程就直接结束。",
        ],
        "flow": ["用户指令", "Agent Node", "工具执行", "结果汇总", "分析报告"],
    },
    {
        "title": "工具链设计",
        "subtitle": "三个核心工具",
        "bullets": [
            "simple_query 只查某个点位某一刻的状态。",
            "deep_analysis 会多取一段历史趋势，帮助判断拥堵情况。",
            "predict_traffic 会调用 PDFormer，预测未来一小时车流。",
            "三个工具都使用 entity_id 和 time，所以大模型调用比较统一。",
        ],
        "cards": [
            ("simple_query", "流量、占有率、速度", "基础查询"),
            ("deep_analysis", "过去 3 小时趋势", "深度研判"),
            ("predict_traffic", "未来 12 个时间步", "流量预测"),
        ],
    },
    {
        "title": "PDFormer 预测模型",
        "subtitle": "预测工具实现",
        "bullets": [
            "预测工具不是只看目标传感器一个点。",
            "它会收集全路网 170 个传感器的历史数据。",
            "输入使用过去 12 个时间步，也就是过去一小时。",
            "输出是未来 12 个时间步，也就是未来一小时。",
        ],
    },
    {
        "title": "预测实验结果",
        "subtitle": "预测值与真实值对比",
        "bullets": [
            "预测曲线整体能够跟住真实流量变化。",
            "在 PeMS08 上，第 1 个预测步的 MAE 是 11.76。",
            "预测到 1 小时时，MAE 增长到 13.59。",
        ],
        "images": [FIG / "image5.png", FIG / "image6.png"],
    },
    {
        "title": "训练与指标",
        "subtitle": "收敛过程与评价指标",
        "bullets": [
            "训练损失和验证损失整体都在下降。",
            "实验用 MAE、MAPE 和 RMSE 三个指标看误差。",
            "PeMS08 的 masked_MAPE 基本保持在 8% 到 9% 之间。",
        ],
        "images": [FIG / "image9.png", FIG / "image10.png"],
    },
    {
        "title": "系统集成示例",
        "subtitle": "基础流量查询",
        "bullets": [
            "用户问：1 号传感器在 2016 年 7 月 1 日上午十点的流量。",
            "智能体识别出这是 simple_query 任务。",
            "系统返回流量 416 辆/小时、占有率 13.5%、速度 61.0 km/h。",
        ],
        "image": FIG / "image11.png",
    },
    {
        "title": "深度态势分析",
        "subtitle": "历史趋势与拥堵研判",
        "bullets": [
            "用户让系统分析 2 号探头下午五点半的情况。",
            "系统先查当前状态，然后回看历史趋势。",
            "报告识别出 16:45–17:05 期间两次流量高峰。",
            "系统最后判断该路段没有出现拥堵。",
        ],
        "image": FIG / "image12.png",
    },
    {
        "title": "未来流量预测",
        "subtitle": "PDFormer 协同推理",
        "bullets": [
            "用户让系统预测 3 号传感器未来半小时车流。",
            "智能体调用 predict_traffic 工具。",
            "系统返回 08:00–08:30 的 7 个预测点。",
            "报告指出流量从 325 辆波动上升到 350 辆。",
        ],
        "image": FIG / "image13.png",
    },
    {
        "title": "智能体实验",
        "subtitle": "端到端评估",
        "bullets": [
            "我把测试用例分成七类，总共 37 条。",
            "用例覆盖基础查询、深度分析、预测和多工具协同。",
            "37 条用例的意图识别全部正确。",
            "时间格式也覆盖了中文、点分隔、斜杠分隔和相对时间。",
        ],
        "metrics": [
            ("用例数", "37", "覆盖七类场景"),
            ("正确数", "37", "意图识别正确"),
            ("准确率", "100%", "表格统计结果"),
        ],
    },
    {
        "title": "问题与不足",
        "subtitle": "实验中发现的问题",
        "bullets": [
            "多工具任务里会出现数字歧义。",
            "比如 22 号传感器有时会被模型误解成 22 日。",
            "PeMS08 只覆盖 2016 年 7–8 月，所以相对时间查询会受限制。",
            "当前工具还比较少，暂时没有路径规划和信号灯优化。",
        ],
    },
    {
        "title": "总结与展望",
        "subtitle": "主要结论",
        "bullets": [
            "本课题完成了一个基于大模型的交通智能体原型。",
            "系统能把自然语言问题转成工具调用。",
            "系统能把查询、分析和预测结果整理成中文报告。",
            "后续工作会优先接入实时数据，并继续扩展交通管理工具。",
        ],
        "ending": True,
    },
]


def slide_cover(tree: ET.Element, data: dict) -> None:
    add_rect(tree, 0, 0, 13.333, 7.5, PALE)
    add_rect(tree, 0, 0, 3.05, 7.5, BLUE)
    add_rect(tree, 3.05, 0, 0.08, 7.5, ORANGE)
    add_text(tree, 0.48, 0.65, 2.2, 0.5, "BUAA", size=25, color=WHITE, bold=True, align="c")
    add_text(tree, 0.42, 1.18, 2.35, 0.45, "本科毕业设计答辩", size=17, color=WHITE, align="c")
    add_text(tree, 3.8, 2.12, 8.3, 1.15, data["title"], size=34, color=DEEP, bold=True)
    add_rect(tree, 3.86, 3.52, 2.05, 0.06, ORANGE)
    add_text(tree, 3.8, 4.06, 2.2, 0.35, "答辩人：周佳琪", size=17, color=TEXT)
    add_text(tree, 6.0, 4.06, 2.2, 0.35, "导师：王静远", size=17, color=TEXT)
    add_text(tree, 3.8, 4.52, 4.5, 0.32, "计算机科学与技术  计算机学院", size=14, color=GRAY)
    add_text(tree, 3.8, 4.9, 3.0, 0.28, "2026 年 6 月", size=13, color=GRAY)
    add_text(tree, 10.15, 7.12, 2.1, 0.22, "北京航空航天大学", size=10, color=GRAY, align="r", margin=0)


def slide_contents(tree: ET.Element, data: dict) -> None:
    add_rect(tree, 0, 0, 13.333, 7.5, PALE)
    add_rect(tree, 0, 0, 4.2, 7.5, BLUE)
    add_text(tree, 0.72, 1.15, 3.0, 0.65, "CONTENTS", size=30, color=WHITE, bold=True)
    add_text(tree, 0.78, 1.9, 2.1, 0.42, "目录", size=24, color=WHITE, bold=True)
    add_rect(tree, 0.8, 2.52, 1.3, 0.055, ORANGE)
    for i, item in enumerate(data["items"], 1):
        y = 1.52 + (i - 1) * 1.13
        add_text(tree, 4.78, y, 0.55, 0.38, f"{i:02d}", size=20, color=ORANGE, bold=True, align="c")
        add_text(tree, 5.65, y, 4.8, 0.38, item, size=22, color=TEXT, bold=True)
        add_rect(tree, 5.63, y + 0.52, 4.2, 0.02, "D7E3EA")
    add_text(tree, 9.95, 6.85, 2.5, 0.25, "德才兼备  知行合一", size=13, color=GRAY, align="r")


def slide_regular(tree: ET.Element, idx: int, data: dict, rels: ET.Element, media_jobs: list[tuple[Path, str]]) -> None:
    add_rect(tree, 0, 0, 13.333, 7.5, WHITE)
    add_title(tree, idx, data["title"], data.get("subtitle"))

    if data.get("flow"):
        add_bullets(tree, 1.0, 1.45, data.get("bullets", []), w=4.8, size=17)
        xs = [6.45, 7.55, 8.65, 9.75, 10.85]
        for j, label in enumerate(data["flow"]):
            add_rect(tree, xs[j], 2.42, 0.92, 0.92, LIGHT, line="C8DCEA")
            add_text(tree, xs[j] + 0.05, 2.68, 0.82, 0.22, label, size=12, color=DEEP, bold=True, align="c")
            if j < len(xs) - 1:
                add_text(tree, xs[j] + 0.88, 2.65, 0.32, 0.25, "→", size=20, color=ORANGE, bold=True, align="c")
        return

    if data.get("metrics"):
        add_bullets(tree, 1.0, 1.45, data.get("bullets", []), w=5.0, size=17)
        for j, (title, value, note) in enumerate(data["metrics"]):
            add_metric_card(tree, 6.7, 1.55 + j * 1.58, title, value, note, ORANGE if j == 2 else BLUE)
        return

    if data.get("cards"):
        add_bullets(tree, 1.0, 1.35, data.get("bullets", []), w=5.2, size=17)
        for j, (title, value, note) in enumerate(data["cards"]):
            add_metric_card(tree, 6.65, 1.35 + j * 1.42, title, value, note)
        return

    if data.get("images"):
        add_bullets(tree, 0.9, 1.28, data.get("bullets", []), w=4.6, size=16)
        boxes = [(5.65, 1.45, 3.35, 4.6), (9.15, 1.45, 3.35, 4.6)]
        for img, box in zip(data["images"], boxes):
            if not img.exists():
                continue
            target = f"../media/defense_{img.name}"
            rid = add_image_rel(rels, target)
            media_jobs.append((img, f"ppt/media/defense_{img.name}"))
            ix, iy, iw, ih = image_box(img, *box)
            add_picture(tree, rid, ix, iy, iw, ih, img.name)
        return

    if data.get("image"):
        img = data["image"]
        add_bullets(tree, 0.9, 1.25, data.get("bullets", []), w=4.5, size=16)
        if img.exists():
            target = f"../media/defense_{img.name}"
            rid = add_image_rel(rels, target)
            media_jobs.append((img, f"ppt/media/defense_{img.name}"))
            ix, iy, iw, ih = image_box(img, 5.55, 1.35, 6.85, 4.95)
            add_picture(tree, rid, ix, iy, iw, ih, img.name)
        return

    add_bullets(tree, 1.28, 1.42, data.get("bullets", []), w=9.7, size=20)
    if data.get("ending"):
        add_text(tree, 7.95, 5.85, 2.1, 0.32, "答辩人：周佳琪", size=13, color=GRAY)
        add_text(tree, 10.0, 5.85, 2.0, 0.32, "导师：王静远", size=13, color=GRAY)


def prune_presentation(zip_data: dict[str, bytes]) -> None:
    pres_path = "ppt/presentation.xml"
    rels_path = "ppt/_rels/presentation.xml.rels"
    pres = ET.fromstring(zip_data[pres_path])
    rels = ET.fromstring(zip_data[rels_path])
    sld_lst = pres.find("p:sldIdLst", NS)
    if sld_lst is not None:
        sld_ids = list(sld_lst.findall("p:sldId", NS))
        for idx, item in enumerate(sld_ids, 1):
            if idx > len(SLIDES):
                rid = item.get(qn("r", "id"))
                sld_lst.remove(item)
                for rel in list(rels.findall("rel:Relationship", NS)):
                    if rel.get("Id") == rid:
                        rels.remove(rel)
    zip_data[pres_path] = ET.tostring(pres, encoding="utf-8", xml_declaration=True)
    zip_data[rels_path] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)

    ct_path = "[Content_Types].xml"
    ct = ET.fromstring(zip_data[ct_path])
    for override in list(ct.findall("ct:Override", NS)):
        part = override.get("PartName", "")
        m_slide = re.match(r"/ppt/slides/slide(\d+)\.xml$", part)
        m_notes = re.match(r"/ppt/notesSlides/notesSlide(\d+)\.xml$", part)
        if (m_slide and int(m_slide.group(1)) > len(SLIDES)) or m_notes:
            ct.remove(override)
    zip_data[ct_path] = ET.tostring(ct, encoding="utf-8", xml_declaration=True)

    app_path = "docProps/app.xml"
    if app_path in zip_data:
        app = ET.fromstring(zip_data[app_path])
        for child in app.iter():
            if child.tag.endswith("Slides"):
                child.text = str(len(SLIDES))
        zip_data[app_path] = ET.tostring(app, encoding="utf-8", xml_declaration=True)


def main() -> None:
    if not TEMPLATE.exists():
        raise FileNotFoundError(TEMPLATE)

    with zipfile.ZipFile(TEMPLATE, "r") as zin:
        zip_data = {name: zin.read(name) for name in zin.namelist()}

    media_jobs: list[tuple[Path, str]] = []

    for idx, data in enumerate(SLIDES, 1):
        slide_path = f"ppt/slides/slide{idx}.xml"
        rel_path = f"ppt/slides/_rels/slide{idx}.xml.rels"
        root = ET.fromstring(zip_data[slide_path])
        rels = ET.fromstring(zip_data[rel_path]) if rel_path in zip_data else ET.Element(qn("rel", "Relationships"))
        clean_slide_rels(rels)
        tree = clear_slide(root)
        if data.get("kind") == "cover":
            slide_cover(tree, data)
        elif data.get("kind") == "contents":
            slide_contents(tree, data)
        else:
            slide_regular(tree, idx, data, rels, media_jobs)
        zip_data[slide_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        zip_data[rel_path] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)

    prune_presentation(zip_data)

    remove_patterns = [
        re.compile(r"ppt/slides/slide(\d+)\.xml$"),
        re.compile(r"ppt/slides/_rels/slide(\d+)\.xml\.rels$"),
        re.compile(r"ppt/notesSlides/.*"),
    ]

    for src, arcname in media_jobs:
        zip_data[arcname] = src.read_bytes()

    tmp = OUTPUT.with_suffix(".tmp.pptx")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in zip_data.items():
            skip = False
            for pat in remove_patterns[:2]:
                match = pat.match(name)
                if match and int(match.group(1)) > len(SLIDES):
                    skip = True
            if remove_patterns[2].match(name):
                skip = True
            if skip:
                continue
            zout.writestr(name, data)

    if OUTPUT.exists():
        OUTPUT.unlink()
    shutil.move(tmp, OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
