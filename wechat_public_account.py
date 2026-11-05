# -*- coding: utf-8 -*-
"""微信桌面端自动化主流程（公众号采集用）——按步骤推进，每步可单独跑。

设计原则
--------
1. 确定性的事交给代码：进程 / 窗口 / HWND / 显示器几何 / UIA，快、免费、可重跑；
2. 判不准的事交给视觉 LLM：现在是二维码页还是已登录？任务栏微信图标在哪？
   统一协议——模型只输出 JSON，程序按字段分发，所以界面改版不用改代码；
3. 每一步都截图存档 + 落 JSON 结果，出问题能回看，不靠猜。

LLM 决策协议（所有步骤共用）
---------------------------
{
  "state": "qrcode|phone_confirm|main|loading|error|other",
  "evidence": "图中关键文字（<=20字）",
  "action": {
      "type": "click|double_click|wait|none",
      "target": "要点的东西",
      "center": [x, y],              # 图片像素坐标（备用）
      "center_norm": [0.0, 0.0],     # 相对整图的比例（优先，抗缩放）
      "double": false
  },
  "confidence": 0.95
}

用法
----
    # 只跑步骤1：确认微信启动并进入主界面（含扫码轮询）
    D:\\miniconda3\\envs\\ufo\\python.exe -X utf8 wechat_public_account.py --stage launch

    # 只跑步骤2：确认微信窗口可见、没最小化、没跑出屏幕（副屏也照常可操作，不搬窗口）
    python wechat_public_account.py --stage ready        # screen 是旧别名

    # 只跑步骤3：确认在主聊天窗，不在（设置/附属窗口/子页面/白屏）就退回来
    python wechat_public_account.py --stage chat

    # 只跑步骤4：「公众号」页 → 右上角放大镜 → 搜关键词 → 读结果 → 点开对应公众号主页
    python wechat_public_account.py --stage search --keyword 建约车评
    python wechat_public_account.py --stage search --keyword 建约车评 --no-open   # 只搜不点开

    # 步骤2→3→4 一口气连跑：每步先做前置判定，不满足就回退上一步重做
    python wechat_public_account.py --stage flow --keyword 建约车评

    # 全流程（步骤1 -> 步骤2 -> 步骤3）
    python wechat_public_account.py

    python wechat_public_account.py --scan-timeout 0     # 不等待扫码（调试）
    python wechat_public_account.py --dry-run            # 只判定、不点击不启动
"""

from __future__ import annotations

import argparse
import base64
import io
import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import requests
from PIL import Image, ImageGrab

HERE = Path(__file__).resolve().parent
SHOT_DIR = HERE / "screenshots"
OUT_DIR = HERE / "out"
LOG_DIR = HERE / "logs"
for _d in (SHOT_DIR, OUT_DIR, LOG_DIR):
    _d.mkdir(exist_ok=True)

ENV_FILE = HERE / ".env"


def _load_dotenv(path: Path = ENV_FILE) -> int:
    """把技能目录下的 `.env` 灌进 `os.environ`（**已存在的环境变量优先，不覆盖**）。

    密钥只放 `.env`，不进代码也不进 git（见 `.gitignore`）。格式：`KEY=VALUE`，`#` 开头是注释，
    值两侧的引号可选。只用标准库，不引 python-dotenv。返回实际写入的键数。
    """
    try:
        raw = path.read_text(encoding="utf-8")
    except OSError:
        return 0
    n = 0
    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        key, val = key.strip(), val.strip()
        if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
            val = val[1:-1]
        if key and key not in os.environ:
            os.environ[key] = val
            n += 1
    return n


LOADED_ENV_KEYS = _load_dotenv()

VL_BASE = os.environ.get("VL_BASE", "http://119.84.39.2:11435/v1")
# 密钥不在代码里：只从 .env / 环境变量取；缺了就在真正要调用时报清楚该往哪写
VL_KEY = os.environ.get("VL_KEY", "")
VL_MODEL = os.environ.get("VL_MODEL", "qwen3.8-27b")
VL_MAX_WIDTH = int(os.environ.get("VL_MAX_WIDTH", "1600"))
# 该端点默认带 thinking，会吃掉 max_tokens 导致 content 为空；判图用不到长推理，默认关掉
VL_THINKING = os.environ.get("VL_THINKING", "0") not in ("0", "false", "False", "")


def _vl_headers() -> dict:
    if not VL_KEY:
        raise RuntimeError(
            f"缺 VL_KEY：把密钥写进 {ENV_FILE}（一行 `VL_KEY=sk-...`）或设同名环境变量；"
            "代码里不再内置密钥"
        )
    return {"Authorization": f"Bearer {VL_KEY}", "Content-Type": "application/json"}

MAIN_CLASS = "mmui::MainWindow"
WX_EXE_NAMES = ("Weixin.exe", "WeChat.exe")
_LOG_FILE = LOG_DIR / f"run_{datetime.now():%Y%m%d_%H%M%S}.log"


# --------------------------------------------------------------------------
# 日志
# --------------------------------------------------------------------------
def log(msg: str) -> None:
    """日志走 stderr，stdout 只留给最终 JSON。"""
    line = f"[{datetime.now():%H:%M:%S}] {msg}"
    print(line, file=sys.stderr, flush=True)
    try:
        with _LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def _setup_console() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            if (stream.encoding or "").lower() not in ("utf-8", "utf8"):
                stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]
        except (AttributeError, ValueError):
            pass


# --------------------------------------------------------------------------
# 屏幕 / 截图（多屏：原点可能不是 0,0，必须按虚拟屏原点裁剪）
# --------------------------------------------------------------------------
def dpi_awareness() -> int:
    """当前进程的 DPI 感知级别：0= unaware / 1= system / 2= per-monitor。

    必须显式声明 argtypes/restype：不声明时 `byref(c_int)` 会被当 32 位指针传，
    返回值也可能被截断，实测会误报成 1（实际是 2）。
    """
    import ctypes
    from ctypes import byref, c_int

    try:
        shcore = ctypes.windll.shcore
        shcore.GetProcessDpiAwareness.argtypes = [ctypes.c_void_p, ctypes.POINTER(c_int)]
        shcore.GetProcessDpiAwareness.restype = ctypes.c_long
        val = c_int()
        if shcore.GetProcessDpiAwareness(None, byref(val)) == 0:
            return int(val.value)
    except Exception:  # noqa: BLE001
        pass
    try:
        return 1 if ctypes.windll.user32.IsProcessDPIAware() else 0
    except Exception:  # noqa: BLE001
        return -1


def enable_dpi_awareness() -> int:
    """把进程设成 per-monitor DPI 感知（必须的，见下）。

    坑（2026-11-02，缩放 125%）：进程若对 DPI 无感知，`GetWindowRect` 拿到的是**逻辑坐标**
    (1920 屏报 1536)，而 `ImageGrab` 截回来的是**物理像素** (1920)。两套坐标系混用后，
    "从截图算出来的点"换算成点击点会按 1.25 倍系统性偏移（离屏幕原点越远偏得越多，
    几百像素级），表现就是"点了不准/点了没反应"。设成 per-monitor 感知后两边统一为物理像素。
    """
    try:
        import ctypes

        # Win10 1703+：PER_MONITOR_AWARE_V2 = -4
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return dpi_awareness()
    except Exception:  # noqa: BLE001
        pass
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:  # noqa: BLE001
        try:
            import ctypes

            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:  # noqa: BLE001
            pass
    return dpi_awareness()


# 进程一导入就定死 DPI 感知，后面所有窗口坐标/截图/点击都在物理像素同一套坐标系里
DPI = enable_dpi_awareness()


# --------------------------------------------------------------------------
# 无障碍（UIA）健康检查：微信有没有建无障碍树，取决于**它是怎么被启动的**（见 §11.30）
# --------------------------------------------------------------------------
SPI_GETSCREENREADER = 0x0046
SPI_SETSCREENREADER = 0x0047
SPIF_SENDCHANGE = 0x0002
SPIF_UPDATEINIFILE = 0x0001


# 本进程是否已经连过 UIA。**启动微信的进程在启动前绝不能连 UIA**（§11.30）：
# 连了之后微信收到的第一次 WM_GETOBJECT 早于 MMUI 就绪，Qt 的 UIA provider 会把"空树"缓存住，
# 之后无论怎么查都只剩 Qt 空壳（Ctrl+Alt+W 也救不回来）。
_UIA_TOUCHED = False


def _mark_uia() -> None:
    global _UIA_TOUCHED
    _UIA_TOUCHED = True


def screen_reader_flag() -> int:
    """系统「读屏软件」标志（SPI_GETSCREENREADER）：1=有读屏软件，0=没有，-1=读失败。"""
    import ctypes

    v = ctypes.c_int()
    try:
        ctypes.windll.user32.SystemParametersInfoW(SPI_GETSCREENREADER, 0, ctypes.byref(v), 0)
    except Exception:  # noqa: BLE001
        return -1
    return int(v.value)


def set_screen_reader_flag(value: int) -> bool:
    """置/清读屏标志（写用户配置并广播 WM_SETTINGCHANGE）。

    **注意：它不是「UIA 读不到微信」的原因，两次独立验证都排除掉了。**
    §11.30（11-03）的 2×2 对照先排除一次；11-04 又按「**启动前**置 1 → 再启动」严格重做一次
    （`screen_reader_flag=1` 已确认写进去），连续重启的实例仍然没有无障碍树。
    Weixin.dll 里确实有 `clicfg_accessible_check_narrator` / `mmui::accessible::AccessibleChecker`
    / `kDisableAccessibleKey` 这组符号，但它们不是决定因素。
    这里只保留作诊断/对照用（`--stage a11y` 会把它当一项数据打出来）。
    """
    import ctypes

    try:
        return bool(
            ctypes.windll.user32.SystemParametersInfoW(
                SPI_SETSCREENREADER, int(value), None, SPIF_SENDCHANGE | SPIF_UPDATEINIFILE
            )
        )
    except Exception:  # noqa: BLE001
        return False


def ensure_screen_reader_flag() -> dict:
    """启动微信之前调用：保证读屏标志=1，否则微信不会建无障碍树。"""
    before = screen_reader_flag()
    if before == 1:
        return {"before": before, "after": 1, "set_ok": True, "changed": False}
    ok = set_screen_reader_flag(1)
    return {"before": before, "after": screen_reader_flag(), "set_ok": ok, "changed": True}


MMUI_SUBWINDOW_CLASS = "MMUIRenderSubWindowHW"


def mmui_tree_probe() -> dict:
    """裸 UIA 直接量「MMUI 内容树建了没有」，不依赖任何类名假设。

    微信 4.x 是两层窗口：外层 Qt 窗（class 带 `Qt`）+ 内层自绘子窗 `MMUIRenderSubWindowHW`。
    会话列表那套控件属于 MMUI，**它的无障碍树是运行时才挂上去的**：没挂时从任何入口进去都是
    0 子节点；挂上时子节点立刻可见。所以「树在不在」最硬的指标是**这个子窗的第一层子节点数**。

    为什么不用类名判（2026-11-04 实测）：有的实例 UIA 里根本没有 `mmui::MainWindow` 这个类名，
    树挂在 Qt 主窗上（`find_main_window()` 靠 HWND 兜底读到 10 条会话），而旧 `a11y_diagnosis()`
    只按类名找 → 把**能用的实例**报成 `a11y_disabled`（同一时刻 `session_items=10` 与
    `list_items=-1` 同时成立）。

    只查 `TreeScope_Children`（第一层）且限量，绝不无参 `descendants()`（护栏 9）。
    """
    out: dict = {
        "main_hwnd": None,
        "main_class": None,
        "main_children": -1,
        "mmui_subwindow_hwnd": None,
        "mmui_subwindow_children": -1,
    }
    found = find_main_hwnd()
    if found is None:
        return out
    out["main_hwnd"] = found["hwnd"]
    out["main_class"] = found["class"]
    try:
        import comtypes.client
        import win32gui

        _mark_uia()  # 裸 comtypes 也算连过 UIA，防呆标记要打（§11.30）

        uia_lib = comtypes.client.GetModule("UIAutomationCore.dll")
        uia = comtypes.client.CreateObject(uia_lib.CUIAutomation, interface=uia_lib.IUIAutomation)
        true_cond = uia.CreateTrueCondition()

        el = uia.ElementFromHandle(found["hwnd"])
        kids = None
        if el is not None:
            kids = el.FindAll(4, true_cond)  # 4 = TreeScope_Children
            out["main_children"] = kids.Length if kids is not None else 0

        subs: list[int] = []

        def _cb(hwnd, _p):
            try:
                if (win32gui.GetClassName(hwnd) or "") == MMUI_SUBWINDOW_CLASS:
                    subs.append(int(hwnd))
            except Exception:  # noqa: BLE001
                pass
            return True

        win32gui.EnumChildWindows(found["hwnd"], _cb, None)
        if subs:
            out["mmui_subwindow_hwnd"] = subs[0]

        # 数子窗第一层子节点：**优先从主窗的 UIA 子节点里找同名元素**，再退回按子 HWND 取元素。
        # 实测 `ElementFromHandle(子 HWND)` 时好时坏（返回 None 且不报错；单独调用能拿到 0，
        # 放进 a11y_diagnosis 的调用链里就变 None，2026-11-04），从父元素往下走则稳定。
        if kids is not None:
            for i in range(kids.Length):
                try:
                    k = kids.GetElement(i)
                    if (k.CurrentClassName or "") != MMUI_SUBWINDOW_CLASS:
                        continue
                    sk = k.FindAll(4, true_cond)
                    out["mmui_subwindow_children"] = sk.Length if sk is not None else 0
                    break
                except Exception:  # noqa: BLE001
                    continue
        if out["mmui_subwindow_children"] < 0 and subs:
            sel = uia.ElementFromHandle(subs[0])
            if sel is not None:
                sk = sel.FindAll(4, true_cond)
                out["mmui_subwindow_children"] = sk.Length if sk is not None else 0
    except Exception as exc:  # noqa: BLE001
        out["probe_error"] = f"{type(exc).__name__}: {exc}"
    return out


def a11y_diagnosis() -> dict:
    """UIA 读不到东西时给出可判定的原因，而不是笼统的「找不到主窗」。

    verdict:
      ok            —— 读得到会话条目（不管树挂在 `mmui::MainWindow` 还是 Qt 主窗上）
      not_running   —— 微信没在运行
      a11y_disabled —— 微信起来了但 MMUI 内容树没建（子窗 0 子节点、会话读不到）
      tree_broken   —— 树挂上了（子窗有子节点）但会话条目为 0（见 §11.1）
    """
    flag = screen_reader_flag()
    pids = wechat_pids()
    wins = wechat_top_windows(pids) if pids else []
    probe = mmui_tree_probe() if pids else {}
    mmui = False
    items = -1
    sessions = -1
    via = None
    try:
        from pywinauto import Desktop

        for w in Desktop(backend="uia").windows():
            try:
                if w.element_info.class_name == "mmui::MainWindow":
                    mmui = True
                    via = "class:mmui::MainWindow"
                    items = len(w.descendants(control_type="ListItem"))
                    sessions = session_count(w)
                    break
            except Exception:  # noqa: BLE001
                continue
        if not mmui:
            # 兜底（与 find_main_window 一致）：树可能挂在 Qt 主窗上，类名不是 mmui::MainWindow
            found = find_main_hwnd()
            if found is not None:
                w = Desktop(backend="uia").window(handle=found["hwnd"])
                via = f"hwnd:{found['hwnd']}"
                items = len(w.descendants(control_type="ListItem"))
                sessions = session_count(w)
    except Exception:  # noqa: BLE001
        pass

    # 注意别写 `probe.get(...) or -1`：0 是**有效值**（确认没挂树），会被 `or` 吞成 -1
    raw_sub = probe.get("mmui_subwindow_children", -1)
    sub_children = int(raw_sub) if isinstance(raw_sub, (int, float)) else -1
    out: dict = {
        "screen_reader_flag": flag,
        "wechat_pids": len(pids),
        "top_windows": [{"class": w["class"], "title": w["title"]} for w in wins[:3]],
        "mmui_main_window": mmui,
        "matched_via": via,
        "list_items": items,
        "session_list_items": sessions,
        "main_hwnd": probe.get("main_hwnd"),
        "main_class": probe.get("main_class"),
        "main_children": probe.get("main_children", -1),
        "mmui_subwindow_hwnd": probe.get("mmui_subwindow_hwnd"),
        "mmui_subwindow_children": sub_children,
        "probe_error": probe.get("probe_error"),
        "plugin_procs": len(plugin_residue_pids()),
    }
    tree_up = (items or 0) > 0 or (sessions or 0) > 0 or sub_children > 0
    if tree_up:
        out.update(verdict="ok", cause=None, fix=None)
    elif not pids:
        out.update(verdict="not_running", cause="微信没在运行", fix="跑步骤 1 启动微信")
    elif sub_children <= 0:  # 0=确认没挂树；-1=连子窗的 UIA 元素都取不到，同样按没树处理
        cause = (
            f"这个微信实例没建 MMUI 无障碍树：Qt 主窗在（hwnd={probe.get('main_hwnd')} "
            f"class={probe.get('main_class')}），但自绘子窗 {MMUI_SUBWINDOW_CLASS} 第一层 "
            f"{sub_children if sub_children == 0 else '取不到元素'} 子节点，"
            f"会话 ListItem 也读不到。"
            "**根因（2026-11-05 定位）**：微信内嵌 Qt 的 `QAccessible` 运行时 gate 字节 = 0。"
            "冷启动没检测到读屏器时它就是 0，`WM_GETOBJECT` 只给 Qt 外壳（对自绘子窗发 `WM_GETOBJECT` "
            "返回 lParam=0，即那个 HWND 根本没有 UIA provider）；微信重启/升级后该字节复位为 0，"
            "所以表现为「昨天那个实例能用几小时、今天新实例全是空壳」。"
            "**已实测排除**：启动方式（Popen/ShellExecute/explorer 都起过没树的实例，也起过有树的）、"
            "系统读屏标志（启动前置 1 仍没树）、屏幕与 DPI、`Ctrl+Alt+W` 重建、清残留插件进程"
            "（清干净后连续 3 次重启，45 秒连读 0/45 轮可读）。"
        )
        out.update(
            verdict="a11y_disabled",
            cause=cause,
            fix=(
                "跑 `--stage a11y --hot-gate`：把 Weixin.dll 里的 Qt accessibility gate 字节热写成 1，"
                "**不用重启微信**，树当场物化（本机 4.2.0.11 命中 Weixin.dll+0xb507c08）"
            ),
        )
    else:
        out.update(
            verdict="tree_broken",
            cause=(
                f"MMUI 子窗有子节点（{sub_children}）但会话条目为 0（无障碍树被读崩/未刷新，见 §11.1）"
            ),
            fix="recover_uia_tree()（Ctrl+Alt+W）重建，必要时 --repair-a11y 重启微信",
        )
    return out


# --------------------------------------------------------------------------
# Qt accessibility gate：热激活（不重启微信就能让 mmui 控件树物化）
# --------------------------------------------------------------------------
# 微信 4.x 内嵌 Qt 的 `QAccessible` 查询前会看一个运行时 active gate 字节：冷启动没检测到读屏器
# 时它是 0，`WM_GETOBJECT` 只给 Qt 外壳（`Qt51514QWindowIcon` + `MMUIRenderSubWindowHW` 空壳，
# 对子窗发 `WM_GETOBJECT` 返回 lParam=0 = 那个 HWND 根本没有 provider）；把它热写成 1，**当前进程
# 立刻**物化出 `mmui::*` 完整控件树（2026-11-05 实测：写后 `session_list_items` 0→11、
# `list_items`=21，连续复查稳定）。微信重启/升级后该字节复位为 0——**这才是「昨天能用几小时、
# 今天新实例全是空壳」的真正原因**，与启动方式（Popen/ShellExecute）、系统读屏标志、屏幕/DPI、
# `Ctrl+Alt+W`、清残留插件进程全都无关（以上全部实测排除，见 §11.30 的更正）。
# 特征码与「写一个候选→校验→不对就回滚换下一个」的做法取自 wechatauto-replica（Apache-2.0）。
# 风险：写第三方进程内存可能触发微信风控（该项目作者账号被强制重登过一次），只在采集机上用。
GATE_PATTERN = re.compile(rb"\x48\x85\xc9\x0f\x84....\x80\x3d(?P<disp>.{4})\x00\x0f\x84", re.DOTALL)
GATE_CORE_STRING = b"qt.accessibility.core"
# LEA rip-relative：48 8D <modrm(mod=00,rm=101)> <disp32>
_LEA_RE = re.compile(
    rb"\x48\x8d[\x05\x0d\x15\x1d\x25\x2d\x35\x3d\x45\x4d\x55\x5d\x65\x6d\x75\x7d](?P<disp>....)"
)
IMAGE_SCN_MEM_EXECUTE = 0x20000000
IMAGE_SCN_MEM_WRITE = 0x80000000
# 已实测命中的版本兜底表（扫描优先，扫不到才用它）
GATE_RVA_BY_VERSION = {"4.2.0.11": 0x0B507C08}
# gate RVA 按 Weixin.dll 身份（版本目录+大小+mtime）落盘：196MB 的 DLL 每次重扫不值
GATE_CACHE_FILE = OUT_DIR / "gate_cache.json"


def _dll_identity(dll_path: str) -> str:
    try:
        st = os.stat(dll_path)
        return f"{Path(dll_path).parent.name}|{st.st_size}|{int(st.st_mtime)}"
    except OSError:
        return dll_path


def _gate_cache() -> dict:
    try:
        with GATE_CACHE_FILE.open(encoding="utf-8") as f:
            raw = json.load(f)
        return raw if isinstance(raw, dict) else {}
    except (OSError, ValueError):
        return {}


def _gate_cache_put(identity: str, **fields) -> None:
    data = _gate_cache()
    entry = data.setdefault(identity, {})
    entry.update(fields)
    try:
        GATE_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        with GATE_CACHE_FILE.open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
    except OSError:
        pass


def _pe_sections(data: bytes) -> list[dict]:
    if data[:2] != b"MZ":
        return []
    pe_off = int.from_bytes(data[0x3C:0x40], "little")
    if data[pe_off : pe_off + 4] != b"PE\0\0":
        return []
    n_sec = int.from_bytes(data[pe_off + 6 : pe_off + 8], "little")
    tab = pe_off + 24 + int.from_bytes(data[pe_off + 20 : pe_off + 22], "little")
    out = []
    for i in range(n_sec):
        e = tab + i * 40
        vsize, va, rawsize, raw = (int.from_bytes(data[e + 8 + j * 4 : e + 12 + j * 4], "little") for j in range(4))
        out.append(
            {
                "name": data[e : e + 8].rstrip(b"\0").decode("ascii", "replace"),
                "va": va,
                "vsize": vsize,
                "raw": raw,
                "rawsize": rawsize,
                "chars": int.from_bytes(data[e + 36 : e + 40], "little"),
            }
        )
    return out


def _off_to_rva(sections: list[dict], off: int) -> int | None:
    for s in sections:
        if s["raw"] <= off < s["raw"] + s["rawsize"]:
            return s["va"] + (off - s["raw"])
    return None


def _sec_for_rva(sections: list[dict], rva: int) -> dict | None:
    for s in sections:
        if s["va"] <= rva < s["va"] + max(s["vsize"], s["rawsize"]):
            return s
    return None


def _lea_xrefs(data: bytes, sections: list[dict], target_rva: int) -> list[int]:
    """可执行段里以 RIP 相对寻址引用 `target_rva` 的 LEA 指令 RVA（用来给候选排序）。"""
    out: list[int] = []
    for mt in _LEA_RE.finditer(data):
        rva = _off_to_rva(sections, mt.start())
        if rva is None:
            continue
        if rva + 7 + int.from_bytes(mt.group("disp"), "little", signed=True) == target_rva:
            out.append(rva)
    return out


def gate_candidates(dll_path: str) -> list[int]:
    """按可信度返回 gate 字节的全部 RVA 候选（越靠前越可能是真 gate）。

    版本升级后 RVA 会漂移，所以返回候选序列，由 `hot_activate_gate()` 逐个写、
    用「mmui 树是否真的物化」判定，成功者按 DLL 身份落盘，下次直接用。
    """
    identity = _dll_identity(dll_path)
    cached = _gate_cache().get(identity, {}).get("candidates")
    if isinstance(cached, list):
        return [int(c) for c in cached]
    try:
        data = Path(dll_path).read_bytes()
    except OSError:
        return []
    sections = _pe_sections(data)
    if not sections:
        return []
    core_off = data.find(GATE_CORE_STRING)
    core_rva = _off_to_rva(sections, core_off) if core_off >= 0 else None
    core_xrefs = _lea_xrefs(data, sections, core_rva) if core_rva is not None else []

    found: list[tuple[int, int]] = []
    for mt in GATE_PATTERN.finditer(data):
        mrva = _off_to_rva(sections, mt.start())
        drva = _off_to_rva(sections, mt.start("disp"))
        if mrva is None or drva is None:
            continue
        msec = _sec_for_rva(sections, mrva)
        if not msec or not (msec["chars"] & IMAGE_SCN_MEM_EXECUTE):
            continue
        disp = int.from_bytes(mt.group("disp"), "little", signed=True)
        target = (drva - 2) + 7 + disp  # 80 3D <disp32> 00
        tsec = _sec_for_rva(sections, target)
        if not tsec or not (tsec["chars"] & IMAGE_SCN_MEM_WRITE):
            continue
        dist = min((abs(mrva - x) for x in core_xrefs), default=0x7FFFFFFF)
        found.append((dist, target))
    if not found:
        _gate_cache_put(identity, candidates=[])
        return []
    found.sort()
    near = [rva for dist, rva in found if dist <= 0x20000]
    ordered = list(dict.fromkeys(near or [rva for _d, rva in found]))
    _gate_cache_put(identity, candidates=ordered)
    return ordered


def gate_candidate_rvas(dll_path: str) -> list[int]:
    """候选顺序：已验证缓存 > 特征扫描 > 版本兜底表。"""
    out: list[int] = []
    identity = _dll_identity(dll_path)
    verified = _gate_cache().get(identity, {}).get("verified")
    if isinstance(verified, int):
        out.append(int(verified))
    out += [r for r in gate_candidates(dll_path) if r not in out]
    fallback = GATE_RVA_BY_VERSION.get(Path(dll_path).parent.name)
    if fallback is not None and fallback not in out:
        out.append(int(fallback))
    return out


def psapi_modules(pid: int) -> list[dict]:
    """进程模块 + 基址。psutil 7 删了 `Process.modules()`，pywin32 没有
    `GetModuleInformation`，Toolhelp 的 ctypes 结构体在 x64 下又容易错位——
    用 PSAPI：`EnumProcessModulesEx` 返回的 HMODULE 本身就是基址。"""
    import ctypes
    import ctypes.wintypes as wt

    k32 = ctypes.windll.kernel32
    k32.OpenProcess.restype = wt.HANDLE
    k32.OpenProcess.argtypes = [wt.DWORD, wt.BOOL, wt.DWORD]
    k32.K32EnumProcessModulesEx.restype = wt.BOOL
    k32.K32EnumProcessModulesEx.argtypes = [
        wt.HANDLE,
        ctypes.POINTER(wt.HMODULE),
        wt.DWORD,
        ctypes.POINTER(wt.DWORD),
        wt.DWORD,
    ]
    k32.K32GetModuleFileNameExW.restype = wt.DWORD
    k32.K32GetModuleFileNameExW.argtypes = [wt.HANDLE, wt.HMODULE, wt.LPWSTR, wt.DWORD]
    k32.CloseHandle.argtypes = [wt.HANDLE]

    h = k32.OpenProcess(0x0400 | 0x0010, False, pid)  # QUERY_INFORMATION | VM_READ
    if not h:
        return []
    try:
        cap = 2048
        arr = (wt.HMODULE * cap)()
        needed = wt.DWORD(0)
        if not k32.K32EnumProcessModulesEx(h, arr, ctypes.sizeof(arr), ctypes.byref(needed), 3):
            return []
        out: list[dict] = []
        for i in range(min(cap, needed.value // ctypes.sizeof(wt.HMODULE))):
            buf = ctypes.create_unicode_buffer(32768)
            if not k32.K32GetModuleFileNameExW(h, arr[i], buf, 32768):
                continue
            path = buf.value
            out.append({"name": path.rsplit("\\", 1)[-1], "path": path, "base": int(arr[i] or 0)})
        return out
    finally:
        k32.CloseHandle(h)


def hot_activate_gate(pid: int, verify: bool = True, max_candidates: int = 4) -> dict:
    """把运行中微信进程里的 Qt accessibility gate 字节热写成 1（不重启、不注入代码）。

    每个候选写完就用「会话列表读得到吗」校验：不对就**回滚原值**再试下一个；
    成功则按 DLL 身份记进 `out/gate_cache.json`，下次直接命中、不必重扫 196MB DLL。
    """
    mods = [mo for mo in psapi_modules(pid) if mo["name"].lower() == "weixin.dll"]
    if not mods:
        return {"ok": False, "error": f"PID {pid} 里没有 Weixin.dll"}
    mod = mods[0]
    cands = gate_candidate_rvas(mod["path"])[:max_candidates]
    if not cands:
        return {"ok": False, "error": f"扫不到 gate 候选：{mod['path']}", "dll": mod["path"]}

    import ctypes

    k32 = ctypes.windll.kernel32
    access = 0x0400 | 0x0010 | 0x0020 | 0x0008  # QUERY | VM_READ | VM_WRITE | VM_OPERATION
    k32.OpenProcess.restype = ctypes.c_void_p
    k32.OpenProcess.argtypes = [ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong]
    handle = k32.OpenProcess(access, 0, pid)
    if not handle:
        return {"ok": False, "error": f"OpenProcess 失败 PID {pid}", "tried": cands}
    tried: list[dict] = []
    try:

        def rbyte(addr: int) -> int | None:
            buf = (ctypes.c_ubyte * 1)()
            got = ctypes.c_size_t(0)
            k32.ReadProcessMemory.argtypes = [
                ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)
            ]
            ok = k32.ReadProcessMemory(handle, ctypes.c_void_p(addr), buf, 1, ctypes.byref(got))
            return int(buf[0]) if ok and got.value == 1 else None

        def wbyte(addr: int, val: int) -> bool:
            buf = (ctypes.c_ubyte * 1)(val & 0xFF)
            put = ctypes.c_size_t(0)
            k32.WriteProcessMemory.argtypes = k32.ReadProcessMemory.argtypes
            ok = k32.WriteProcessMemory(handle, ctypes.c_void_p(addr), buf, 1, ctypes.byref(put))
            return bool(ok and put.value == 1)

        for rva in cands:
            addr = mod["base"] + rva
            cur = rbyte(addr)
            if cur is None:
                tried.append({"rva": hex(rva), "read": False})
                continue
            wrote = False
            if cur != 1:
                if not wbyte(addr, 1):
                    tried.append({"rva": hex(rva), "from": cur, "write": False})
                    continue
                wrote = True
            if not verify:
                k32.CloseHandle.argtypes = [ctypes.c_void_p]
                k32.CloseHandle(handle)
                return {"ok": True, "rva": hex(rva), "from": cur, "verified": False, "tried": tried}
            time.sleep(1.0)
            if mmui_tree_ready(2.0):
                _gate_cache_put(_dll_identity(mod["path"]), verified=int(rva))
                tried.append({"rva": hex(rva), "from": cur, "hit": True})
                return {
                    "ok": True,
                    "rva": hex(rva),
                    "from": cur,
                    "verified": True,
                    "dll": mod["path"],
                    "tried": tried,
                }
            if wrote:
                wbyte(addr, cur)  # 候选不对：回滚
            tried.append({"rva": hex(rva), "from": cur, "hit": False})
        return {"ok": False, "error": f"{len(cands)} 个候选都没让 mmui 树物化", "tried": tried}
    finally:
        k32.CloseHandle.argtypes = [ctypes.c_void_p]
        k32.CloseHandle(handle)


def mmui_tree_ready(timeout: float = 1.5) -> bool:
    """mmui 控件树是否可用（热激活的校验判据）。

    **不能借用 `find_main_window()`**：那个函数在找不到时会调 `maybe_wake_gate()`，
    而这里正是被 `wake_a11y()` 调用来做校验的，会绕成递归。
    """
    _mark_uia()
    from pywinauto import Desktop

    deadline = time.time() + max(timeout, 0)
    while True:
        try:
            for w in Desktop(backend="uia").windows():
                try:
                    if w.element_info.class_name == MAIN_CLASS and session_count(w) > 0:
                        return True
                except Exception:  # noqa: BLE001
                    continue
        except Exception:  # noqa: BLE001
            pass
        found = find_main_hwnd()  # 树挂在 Qt 窗上的实例（class 不是 mmui::MainWindow）
        if found is not None:
            try:
                w = Desktop(backend="uia").window(handle=found["hwnd"])
                if session_count(w) > 0:
                    return True
            except Exception:  # noqa: BLE001
                pass
        if time.time() >= deadline:
            return False
        time.sleep(0.3)


def wake_a11y(force: bool = False) -> dict:
    """让微信的 mmui 无障碍树可用：置读屏标志 + 逐个微信进程热写 gate 字节并校验。

    已经是好的实例（能读到会话）默认直接跳过，`force=True` 强制重写一遍。
    """
    if mmui_tree_ready(1.5) and not force:
        # 必须带 `ready`：`maybe_wake_gate()` 按这个键判定要不要重试找窗口，缺键会被当成激活失败
        return {"ok": True, "ready": True, "skipped": "会话列表已可读，无需热激活"}
    flag = ensure_screen_reader_flag()
    attempts: list[dict] = []
    for pid in wechat_pids():
        try:
            r = hot_activate_gate(pid)
        except Exception as exc:  # noqa: BLE001
            r = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
        attempts.append({"pid": pid, **r})
        if r.get("ok"):
            break
    ok = any(a.get("ok") for a in attempts)
    return {"ok": ok, "screen_reader_flag": flag, "attempts": attempts, "ready": mmui_tree_ready(2.0)}


# 微信会自己把 gate 字节复位。2026-11-05 用 `tests/probe_gate_reset.py` 量过两轮（只看 1→0）：
# **复位周期约 38~43 秒**（run1 在 17.0s / 55.2s，run2 在 62.8s）。复位后 mmui 树不是立刻消失，
# 还能残留几秒（会看到 `gate=0 sessions=10`），下一拍才塌成 `sessions=-1`。
# 所以"激活一次管几分钟"是错的，长流程必须持续补。限流按这个周期定：
# 间隔 ≥8s（塌掉后一轮轮询内就能救回）、上限 60 次（够跑完整 1→6 含步骤 6 逐篇抓正文）。
_GATE_WAKE = {"n": 0, "t": 0.0, "gap": 8.0, "max": 60}


def maybe_wake_gate(reason: str = "") -> bool:
    if not wechat_pids():  # 微信没在运行，别白跑一次 PSAPI + 写内存
        return False
    if _GATE_WAKE["n"] >= _GATE_WAKE["max"] or time.time() - _GATE_WAKE["t"] < _GATE_WAKE["gap"]:
        return False
    _GATE_WAKE["n"] += 1
    _GATE_WAKE["t"] = time.time()
    r = wake_a11y()
    log(f"自动热激活 gate（{reason}）：ok={r.get('ok')} ready={r.get('ready')}")
    return bool(r.get("ready"))


def virtual_screen() -> tuple[int, int, int, int]:
    import win32api

    return (
        win32api.GetSystemMetrics(76),  # SM_XVIRTUALSCREEN
        win32api.GetSystemMetrics(77),  # SM_YVIRTUALSCREEN
        win32api.GetSystemMetrics(78),  # SM_CXVIRTUALSCREEN
        win32api.GetSystemMetrics(79),  # SM_CYVIRTUALSCREEN
    )


def virtual_screen_rect() -> list[int]:
    x, y, w, h = virtual_screen()
    return [x, y, x + w, y + h]


def monitors() -> list[dict]:
    import win32api
    import win32con

    out: list[dict] = []
    for hmon, _hdc, _rect in win32api.EnumDisplayMonitors(None, None):
        info = win32api.GetMonitorInfo(hmon)
        out.append(
            {
                "rect": list(info["Monitor"]),
                "work": list(info["Work"]),
                "primary": bool(info["Flags"] & win32con.MONITORINFOF_PRIMARY),
            }
        )
    return out


def primary_work_area() -> list[int]:
    for m in monitors():
        if m["primary"]:
            return [int(v) for v in m["work"]]
    return [0, 0, 1920, 1040]


def primary_monitor_rect() -> list[int]:
    """主屏整块（含任务栏），视觉找任务栏图标时用它，避免多屏坐标歧义。"""
    for m in monitors():
        if m["primary"]:
            return [int(v) for v in m["rect"]]
    return [0, 0, 1920, 1080]


def grab(rect: list[int] | None = None) -> Image.Image:
    """抓全虚拟屏再裁剪，避免 bbox 直接截到别的屏。"""
    vx0, vy0, _, _ = virtual_screen()
    full = ImageGrab.grab(all_screens=True)
    if rect is None:
        return full
    l, t, r, b = [int(v) for v in rect]
    return full.crop((l - vx0, t - vy0, r - vx0, b - vy0))


def save_shot(img: Image.Image, name: str) -> str:
    path = SHOT_DIR / f"{name}.png"
    img.save(path)
    return str(path)


# --------------------------------------------------------------------------
# 视觉 LLM
# --------------------------------------------------------------------------
def _encode(img: Image.Image, max_width: int = VL_MAX_WIDTH) -> tuple[str, tuple[int, int]]:
    """缩到 max_width 内再用 JPEG 发（省 token），返回 data url 与实际发送尺寸。"""
    w, h = img.size
    sent = img
    if w > max_width:
        sent = img.resize((max_width, max(1, int(h * max_width / w))), Image.LANCZOS)
    buf = io.BytesIO()
    sent.convert("RGB").save(buf, "JPEG", quality=88)
    b64 = base64.b64encode(buf.getvalue()).decode()
    return f"data:image/jpeg;base64,{b64}", sent.size


def _parse_json(raw: str) -> dict:
    s = (raw or "").strip()
    s = re.sub(r"^```[a-zA-Z]*\s*", "", s)
    s = re.sub(r"\s*```$", "", s).strip()
    try:
        obj = json.loads(s)
        return obj if isinstance(obj, dict) else {"_parse_error": True, "_raw": s[:300]}
    except json.JSONDecodeError:
        pass
    # 兜底：抓第一个配平的 {...}
    depth, start = 0, -1
    for i, ch in enumerate(s):
        if ch == "{":
            if depth == 0:
                start = i
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0 and start >= 0:
                try:
                    obj = json.loads(s[start : i + 1])
                    if isinstance(obj, dict):
                        return obj
                except json.JSONDecodeError:
                    pass
                depth, start = 0, -1
    return {"_parse_error": True, "_raw": s[:300]}


def llm_json(img: Image.Image, prompt: str, max_tokens: int = 1200, retries: int = 2) -> dict:
    """问视觉模型要一个 JSON 对象。失败时返回 {"_error": ...}，不抛异常。

    实测（2026-11-02）：该端点带 thinking，reasoning 会吃掉 token，
    max_tokens 偏小时 content 可能整段为空 —— 所以默认关掉 thinking
    （enable_thinking=False，响应快 10 倍且更稳），并在 content 为空时
    退回 reasoning_content 里抓 JSON。
    """
    url, sent_size = _encode(img)
    payload = {
        "model": VL_MODEL,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": url}},
                ],
            }
        ],
        "max_tokens": max_tokens,
        "temperature": 0,
        "chat_template_kwargs": {"enable_thinking": VL_THINKING},
    }
    last = ""
    for i in range(retries + 1):
        try:
            resp = requests.post(
                f"{VL_BASE}/chat/completions",
                json=payload,
                headers=_vl_headers(),
                timeout=180,
            )
            if resp.status_code != 200:
                last = f"HTTP {resp.status_code}: {resp.text[:200]}"
                time.sleep(1.0)
                continue
            choice = (resp.json().get("choices") or [{}])[0]
            msg = choice.get("message") or {}
            raw = (msg.get("content") or "").strip()
            if not raw:  # thinking 型回答可能只给了 reasoning_content
                raw = (msg.get("reasoning_content") or "").strip()
            if not raw:
                last = f"空回答 finish={choice.get('finish_reason')}"
                payload["max_tokens"] = int(payload["max_tokens"]) * 2
                continue
            data = _parse_json(raw)
            data["_raw"] = raw[:400]
            data["_sent_size"] = [int(sent_size[0]), int(sent_size[1])]
            return data
        except Exception as exc:  # noqa: BLE001
            last = f"{type(exc).__name__}: {exc}"
            time.sleep(1.0)
    log(f"LLM 调用失败: {last}")
    return {"_error": last}


def _salvage_strings(raw: str, key: str) -> list[str]:
    """JSON 被 max_tokens 截断时，从 `"key": [ ...` 里逐个抢救字符串元素。

    实测（2026-11-03）：5584 字的全文整理成 paragraphs，输出比输入还长，
    4000 token 直接截断 → 整个 JSON 解析失败、正文全丢。截断了也要把已生成的段落救回来。
    """
    i = raw.find(f'"{key}"')
    j = raw.find("[", i) if i >= 0 else -1
    if j < 0:
        return []
    out: list[str] = []
    k = j + 1
    while k < len(raw):
        while k < len(raw) and raw[k] in " \n\r\t,":
            k += 1
        if k >= len(raw) or raw[k] != '"':
            break
        k += 1
        buf: list[str] = []
        closed = False
        while k < len(raw):
            ch = raw[k]
            if ch == "\\":
                nxt = raw[k + 1 : k + 2]
                buf.append({"n": "\n", "t": "\t", "r": "\r"}.get(nxt, nxt))
                k += 2
                continue
            if ch == '"':
                closed = True
                k += 1
                break
            buf.append(ch)
            k += 1
        if not closed:
            break  # 截断在字符串中间 → 这半条丢掉，前面的照收
        s = "".join(buf).strip()
        if s:
            out.append(s)
    return out


def llm_text_json(prompt: str, max_tokens: int = 8000, retries: int = 2) -> dict:
    """纯文本问模型要一个 JSON 对象（不带图，用于整理剪贴板里的全文）。"""
    payload = {
        "model": VL_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": max_tokens,
        "temperature": 0,
        "chat_template_kwargs": {"enable_thinking": VL_THINKING},
    }
    last = ""
    for _ in range(retries + 1):
        try:
            resp = requests.post(
                f"{VL_BASE}/chat/completions",
                json=payload,
                headers=_vl_headers(),
                timeout=180,
            )
            if resp.status_code != 200:
                last = f"HTTP {resp.status_code}: {resp.text[:200]}"
                time.sleep(1.0)
                continue
            choice = (resp.json().get("choices") or [{}])[0]
            msg = choice.get("message") or {}
            raw = (msg.get("content") or "").strip()
            if not raw:
                raw = (msg.get("reasoning_content") or "").strip()
            if not raw:
                last = f"空回答 finish={choice.get('finish_reason')}"
                payload["max_tokens"] = int(payload["max_tokens"]) * 2
                continue
            data = _parse_json(raw)
            data["_raw"] = raw  # 截断时给 `_salvage_strings` 抢救用
            return data
        except Exception as exc:  # noqa: BLE001
            last = f"{type(exc).__name__}: {exc}"
            time.sleep(1.0)
    log(f"LLM(纯文本) 调用失败: {last}")
    return {"_error": last}


def action_point(data: dict, img_rect: list[int]) -> tuple[int, int] | None:
    """把模型给的坐标换算成屏幕绝对坐标。优先 center_norm（相对比例，抗缩放）。"""
    act = data.get("action")
    if not isinstance(act, dict):
        return None
    if act.get("type") not in ("click", "double_click"):
        return None  # 模型说不用点，别把 [0,0] 当成真坐标
    l, t, r, b = [int(v) for v in img_rect]

    norm = act.get("center_norm")
    if isinstance(norm, (list, tuple)) and len(norm) == 2:
        try:
            nx, ny = float(norm[0]), float(norm[1])
            if 0.0 <= nx <= 1.0 and 0.0 <= ny <= 1.0:
                return int(l + nx * (r - l)), int(t + ny * (b - t))
        except (TypeError, ValueError):
            pass

    center = act.get("center")
    if isinstance(center, (list, tuple)) and len(center) == 2:
        try:
            cx, cy = float(center[0]), float(center[1])
        except (TypeError, ValueError):
            return None
        sent = data.get("_sent_size") or [r - l, b - t]
        sx = (r - l) / float(sent[0] or 1)
        sy = (b - t) / float(sent[1] or 1)
        return int(l + cx * sx), int(t + cy * sy)
    return None


# --------------------------------------------------------------------------
# 进程 / 窗口 / 点击
# --------------------------------------------------------------------------
def _wx_names() -> set[str]:
    return {n.lower() for n in WX_EXE_NAMES}


def wechat_pids() -> list[int]:
    import psutil

    pids: list[int] = []
    for p in psutil.process_iter(["name"]):
        try:
            if (p.info["name"] or "").lower() in _wx_names():
                pids.append(int(p.pid))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return pids


def wechat_top_windows(pids: list[int] | None = None) -> list[dict]:
    """按 pid 过滤出微信的可见顶层窗口（真实 HWND，Qt 窗口也在这里）。"""
    import win32gui
    import win32process

    pid_set = set(pids if pids is not None else wechat_pids())
    if not pid_set:
        return []
    found: list[dict] = []

    def _cb(hwnd, _param):
        try:
            _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid not in pid_set or not win32gui.IsWindowVisible(hwnd):
                return True
            l, t, r, b = win32gui.GetWindowRect(hwnd)
            if (r - l) < 200 or (b - t) < 200:
                return True
            found.append(
                {
                    "hwnd": int(hwnd),
                    "pid": int(pid),
                    "class": win32gui.GetClassName(hwnd),
                    "title": win32gui.GetWindowText(hwnd),
                    "rect": [int(l), int(t), int(r), int(b)],
                    "iconic": bool(win32gui.IsIconic(hwnd)),
                }
            )
        except Exception:  # noqa: BLE001
            pass
        return True

    win32gui.EnumWindows(_cb, None)
    return found


def biggest_wechat_hwnd() -> int:
    """微信可见顶层窗口里面积最大的那个 HWND。

    UIA 枚举不到 `mmui::MainWindow` 时（Qt 无障碍树没建起来，见 §11.23），
    用它当兜底锚点：热键重建 UIA 树、取窗口矩形都靠它。
    """
    ws = [w for w in wechat_top_windows() if not w.get("iconic")]
    if not ws:
        return 0
    w = max(ws, key=lambda x: (x["rect"][2] - x["rect"][0]) * (x["rect"][3] - x["rect"][1]))
    return int(w["hwnd"])


def _hwnd_rect(hwnd: int) -> list[int] | None:
    import win32gui

    try:
        l, t, r, b = win32gui.GetWindowRect(int(hwnd))
        return [int(l), int(t), int(r), int(b)]
    except Exception:  # noqa: BLE001
        return None


def find_main_window(timeout: float = 2.0):
    """UIA 找微信主窗口。微信 4.x 的 mmui::MainWindow 不是顶层 HWND，只能靠 UIA 枚举。"""
    _mark_uia()
    from pywinauto import Desktop

    deadline = time.time() + max(timeout, 0)
    fallback = None
    while True:
        for w in Desktop(backend="uia").windows():
            try:
                if w.element_info.class_name == MAIN_CLASS:
                    return w
            except Exception:  # noqa: BLE001
                continue
        # 兜底（只在超时后返回）：实测有实例的 UIA class 报成 `Qt51514QWindowIcon`
        # （不是 `mmui::MainWindow`），按 class 名硬找会永远找不到 → 步骤2 直接判"窗口不存在"。
        # 但**不能一有机会就用**：刚登录时 UIA 子树还没建好，过早返回一个空壳会把
        # step1 的"等主界面"变成"立刻认定登录成功（session=0）"，后面步骤全卡在会话列表为空。
        if fallback is None:
            found = find_main_hwnd()
            if found is not None:
                try:
                    w = Desktop(backend="uia").window(handle=found["hwnd"])
                    # 只有子树真读得到会话才算"主窗可用了"，空壳一律当没找到
                    if w.element_info.class_name and session_count(w) > 0:
                        fallback = w
                except Exception:  # noqa: BLE001
                    pass
        if time.time() >= deadline:
            if fallback is None and maybe_wake_gate("find_main_window 超时"):
                # gate 是当场生效的：激活完立刻再找一遍，别让调用方以为微信没登录
                try:
                    for w in Desktop(backend="uia").windows():
                        try:
                            if w.element_info.class_name == MAIN_CLASS:
                                return w
                        except Exception:  # noqa: BLE001
                            continue
                except Exception:  # noqa: BLE001
                    pass
                found = find_main_hwnd()
                if found is not None:
                    try:
                        w = Desktop(backend="uia").window(handle=found["hwnd"])
                        if session_count(w) > 0:
                            return w
                    except Exception:  # noqa: BLE001
                        pass
            return fallback
        time.sleep(0.3)


def window_rect(win) -> list[int]:
    """取窗口屏幕矩形。

    微信重建窗口后旧的 UIA 元素会失效（rectangle() 抛 COMError 灾难性故障），
    所以 UIA 拿不到就退回 win32 的 HWND 矩形。
    """
    try:
        r = win.rectangle()
    except Exception:  # noqa: BLE001
        found = find_main_hwnd()
        if found:
            return list(found["rect"])
        raise
    try:
        return [int(r.left), int(r.top), int(r.right), int(r.bottom)]
    except AttributeError:
        return [int(v) for v in r]


def window_hwnd(win) -> int:
    for getter in (lambda: win.handle, lambda: win.element_info.handle):
        try:
            h = getter()
            if h:
                return int(h)
        except Exception:  # noqa: BLE001
            continue
    return 0


def session_count(win) -> int:
    """会话列表条目数——有值即已登录进入主界面。

    关键：必须把 `control_type` 交给 UIA 服务端过滤（`descendants(control_type=...)`
    内部是 FindAll，返回元素极少，实测 0.01s）；绝不能用无参 `descendants()` 取回
    全部元素再在 Python 里逐个读属性 —— 微信 4.x 的 Qt 无障碍桥会被拖到无响应
    （实测主进程 CPU 100%、吃 3.4GB、窗口全部失联），`COMError 灾难性故障` 就是前兆。
    另外本机 pywinauto 的 UIA 条件不支持 auto_id/automation_id，只能按类型筛后本地挑。
    """
    try:
        for lst in win.descendants(control_type="List"):
            if lst.element_info.automation_id == "session_list":
                return len(lst.children(control_type="ListItem"))
        return 0
    except Exception:  # noqa: BLE001
        return 0


def wechat_hung() -> bool | None:
    """微信主窗是否无响应。None 表示找不到主窗。

    注意：`win32gui` 里没有 IsHungAppWindow，必须走 ctypes 调 user32。
    """
    import ctypes

    found = find_main_hwnd()
    if found is None:
        return None
    try:
        return bool(ctypes.windll.user32.IsHungAppWindow(found["hwnd"]))
    except Exception:  # noqa: BLE001
        return None


def click_screen(x: int, y: int, double: bool = False) -> None:
    import win32api
    import win32con

    win32api.SetCursorPos((int(x), int(y)))
    time.sleep(0.25)

    def _one() -> None:
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(0.12)

    _one()
    if double:
        _one()
    time.sleep(0.4)


def force_foreground(hwnd: int) -> dict:
    """恢复并置前。SetForegroundWindow 常被系统拒绝，用 ALT 键技巧解锁。"""
    import win32api
    import win32con
    import win32gui

    out: dict = {"hwnd": hwnd}
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            out["restored"] = True
    except Exception as exc:  # noqa: BLE001
        out["restore_error"] = str(exc)
    try:
        win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
        win32gui.SetForegroundWindow(hwnd)
        win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
        out["foreground"] = True
    except Exception as exc:  # noqa: BLE001
        out["foreground_error"] = str(exc)
    time.sleep(0.8)
    try:
        out["fg_is_wechat"] = win32gui.GetForegroundWindow() == hwnd
    except Exception:  # noqa: BLE001
        pass
    return out


def bring_to_front(hwnd: int) -> dict:
    """把微信主窗真正怼到最前（含"被 IDE 挡住"的情况）。

    坑：窗口"可见"不等于"没被遮挡"。若 VS Code 盖在微信上面，
    grab(窗口矩形) 截到的是 VS Code，视觉模型就会判成 other；
    而 UIA 的 click_input 走真实鼠标，也会点到 VS Code 上。
    所以：先 SetForegroundWindow，确认失败就用 TOPMOST 抖一下（不常驻置顶）。
    """
    import win32con
    import win32gui

    out = force_foreground(hwnd)
    if out.get("fg_is_wechat"):
        return out
    try:
        flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, flags)
        time.sleep(0.2)
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, flags)
        out["topmost_bump"] = True
        out["fg_after_bump"] = win32gui.GetForegroundWindow() == hwnd
    except Exception as exc:  # noqa: BLE001
        out["topmost_error"] = str(exc)
    time.sleep(0.5)
    out["fg_final"] = win32gui.GetForegroundWindow() == hwnd
    return out


def find_main_hwnd() -> dict | None:
    """按 win32 找微信主窗（Qt 主窗），返回 {hwnd, rect, title, class, iconic, visible}。

    不依赖 UIA：微信重建窗口后旧 UIA 元素会失效（COMError 灾难性故障），
    而 HWND 一直在，用它做矩形/可见性判定更稳。
    """
    import win32gui
    import win32process

    pids = set(wechat_pids())
    if not pids:
        return None

    cands: list[dict] = []

    def _cb(hwnd, _param):
        try:
            _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid not in pids:
                return True
            cls = win32gui.GetClassName(hwnd) or ""
            title = win32gui.GetWindowText(hwnd) or ""
            if "Qt" not in cls or not title.startswith(("微信", "Weixin", "WeChat")):
                return True
            l, t, r, b = win32gui.GetWindowRect(hwnd)
            if (r - l) < 600 or (b - t) < 400:
                return True
            cands.append(
                {
                    "hwnd": int(hwnd),
                    "rect": [int(l), int(t), int(r), int(b)],
                    "title": title,
                    "class": cls,
                    "iconic": bool(win32gui.IsIconic(hwnd)),
                    "visible": bool(win32gui.IsWindowVisible(hwnd)),
                }
            )
        except Exception:  # noqa: BLE001
            pass
        return True

    win32gui.EnumWindows(_cb, None)
    if not cands:
        return None
    return max(cands, key=lambda c: (c["rect"][2] - c["rect"][0]) * (c["rect"][3] - c["rect"][1]))


def force_show_main_hwnd() -> dict:
    """确定性唤出：即使 UIA 枚举不到（窗口被隐藏到托盘/不可见），
    仍按 win32 找微信的 Qt 主窗 HWND 并 Show/Restore + 置前。"""
    import win32con
    import win32gui

    if not wechat_pids():
        return {"ok": False, "error": "微信进程不存在"}
    found = find_main_hwnd()
    if found is None:
        return {"ok": False, "error": "没有找到微信主窗 HWND（可能真的没启动）"}

    hwnd = found["hwnd"]
    out: dict = {"hwnd": hwnd, "rect_before": found["rect"], "visible_before": found["visible"]}
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        out["shown"] = True
    except Exception as exc:  # noqa: BLE001
        out["show_error"] = str(exc)
    time.sleep(0.6)
    out.update(force_foreground(hwnd))
    out["ok"] = True
    return out


def kill_wechat() -> list[int]:
    import psutil

    killed: list[int] = []
    for p in psutil.process_iter(["name"]):
        try:
            if (p.info["name"] or "").lower() in _wx_names():
                p.kill()
                killed.append(int(p.pid))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    time.sleep(1.5)
    return killed


# 微信自己的插件/CEF 运行时进程。`kill_wechat()` 只按 Weixin.exe / WeChat.exe 杀，**杀不到它们**，
# 它们会跨实例存活（2026-11-04 实测：残留一个 `RadiumWMPF\WeChatAppEx.exe` 时，新起的微信
# 连续三次都不建 MMUI 无障碍树；把它一起清掉后重启，立刻 `session_items=10`）。
WECHAT_PLUGIN_PROC_NAMES = ("wechatappex.exe",)


def plugin_residue_pids(only_xwechat: bool = True) -> list[dict]:
    """列出属于微信的插件进程（`WeChatAppEx.exe`，路径在 `...\\xwechat\\xplugin\\...`）。

    同名进程还有一个来自 Androws 模拟器（`...\\Androws\\WmpfRuntime\\...`）：跟微信无关，
    且普通权限杀不掉（`Access denied`），默认按路径排除掉，免得把恢复动作误导到死路上。
    """
    import psutil

    out: list[dict] = []
    for p in psutil.process_iter(["name", "exe"]):
        try:
            if (p.info["name"] or "").lower() not in WECHAT_PLUGIN_PROC_NAMES:
                continue
            path = str(p.info.get("exe") or "")
            if only_xwechat and "xwechat" not in path.lower():
                continue
            out.append({"pid": int(p.pid), "exe": path})
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return out


def kill_plugin_residue(only_xwechat: bool = True) -> dict:
    """杀掉微信插件进程（先杀插件，再杀微信本体，插件就没机会把旧状态带进新实例）。"""
    import psutil

    targets = plugin_residue_pids(only_xwechat)
    killed: list[int] = []
    denied: list[int] = []
    for t in targets:
        try:
            psutil.Process(t["pid"]).kill()
            killed.append(t["pid"])
        except psutil.NoSuchProcess:
            continue
        except psutil.AccessDenied:
            denied.append(t["pid"])
    time.sleep(1.0)
    return {"found": len(targets), "killed": killed, "access_denied": denied}


def repair_a11y_tree(exe: str = "", dry_run: bool = False) -> dict:
    """实测有效的恢复动作：清残留插件进程 → 杀微信 → ShellExecute 重启（不碰 UIA）。

    注意顺序：本函数**不能在之前调用任何 UIA**（`_UIA_TOUCHED` 的防呆留着，虽然 §11.30 的
    归因已被 11-04 复测推翻，但「先启动、后连 UIA」至少无害且必要）。
    重启后只等窗口出现（win32），不做登录判定——交给调用方复查。
    """
    if _UIA_TOUCHED:
        log("警告：本进程已经连过 UIA，这样重启的微信可能没有无障碍树；建议单独跑一次 --stage launch")
    exe_path = exe or find_wechat_exe()
    if not exe_path:
        return {"ok": False, "error": "未找到微信可执行文件，请用 --wechat-exe 指定"}
    if dry_run:
        return {"ok": False, "dry_run": True, "exe": exe_path, "plugin_residue": plugin_residue_pids()}

    plugins = kill_plugin_residue()
    log(
        f"清残留插件进程 found={plugins['found']} killed={len(plugins['killed'])} "
        f"access_denied={len(plugins['access_denied'])}"
    )
    killed = kill_wechat()
    log(f"杀微信本体 pids={len(killed)}")
    time.sleep(2.0)

    info = launch_wechat(exe_path)
    log(f"重启微信 {info}")
    if not info.get("ok"):
        return {"ok": False, "error": info.get("error") or "启动失败", "plugins": plugins, "killed": killed}

    appear = time.time() + 30
    while time.time() < appear:
        if wechat_top_windows():
            break
        time.sleep(0.5)
    return {
        "ok": True,
        "plugins": plugins,
        "killed_weixin": len(killed),
        "launch": info,
        "top_windows": len(wechat_top_windows()),
    }


# --------------------------------------------------------------------------
# 微信可执行文件定位 / 启动
# --------------------------------------------------------------------------
_EXE_CANDIDATES = (
    r"C:\Program Files\Tencent\Weixin\Weixin.exe",
    r"C:\Program Files (x86)\Tencent\Weixin\Weixin.exe",
    r"C:\Program Files\Tencent\WeChat\WeChat.exe",
    r"C:\Program Files (x86)\Tencent\WeChat\WeChat.exe",
    r"D:\Program Files\Tencent\Weixin\Weixin.exe",
    r"D:\Program Files (x86)\Tencent\Weixin\Weixin.exe",
    r"D:\Tencent\Weixin\Weixin.exe",
)


def find_wechat_exe() -> str:
    import winreg

    # 1) App Paths
    app_paths = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
    for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for exe in WX_EXE_NAMES:
            try:
                with winreg.OpenKey(root, f"{app_paths}\\{exe}") as key:
                    path, _ = winreg.QueryValueEx(key, "")
                if path and Path(path).exists():
                    return str(path)
            except OSError:
                continue

    # 2) Tencent 安装项
    for root, sub in (
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Tencent\Weixin"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Tencent\Weixin"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Tencent\Weixin"),
    ):
        try:
            with winreg.OpenKey(root, sub) as key:
                for val in ("InstallPath", "InstallDir", ""):
                    try:
                        path, _ = winreg.QueryValueEx(key, val)
                    except OSError:
                        continue
                    for exe in WX_EXE_NAMES:
                        cand = Path(str(path)) / exe
                        if cand.exists():
                            return str(cand)
        except OSError:
            continue

    for cand in _EXE_CANDIDATES:
        if Path(cand).exists():
            return cand
    return ""


def launch_wechat(exe: str, dry_run: bool = False) -> dict:
    if not exe:
        return {"ok": False, "error": "exe 为空"}
    if dry_run:
        return {"ok": False, "dry_run": True, "exe": exe}
    # 防呆（§11.30 留下的规矩）：**启动微信的进程在这之前不要连 UIA**。
    # 但归因已在 11-04 复测中被推翻：同一个脚本用 Popen（连过 UIA）起过 `session_items=10` 的实例，
    # 用 ShellExecute（没连 UIA）也连续起过 3 个没树的实例 → 「有没有树」是**实例级概率事件**，
    # 不由启动方式决定。这条规矩留着只是无害，别再拿它当诊断结论（真结论见 a11y_diagnosis 的 cause）。
    if _UIA_TOUCHED:
        log(
            "警告：本进程启动微信前已经连过 UIA —— 这样起的微信不会有无障碍树（§11.30）；"
            "请用没碰过 UIA 的进程启动（例如单独跑 --stage launch），或完全退出微信后重跑"
        )
    # 用 ShellExecute（等同在资源管理器里双击）启动，别用 CreateProcess 直接拉。
    try:
        os.startfile(exe)  # noqa: S606
        return {"ok": True, "exe": exe, "via": "shellexecute"}
    except Exception as exc:  # noqa: BLE001
        log(f"ShellExecute 启动失败，退回 Popen（注意：这样起的微信读不到 UIA）：{exc}")
        try:
            subprocess.Popen([exe], close_fds=True, cwd=str(Path(exe).parent))
            return {"ok": True, "exe": exe, "via": "popen", "warning": "Popen 启动，无障碍树可能为空"}
        except Exception as exc2:  # noqa: BLE001
            return {"ok": False, "exe": exe, "error": f"{type(exc2).__name__}: {exc2}"}


# --------------------------------------------------------------------------
# LLM 提示词
# --------------------------------------------------------------------------
STATE_PROMPT = """这是 Windows 桌面上微信客户端的截图。请判断当前处于哪个阶段，并给出下一步该点哪里。
只输出一个 JSON 对象，不要解释、不要 Markdown 代码块：

{
  "state": "welcome|qrcode|phone_confirm|main|loading|error|other",
  "evidence": "图中关键文字（不超过20字）",
  "action": {"type": "click|wait|none", "target": "要点的东西", "center_norm": [0.0, 0.0], "double": false},
  "confidence": 0.0
}

判定规则：
- 有用户头像/昵称，并且下方有一个绿色按钮写着"进入微信" → welcome
- 屏幕中有大二维码，且有"扫码登录"/"手机扫码"等字样 → qrcode
- 有"请在手机上确认登录"/"已扫码"等提示 → phone_confirm
- 能看到左侧会话列表、顶部搜索框、右侧聊天区 → main
- 只有启动画面/空白/加载中 → loading
- 报错、网络异常、二维码已失效 → error

动作规则（重要，程序会照着点）：
- state=welcome → 必须给 action.type="click"，target="进入微信"，center_norm 指向那个绿色按钮中心
- state=other → 如果画面上明确有可点的按钮（如"进入微信""登录""确定""重试""刷新"），必须给 click 动作和坐标
- state=qrcode/phone_confirm/loading/main → 需要人扫码或已在主界面，action.type 用 "none"

坐标要求：center_norm 用相对整张图宽高的 0~1 比例。
"""

REVEAL_PROMPT = """这是 Windows 桌面全屏截图（若有多块屏幕则是拼接的整幅图）。
请找出微信，判断怎样才能让微信主窗口显示到前台。只输出一个 JSON 对象，不要解释：

{
  "state": "main_window_visible|minimized|taskbar_icon|tray_icon|not_found",
  "evidence": "看到的关键内容（不超过20字）",
  "action": {"type": "click|double_click|none", "target": "要点的图标", "center_norm": [0.0, 0.0]},
  "confidence": 0.0
}

说明：
- 微信主窗口已经完整显示在桌面上 → main_window_visible，action.type = none
- 任务栏上有微信图标但窗口没显示（最小化）→ taskbar_icon / minimized，给出该图标位置
- 只在右下角托盘区看到微信图标 → tray_icon，给出该图标位置（需要双击才能打开）
- 完全找不到微信 → not_found

坐标要求：center_norm 用相对整张图宽高的 0~1 比例。
"""


# --------------------------------------------------------------------------
# 步骤 1：确认微信已启动并进入主界面
# --------------------------------------------------------------------------
def _wechat_view() -> tuple[list[int] | None, dict | None]:
    """微信当前画面的截图区域：优先取面积最大的微信窗口，没有窗口就整屏。"""
    wins = wechat_top_windows()
    if not wins:
        return None, None
    best = max(wins, key=lambda w: (w["rect"][2] - w["rect"][0]) * (w["rect"][3] - w["rect"][1]))
    return best["rect"], best


def probe_state(tag: str) -> dict:
    """截图 + LLM 判定当前处于哪个阶段。"""
    rect, win = _wechat_view()
    img = grab(rect) if rect else grab()
    img_rect = rect or virtual_screen_rect()
    path = save_shot(img, f"{tag}_{datetime.now():%H%M%S}")
    data = llm_json(img, STATE_PROMPT)
    return {"window": win, "img_rect": img_rect, "shot": path, "llm": data}


def close_login_page(dry_run: bool = False, tag: str = "closing") -> dict:
    """关闭登录/错误窗：先让 LLM 找关闭按钮，点不到就杀进程重启。"""
    out: dict = {}
    probe = probe_state(tag)
    data = probe["llm"]
    act = data.get("action") if isinstance(data.get("action"), dict) else {}
    target = str(act.get("target") or "")
    out["llm_state"] = data.get("state")
    out["shot"] = probe["shot"]

    if act.get("type") in ("click", "double_click") and any(
        k in target for k in ("关闭", "close", "Close", "×", "x")
    ):
        pt = action_point(data, probe["img_rect"])
        if pt and not dry_run:
            click_screen(pt[0], pt[1], double=bool(act.get("double")))
            out["clicked"] = list(pt)
            time.sleep(1.5)
            if find_main_window(1.0) is None and not wechat_top_windows():
                out["closed_by_click"] = True
                return out
            out["click_ineffective"] = True

    if dry_run:
        out["dry_run"] = True
        return out
    out["killed"] = kill_wechat()
    return out


def _launch_and_wait_window(exe: str, dry_run: bool, res: dict, ev) -> bool:
    """启动微信并等窗口出现（全程只用 win32，不碰 UIA）。

    必须在任何 UIA 调用之前完成（§11.30）：启动微信的进程如果已经连过 UIA，
    微信这个实例就永远不会有无障碍树。返回 False 表示失败，res 里已写好 error。
    """
    exe_path = exe or find_wechat_exe()
    if not exe_path:
        res.update(result="error", error="未找到微信可执行文件，请用 --wechat-exe 指定")
        ev("找不到 Weixin.exe")
        return False
    info = launch_wechat(exe_path, dry_run=dry_run)
    ev("启动微信", **info)
    if not info.get("ok"):
        res.update(result="error", error=info.get("error") or "启动失败", exe=exe_path)
        return False
    appear = time.time() + 30
    while time.time() < appear:
        if wechat_top_windows():
            break
        time.sleep(0.5)
    ev("窗口已出现", windows=len(wechat_top_windows()))
    return True


def step1_launch(
    scan_timeout: float = 180.0,
    poll: float = 3.0,
    relogin: int = 1,
    exe: str = "",
    dry_run: bool = False,
) -> dict:
    """步骤1：微信没启动就启动；需要扫码就轮询；进不去就关登录窗重来。"""
    res: dict = {"stage": "1_launch", "events": []}

    def ev(kind: str, **kw) -> None:
        item = {"t": f"{datetime.now():%H:%M:%S}", "event": kind}
        item.update(kw)
        res["events"].append(item)
        log(f"{kind} {kw if kw else ''}")

    # 1.0 进程不在 → 先启动。**这一步必须在任何 UIA 调用之前**：启动微信的进程如果已经连过
    #     UIA，微信这个实例就永远不会建无障碍树（§11.30）
    if not wechat_pids():
        if not _launch_and_wait_window(exe, dry_run, res, ev):
            return res

    # 1.1 已经登录进来了？
    win = find_main_window(1.5)
    if win is not None:
        rect = window_rect(win)
        n = session_count(win)
        res.update(result="already_logged_in", state="main", rect=rect, session_items=n)
        ev("主界面已存在，无需启动", rect=rect, session_items=n)
        if n <= 0:
            # 窗口在但读不到会话：不是"没登录"，是 Qt accessibility gate 字节=0（见 §11.30 更正）。
            # 先热激活（不重启微信），成功就当没事发生继续往下跑。
            wake = wake_a11y()
            ev("热激活 accessibility gate", ok=wake.get("ok"), attempts=wake.get("attempts"))
            res["wake_a11y"] = wake
            win2 = find_main_window(2.5)
            n2 = session_count(win2) if win2 is not None else 0
            if n2 > 0:
                res.update(
                    result="already_logged_in", state="main", rect=window_rect(win2), session_items=n2
                )
                ev("热激活后 mmui 树已可用", session_items=n2)
                return res
            diag = a11y_diagnosis()
            res["a11y"] = diag
            res.update(
                result="a11y_disabled",
                error=f"微信已在运行，但无障碍树不可用：{diag.get('cause')}",
                fix=diag.get("fix"),
            )
            ev("无障碍树不可用", verdict=diag.get("verdict"), cause=diag.get("cause"))
        return res

    # 1.2 进程在但没窗口 → 大概率缩到托盘，先尝试唤出，绝不重复启动（会多开/掉线）
    pids = wechat_pids()
    if pids and not wechat_top_windows(pids):
        ev("进程存在但无可见窗口，尝试唤出", pids=pids)
        res["force_show"] = force_show_main_hwnd()
        win = find_main_window(4.0)
        if win is None:
            res["reveal"] = reveal_via_llm(dry_run=dry_run, tag="step1_reveal")
            win = find_main_window(4.0)
        if win is not None:
            rect = window_rect(win)
            res.update(result="logged_in", state="main", rect=rect, session_items=session_count(win))
            ev("唤出成功", rect=rect)
            return res
        res.update(result="error", error="微信进程在运行但窗口唤不出来（可能异常），请手动检查")
        return res

    # 1.3 进程在、窗口也在（登录窗或主窗）→ 直接进 1.4 轮询

    # 1.4 轮询：扫码 → 手机确认 → 进入主界面
    deadline = time.time() + max(scan_timeout, 0.0)
    attempt = 0
    restarts = 0
    last_state = "unknown"

    while True:
        win = find_main_window(0.8)
        if win is not None:
            rect = window_rect(win)
            n = session_count(win)
            res.update(result="logged_in", state="main", rect=rect, session_items=n, attempts=attempt)
            ev("检测到主界面（登录成功）", rect=rect, session_items=n)
            return res

        # 轮询期间窗口若又缩回托盘，先按 HWND 拉出来，否则截图判定没有意义
        if (
            len(res.get("force_show_in_loop", [])) < 3
            and wechat_pids()
            and not wechat_top_windows()
        ):
            ev("轮询中发现窗口不可见，尝试唤出")
            res.setdefault("force_show_in_loop", []).append(force_show_main_hwnd())
            if find_main_window(2.0) is not None:
                continue

        # 微信窗口被 IDE/浏览器盖住时，整屏截图判出来的 state 是别人的界面（实测连续 15 轮
        # 读到 "Task-UFO微信.md"，白等一轮 scan_timeout 还触发多余的"超时重来"）→ 先把微信抬到最前
        front = biggest_wechat_hwnd()
        if front:
            bring_to_front(front)

        attempt += 1
        probe = probe_state(f"step1_{attempt:02d}")
        data = probe["llm"]
        state = str(data.get("state") or "other").lower()
        last_state = state
        res["last_probe"] = {
            "state": state,
            "evidence": data.get("evidence"),
            "confidence": data.get("confidence"),
            "shot": probe["shot"],
            "llm_raw": data.get("_raw"),
        }
        if "_error" in data:
            ev("视觉判定失败", error=data["_error"], shot=probe["shot"])
        else:
            ev("状态判定", state=state, evidence=data.get("evidence"), shot=probe["shot"])

        # 需要"点一下才往前走"的界面：欢迎页的「进入微信」按钮、other 页里的可点按钮。
        # 只等不点会一直卡在 welcome（冷启动实测踩过）。
        act = data.get("action") if isinstance(data.get("action"), dict) else {}
        clicks = res.setdefault("clicks", [])
        if act.get("type") in ("click", "double_click") and state != "main" and len(clicks) < 8:
            pt = action_point(data, probe["img_rect"])
            if pt is None:
                ev("模型建议点击但坐标不可用", state=state, target=act.get("target"))
            else:
                item = {
                    "t": f"{datetime.now():%H:%M:%S}",
                    "state": state,
                    "target": act.get("target"),
                    "point": list(pt),
                    "shot": probe["shot"],
                }
                clicks.append(item)
                ev("按模型建议点击", state=state, target=act.get("target"), point=list(pt))
                if not dry_run:
                    click_screen(
                        pt[0],
                        pt[1],
                        double=act.get("type") == "double_click" or bool(act.get("double")),
                    )
                    time.sleep(1.5)
                continue  # 点完立刻回头检查主窗是否已经出现

        # 视觉已经是主界面，但 UIA 枚举不到 mmui::MainWindow（Qt 无障碍树没建起来，
        # 见 §11.23）——这时绝不能"超时→杀微信重启"，那会白白打断已登录状态。
        # 先试着用热键把 UIA 树重建；仍不行就按视觉结果认定登录成功，交给步骤2 继续修。
        if state == "main":
            hwnd = biggest_wechat_hwnd()
            info: dict = {"uia_missing": True, "hwnd": hwnd}
            if hwnd:
                info["recover"] = recover_uia_tree(hwnd, tries=1)
                if find_main_window(2.0) is not None:
                    continue  # 树恢复了，回到循环开头走正常分支
                info["rect"] = _hwnd_rect(hwnd)
                # 热键没用 → 热写 Qt accessibility gate 字节（2026-11-05 实测有效，不用重启微信）
                wake = wake_a11y()
                info["wake_a11y"] = wake
                ev("热激活 accessibility gate", ok=wake.get("ok"), ready=wake.get("ready"))
                if find_main_window(2.0) is not None:
                    continue
            diag = a11y_diagnosis()
            res["a11y"] = diag
            res.update(
                result="a11y_disabled",
                state="main",
                uia_missing=True,
                attempts=attempt,
                rect=info.get("rect"),
                recover=info.get("recover"),
                error=f"已进主界面但无障碍树不可用：{diag.get('cause')}",
                fix=diag.get("fix"),
            )
            ev("视觉判定已进主界面，但无障碍树不可用", hwnd=hwnd, verdict=diag.get("verdict"))
            return res

        timed_out = time.time() >= deadline
        need_restart = state in ("error",) or timed_out

        if need_restart:
            if restarts < relogin:
                restarts += 1
                ev(
                    "关闭登录窗后重来",
                    reason="登录页报错" if state == "error" else "等待超时",
                    attempt=restarts,
                )
                res[f"restart_{restarts}"] = close_login_page(dry_run=dry_run, tag=f"closing_{restarts}")
                log(
                    "注意：本进程已经连过 UIA，这样重启出来的微信可能没有无障碍树（§11.30）；"
                    "若接着报 a11y_disabled，请完全退出微信后单独跑 --stage launch"
                )
                exe_path = exe or find_wechat_exe()
                info = launch_wechat(exe_path, dry_run=dry_run)
                ev("重新启动", **info)
                if not info.get("ok"):
                    res.update(result="error", error=info.get("error") or "重启失败")
                    return res
                time.sleep(3.0)
                deadline = time.time() + max(scan_timeout, 0.0)
                continue
            res.update(
                result="timeout" if timed_out else "error",
                state=last_state,
                error="等待扫码登录超时" if timed_out else "登录页持续报错且重试次数用尽",
                attempts=attempt,
            )
            ev("放弃", state=last_state)
            return res

        time.sleep(max(poll, 0.5))


# --------------------------------------------------------------------------
# 步骤 2：确保微信可见、非最小化、没跑出屏幕（不强制主屏）
# --------------------------------------------------------------------------
def reveal_via_llm(
    dry_run: bool = False, tag: str = "reveal", area: list[int] | None = None
) -> dict:
    """让 LLM 看截图找微信窗口/任务栏图标并点击唤出。

    多屏时先只看主屏 —— 全屏拼接图会让模型把坐标指到另一块屏；
    主屏判定 not_found 再退回全屏找一次。
    """
    areas = [area or primary_monitor_rect()]
    if not area:
        full = virtual_screen_rect()
        if full != areas[0]:
            areas.append(full)

    tries: list[dict] = []
    for idx, ar in enumerate(areas):
        img = grab(ar)
        path = save_shot(img, f"{tag}_{idx}_{datetime.now():%H%M%S}")
        data = llm_json(img, REVEAL_PROMPT)
        cur: dict = {
            "area": ar,
            "shot": path,
            "state": data.get("state"),
            "evidence": data.get("evidence"),
            "confidence": data.get("confidence"),
            "llm_raw": data.get("_raw"),
        }
        if "_error" in data:
            cur["error"] = data["_error"]
            tries.append(cur)
            break

        state = str(data.get("state") or "")
        if state == "main_window_visible":
            cur["skip"] = "模型认为窗口已可见，不点击"
            tries.append(cur)
            break

        act = data.get("action") if isinstance(data.get("action"), dict) else {}
        if act.get("type") not in ("click", "double_click"):
            tries.append(cur)
            if state == "not_found":
                continue  # 换全屏再找
            break

        pt = action_point(data, ar)
        if pt is None:
            cur["error"] = "模型给了动作但坐标不可用"
            tries.append(cur)
            continue
        cur["target"] = act.get("target")
        cur["point"] = list(pt)
        if dry_run:
            cur["dry_run"] = True
            tries.append(cur)
            break
        click_screen(
            pt[0], pt[1], double=act.get("type") == "double_click" or bool(act.get("double"))
        )
        cur["clicked"] = True
        tries.append(cur)
        time.sleep(2.0)
        break

    return {"state": tries[-1].get("state") if tries else None, "tries": tries}


def intersects_ratio(rect: list[int], area: list[int]) -> float:
    l, t, r, b = rect
    al, at, ar, ab = area
    ix = max(0, min(r, ar) - max(l, al))
    iy = max(0, min(b, ab) - max(t, at))
    return (ix * iy) / max(1, (r - l) * (b - t))


def onscreen_ratio(rect: list[int]) -> float:
    """窗口有多大比例落在「所有显示器拼起来的虚拟屏」里。

    这才是能不能按窗口坐标点击的硬条件（跨屏没问题，出界才点空）。
    """
    return intersects_ratio(rect, virtual_screen_rect())


def monitor_for(rect: list[int]) -> list[int]:
    """挑与该矩形重叠最多的显示器；都不重叠就用主屏。用于越界拉回。"""
    best: list[int] | None = None
    best_ratio = 0.0
    for m in monitors():
        rr = intersects_ratio(rect, m["rect"])
        if rr > best_ratio:
            best_ratio = rr
            best = [int(v) for v in m["rect"]]
    return best or primary_monitor_rect()


def move_window(hwnd: int, target: list[int]) -> dict:
    """移动到目标矩形。

    实测（2026-11-02）：微信是 Qt 窗口，刚从托盘恢复后 SetWindowPos 虽然立即返回，
    但随后会延迟重排、自行把窗口撑到屏幕边界；所以每次设置后要等稳定再量，
    用「宽高误差」判定（不能只看右边界），不对就再设一次，稳定后第二次是精确的。
    """
    import win32con
    import win32gui

    x, y, w, h = target
    out: dict = {"target": target, "tries": []}
    try:
        for _attempt in range(3):
            win32gui.SetWindowPos(
                hwnd, 0, x, y, w, h, win32con.SWP_NOZORDER | win32con.SWP_SHOWWINDOW
            )
            time.sleep(1.0)  # 等 Qt 重排稳定，太短会把中间态当成成功
            got = list(win32gui.GetWindowRect(hwnd))
            out["tries"].append(got)
            if (
                abs(got[0] - x) <= 8
                and abs(got[1] - y) <= 8
                and abs((got[2] - got[0]) - w) <= 24
                and abs((got[3] - got[1]) - h) <= 24
            ):
                break
        out["ok"] = True
        out["rect"] = list(win32gui.GetWindowRect(hwnd))
        return out
    except Exception as exc:  # noqa: BLE001
        out.update(ok=False, error=f"{type(exc).__name__}: {exc}")
        return out


def recover_uia_tree(hwnd: int, tries: int = 2) -> dict:
    """会话列表 UIA 读不到条目时，用微信自己的显示/隐藏热键把 UI 树重建出来。

    坑（2026-11-02 实测）：从托盘唤出后，主窗 UIA 只剩一个空 `Pane MMUIRenderSubWindowHW`，
    `descendants(List/ListItem)` 全是 0，而界面上会话列表明明在。
    试过无效：抖动窗口尺寸、点空白处、SW_MINIMIZE→SW_RESTORE、force_show、直接包子窗口 HWND。
    **有效**：发微信自带的「显示/隐藏主窗口」热键 Ctrl+Alt+W 两次（隐藏→显示），树立刻恢复
    （实测 session_count 0 → 13，List 0 → 2）。
    """
    from pywinauto.keyboard import send_keys

    out: dict = {"attempts": [], "ok": False, "hotkey": "Ctrl+Alt+W"}
    for i in range(1, max(tries, 1) + 1):
        try:
            bring_to_front(hwnd)
            time.sleep(0.5)
            send_keys("^%w")
            time.sleep(1.8)
            send_keys("^%w")
            time.sleep(2.2)
        except Exception as exc:  # noqa: BLE001
            out["attempts"].append({"attempt": i, "error": f"{type(exc).__name__}: {exc}"})
            continue
        win = find_main_window(3.0)
        n = session_count(win) if win is not None else 0
        out["attempts"].append({"attempt": i, "session_count": n})
        if n > 0:
            out.update(ok=True, rounds=i)
            return out
    return out


def step2_window_ready(dry_run: bool = False) -> dict:
    """步骤2：确保微信窗口可见、没最小化、且完整落在屏幕内（任何一块屏都行）。

    刻意**不搬主屏**：后续步骤(3/4)全部按窗口自身矩形取坐标/截图（grab 是抓全虚拟屏再裁剪），
    窗口在副屏照常能操作；强行搬动反而会重排微信内部布局、多一次 HWND/UIA 重取。
    只有窗口跑出屏幕边界（点不到）时才拉回重叠最多的那块屏。
    """
    res: dict = {"stage": "2_window_ready", "events": []}

    def ev(kind: str, **kw) -> None:
        item = {"t": f"{datetime.now():%H:%M:%S}", "event": kind}
        item.update(kw)
        res["events"].append(item)
        log(f"{kind} {kw if kw else ''}")

    res["primary_work"] = primary_work_area()
    res["monitors"] = monitors()
    res["virtual"] = virtual_screen_rect()

    # 2.1 找到窗口；找不到先按 HWND 硬唤出，仍不行再让 LLM 点任务栏/托盘图标
    win = find_main_window(2.0)
    if win is None:
        ev("UIA 未找到主窗，先按 HWND 强制唤出")
        res["force_show"] = force_show_main_hwnd()
        win = find_main_window(4.0)
        if win is None:
            ev("HWND 唤出无效，交给视觉 LLM 点任务栏/托盘图标")
            res["reveal"] = reveal_via_llm(dry_run=dry_run, tag="step2_reveal")
            win = find_main_window(5.0)
        if win is None:
            diag = a11y_diagnosis()
            res["a11y"] = diag
            ev("唤出失败，给出无障碍诊断", verdict=diag.get("verdict"), cause=diag.get("cause"))
            res.update(
                result="error",
                error=f"微信窗口既不存在也无法被唤出：{diag.get('cause')}",
                fix=diag.get("fix"),
            )
            return res
        ev("唤出成功")

    hwnd = window_hwnd(win)
    res["hwnd"] = hwnd
    res["wechat_windows"] = wechat_top_windows()[:3]
    rect_before = window_rect(win)
    res["rect_before"] = rect_before
    res["onscreen_before"] = round(onscreen_ratio(rect_before), 3)
    res["on_primary_before"] = round(intersects_ratio(rect_before, primary_monitor_rect()), 3)

    # 2.2 最小化 / 不在前台 → 恢复并置前（失败再用 LLM 点任务栏图标）
    import win32gui

    iconic = bool(hwnd) and win32gui.IsIconic(hwnd)
    res["iconic_before"] = iconic
    if hwnd and (iconic or not win32gui.IsWindowVisible(hwnd)):
        ev("窗口最小化，尝试恢复置前", iconic=iconic)
        res["foreground"] = force_foreground(hwnd)
        if win32gui.IsIconic(hwnd):
            ev("恢复无效，改用视觉点击任务栏图标")
            res["reveal"] = reveal_via_llm(dry_run=dry_run, tag="step2_taskbar")

    win = find_main_window(2.0) or win
    hwnd = hwnd or window_hwnd(win)

    # 2.3 只有窗口跑到屏幕外（点不到）才拉回；位置本身不干预
    rect = window_rect(win)
    res["rect"] = rect
    ratio = onscreen_ratio(rect)
    res["onscreen"] = round(ratio, 3)
    res["on_primary"] = round(intersects_ratio(rect, primary_monitor_rect()), 3)
    if ratio < 0.98:
        area = monitor_for(rect)
        w, h = rect[2] - rect[0], rect[3] - rect[1]
        aw, ah = area[2] - area[0], area[3] - area[1]
        # 能整块放进目标屏就保持原尺寸（不动微信内部布局），放不下才缩
        w, h = max(min(w, aw), 700), max(min(h, ah), 480)
        x = area[0] + max(0, (aw - w) // 2)
        y = area[1] + max(0, (ah - h) // 2)
        # move_window 收的是 [x, y, w, h]；日志里给人看的是 [left, top, right, bottom]
        ev("窗口超出屏幕，拉回到重叠最多的那块屏", **{"from": rect, "to": [x, y, x + w, y + h]})
        if dry_run:
            res["move"] = {"dry_run": True, "target": [x, y, w, h]}
        elif hwnd:
            res["move"] = move_window(hwnd, [x, y, w, h])
            win = find_main_window(2.0) or win
            rect = window_rect(win)
        else:
            res["move"] = {"ok": False, "error": "拿不到 HWND，无法移动"}
        ev("移动完成", rect=rect)

    # 2.4 UIA 树可能是空的（从托盘唤出/最小化还原后常见）：会话列表读不到就热键重建
    items = session_count(win)
    res["session_items"] = items
    if items <= 0:
        ev("会话列表 UIA 为空，尝试重建", hwnd=hwnd)
        if dry_run:
            res["uia_recover"] = {"dry_run": True}
        else:
            res["uia_recover"] = recover_uia_tree(hwnd)
            if res["uia_recover"].get("ok"):
                win = find_main_window(2.0) or win
                items = session_count(win)
        res["session_items_after"] = items
        ev("重建结果", session_items=items)

    res["rect_after"] = rect
    res["onscreen_after"] = round(onscreen_ratio(rect), 3)
    res["on_primary_after"] = round(intersects_ratio(rect, primary_monitor_rect()), 3)
    res["result"] = "ok" if res["onscreen_after"] >= 0.98 else "partial"
    if res["result"] != "ok":
        res["error"] = "窗口仍有部分在屏幕外，按窗口坐标点击可能落空"
    elif items <= 0:
        # 别再猜"UIA 服务挂了"：读屏标志 / mmui 主窗 / 条目数一查就知道原因（§11.30）
        diag = a11y_diagnosis()
        res["a11y"] = diag
        res["result"] = "error"
        res["error"] = f"无障碍树不可用：{diag.get('cause')}"
        res["fix"] = diag.get("fix")
        ev("无障碍诊断", verdict=diag.get("verdict"), cause=diag.get("cause"), fix=diag.get("fix"))
    return res


# --------------------------------------------------------------------------
# 步骤 3：回到正常主聊天窗
# --------------------------------------------------------------------------
CHAT_PROMPT = """这是 Windows 微信桌面客户端的截图。请判断当前是否处于「正常的主聊天窗」，
并给出回到主聊天窗应该做什么。只输出一个 JSON 对象，不要解释、不要 Markdown 代码块：

{
  "state": "main_chat|chat_detail|sub_page|dialog|search_panel|other",
  "evidence": "图中关键内容（不超过20字）",
  "action": {"type": "click|double_click|key|none", "key": "esc|alt+left", "target": "要点的东西", "center_norm": [0.0, 0.0]},
  "confidence": 0.0
}

判定规则：
- main_chat：左侧是会话列表（多个会话，头像+昵称+消息摘要），右侧是**某个会话的内容区**，
  没有任何浮层遮挡 → 这就是正常主聊天窗。
  特别注意：右侧内容可以是普通聊天，也可以是公众号/服务号的图文消息流（图文卡片、
  「看一看」、图片网格、视频号卡片），只要左侧有选中的会话项、右侧是它的消息内容，就算 main_chat
- chat_detail：聊天窗口里打开了右侧的详情/成员/设置侧栏，或有群公告等覆盖层
- sub_page：左侧会话列表消失或被覆盖，整页变成通讯录、朋友圈、收藏、设置页、小程序，
  或单篇文章详情页（顶部有返回箭头、正文是长文章）等非会话页面
- dialog：弹出了确认框/提示框/更新提示/退出确认等（通常带"确定""取消""×"按钮）
- search_panel：顶部搜索框展开了下拉结果面板
- blank：界面几乎是一整片空白（白屏），只剩最左侧一列导航图标，既没有会话列表也没有聊天内容
- other：判断不出

动作规则（程序会照做）：
- main_chat → action.type 用 "none"，不要点击
- dialog → 必须给 click，target 指向「取消 / 关闭 / × / 稍后」这类安全按钮，并给 center_norm
- blank → 必须给 click，target 指向最左侧最上方的「聊天」图标（对话气泡），center_norm 指向它
- sub_page / chat_detail / search_panel → 优先 action.type="key"、key="esc"（Esc 通常能返回或关闭）；
  只有 Esc 明显不适用时，才给 click 指向返回箭头或关闭按钮
- 严禁点击「退出登录」「删除」「清空聊天记录」「立即更新」等破坏性或不可逆按钮
- 严禁点击窗口右上角的最小化/最大化/关闭三个系统按钮（那个 × 会把微信关到托盘）

坐标要求：center_norm 是相对整张图宽高的 0~1 比例。
"""

# 兜底动作序列：模型不给动作、或同一建议重复无效时按顺序升级（每项只试一次）。
# 注意不要裸发 ESC：实测微信主窗按 Esc 会直接把窗口关到托盘（通讯录页踩过）。
_FALLBACK_PLANS = ("nav_chat", "key:%{LEFT}", "conv_top", "nav_chat")


def _is_system_button_zone(pt: tuple[int, int] | list[int], rect: list[int]) -> bool:
    """窗口右上角是系统的最小化/最大化/关闭按钮，绝不让模型点那里。"""
    l, t, r, b = [int(v) for v in rect]
    x, y = int(pt[0]), int(pt[1])
    return (r - x) <= 90 and (y - t) <= 70


def _in_left_nav_zone(pt: tuple[int, int] | list[int], rect: list[int]) -> bool:
    """点是否落在左侧竖排导航条上。

    模型给这一条的坐标经常偏 20~40px（实测给 211，控件中心在 244），
    所以这里一律转交给 UIA 控件点击，不用模型坐标。
    """
    l, t, r, b = [int(v) for v in rect]
    x, y = int(pt[0]), int(pt[1])
    return (x - l) < 75 and 30 <= (y - t) <= 0.5 * (b - t)


def _chat_view_rect(win=None) -> list[int]:
    """判定用截图区域：主窗 rect + 微信其它较大的可见窗口（对话框可能比主窗更大）的并集。"""
    found = find_main_hwnd()
    if found is not None:
        base = found["rect"]
    elif win is not None:
        base = window_rect(win)
    else:
        return virtual_screen_rect()
    area_base = max(1, (base[2] - base[0]) * (base[3] - base[1]))
    rects = [base]
    for w in wechat_top_windows():
        r = w["rect"]
        if max(1, (r[2] - r[0]) * (r[3] - r[1])) >= 0.15 * area_base:
            rects.append(r)
    l = min(r[0] for r in rects)
    t = min(r[1] for r in rects)
    r2 = max(r[2] for r in rects)
    b = max(r[3] for r in rects)
    vl, vt, vr, vb = virtual_screen_rect()
    return [max(l, vl), max(t, vt), min(r2, vr), min(b, vb)]


def _send_hotkey(key: str) -> str:
    from pywinauto.keyboard import send_keys

    mapping = {
        "esc": "{ESC}",
        "escape": "{ESC}",
        "alt+left": "%{LEFT}",
        "back": "%{LEFT}",
        "backspace": "{BACKSPACE}",
        "enter": "{ENTER}",
    }
    seq = mapping.get((key or "esc").strip().lower(), "{ESC}")
    send_keys(seq)
    return seq


def _left_nav_chat_point(win) -> tuple[int, int]:
    """左侧竖排导航「微信（聊天）」按钮的中心坐标（坐标降级用）。

    实测窗口内偏移 (30, 118)；之前的 0.13 高度比例算出来是「通讯录」，点错了页面。
    优先用 click_left_nav() 走 UIA 控件，这里是最后的坐标兜底。
    """
    l, t, _r, _b = window_rect(win)
    return l + 30, t + 118


def click_left_nav(win, name: str = "微信") -> dict:
    """点左侧竖排导航（微信/通讯录/收藏/发现）。

    微信 4.x 把这些按钮以 Button 控件暴露（`descendants(control_type="Button")`
    服务端过滤，秒回），所以优先用控件的 invoke/click_input，
    比"按窗口比例猜坐标"可靠得多；两者都失败才用固定偏移兜底。
    """
    try:
        for b in win.descendants(control_type="Button"):
            if (b.window_text() or "").strip() != name:
                continue
            r = b.rectangle()
            rect = [int(r.left), int(r.top), int(r.right), int(r.bottom)]
            # Qt 不实现 UIA 的 Invoke pattern，b.invoke() 会静默无效，
            # 必须发真实鼠标事件（click_input 就是点控件矩形中心）
            b.click_input()
            return {"ok": True, "name": name, "rect": rect, "via": "uia_click"}
    except Exception as exc:  # noqa: BLE001
        err = f"{type(exc).__name__}: {exc}"
        pt = _left_nav_chat_point(win)
        click_screen(pt[0], pt[1])
        return {"ok": True, "name": name, "point": list(pt), "via": "coordinate", "uia_error": err}

    pt = _left_nav_chat_point(win)
    click_screen(pt[0], pt[1])
    return {"ok": True, "name": name, "point": list(pt), "via": "coordinate", "note": "未找到同名按钮"}


def _first_conversation_point(win) -> tuple[int, int]:
    """会话列表第一项的中心（窗口内固定偏移，实测命中）。兜底用。"""
    l, t, _r, _b = window_rect(win)
    return l + 150, t + 110


def close_side_windows(main_hwnd: int, dry_run: bool = False) -> list[dict]:
    """关闭除主窗以外的微信顶层窗口（设置、聊天文件、聊天记录迁移与备份、文章详情等）。

    这类窗口都有真实 HWND，用 WM_CLOSE 关远比"截图让模型找 ×"可靠：
    实测模型只能看到画面最前面的那个，被主窗挡住或在后面叠着的附属窗口它发现不了。
    """
    import win32con
    import win32gui

    found = find_main_hwnd()
    main_cls = found["class"] if found else "Qt51514QWindowIcon"
    main_area = 1
    if found:
        ml, mt, mr, mb = found["rect"]
        main_area = max(1, (mr - ml) * (mb - mt))

    out: list[dict] = []
    for w in wechat_top_windows():
        if int(w["hwnd"]) == int(main_hwnd):
            continue
        # 只关和主窗同类的真实功能窗口；Qt 的 Tool/ToolTip/SaveBits 浮层、
        # 托盘消息窗、IME 辅助窗一律不碰（曾误关标题为 Weixin 的 ToolSaveBits 浮层）
        if (w.get("class") or "") != main_cls:
            continue
        l, t, r, b = w["rect"]
        if (r - l) * (b - t) < 0.2 * main_area:
            continue
        item: dict = {"hwnd": w["hwnd"], "title": w["title"], "rect": w["rect"]}
        if dry_run:
            item["dry_run"] = True
        else:
            try:
                win32gui.PostMessage(w["hwnd"], win32con.WM_CLOSE, 0, 0)
                item["closed"] = True
            except Exception as exc:  # noqa: BLE001
                item["error"] = f"{type(exc).__name__}: {exc}"
        out.append(item)
    if out and not dry_run:
        time.sleep(1.2)
    return out


def step3_main_chat(max_rounds: int = 4, dry_run: bool = False) -> dict:
    """确认在主聊天窗；不在（子页面/弹层/对话框）就按 LLM 建议 + 兜底按键退回来。"""
    res: dict = {"stage": "3_main_chat", "events": []}

    def ev(kind: str, **kw) -> None:
        item = {"t": f"{datetime.now():%H:%M:%S}", "event": kind}
        item.update(kw)
        res["events"].append(item)
        log(f"{kind} {kw if kw else ''}")

    win = find_main_window(3.0)
    found = find_main_hwnd()
    if found is None:
        res.update(result="error", error="未找到微信主窗，请先跑 --stage launch / --stage screen")
        return res
    if not found["visible"] or found["iconic"]:
        res["force_show"] = force_show_main_hwnd()
        win = find_main_window(2.0) or win
    # 被别的窗口挡住时，截图截到的是别人 → 必须先真正置前
    res["bring_front"] = bring_to_front(found["hwnd"])
    time.sleep(0.4)

    res["windows_before"] = wechat_top_windows()
    res["session_items_before"] = session_count(win)

    # 3.1 先确定性地关掉附属窗口：设置 / 聊天文件 / 聊天记录迁移与备份 / 文章详情等。
    #     （模型看截图只认得最前面那个，叠在后面或被主窗挡住的它发现不了）
    side = close_side_windows(found["hwnd"], dry_run=dry_run)
    if side:
        res["side_windows"] = side
        ev("关闭附属窗口", count=len(side), titles=[s["title"] for s in side])
        win = find_main_window(2.0) or win

    last_sig: tuple[str, str] | None = None

    for rnd in range(1, max_rounds + 1):
        # 每轮先体检：微信一旦无响应立刻停手，绝不继续点/继续问 UIA
        if wechat_hung():
            res.update(
                result="error",
                error="微信窗口无响应（IsHungAppWindow=True），已中止以免继续恶化",
                rounds_used=rnd - 1,
            )
            ev("微信无响应，中止")
            return res

        win = find_main_window(2.0) or win
        if not dry_run:
            # 每轮都重新置前：上一轮的点击/按键可能把焦点送给了别的程序
            bring_to_front(found["hwnd"])
        rect = _chat_view_rect(win)
        img = grab(rect)
        path = save_shot(img, f"step3_{rnd:02d}_{datetime.now():%H%M%S}")
        data = llm_json(img, CHAT_PROMPT)

        state = str(data.get("state") or "other").lower()
        act = data.get("action") if isinstance(data.get("action"), dict) else {}
        atype = str(act.get("type") or "none").lower()
        item: dict = {
            "round": rnd,
            "state": state,
            "evidence": data.get("evidence"),
            "confidence": data.get("confidence"),
            "action": {"type": atype, "key": act.get("key"), "target": act.get("target")},
            "shot": path,
            "session_items": session_count(win),
        }
        res.setdefault("rounds", []).append(item)
        ev("界面判定", round=rnd, state=state, evidence=data.get("evidence"), action=atype)

        if "_error" in data:
            item["error"] = data["_error"]

        if state == "main_chat" and atype in ("none", ""):
            res.update(
                result="ok",
                rounds_used=rnd,
                rect=rect,
                windows_after=wechat_top_windows(),
                session_items_after=session_count(win),
            )
            ev("已在主聊天窗")
            return res

        # ---- 决定这一轮做什么 ----
        # 危险区按「主窗」判断（合并截图区域的右上角不代表主窗的系统按钮）
        main_rect = (find_main_hwnd() or {}).get("rect") or rect
        plan: dict | None = None
        if atype in ("click", "double_click"):
            pt = action_point(data, rect)
            if pt and _is_system_button_zone(pt, main_rect):
                item["note"] = "模型指向窗口系统按钮区，已拒绝该坐标"
            elif pt and _in_left_nav_zone(pt, main_rect):
                # 左侧导航条的点击交给 UIA 控件，模型给的坐标偏得太多
                item["note"] = "模型点在左侧导航条，改用 UIA 控件点击「微信」"
                plan = {"type": "nav", "name": "微信", "reason": "left_nav_from_llm"}
            elif pt:
                plan = {"type": atype, "point": list(pt), "target": act.get("target")}
            else:
                item["note"] = "模型给了点击但坐标不可用"
        elif atype == "key":
            key = str(act.get("key") or "esc").strip().lower()
            if key in ("esc", "escape") and state != "search_panel":
                # 微信主窗里按 Esc 会直接关到托盘（实测），只有搜索面板内才安全
                item["note"] = "ESC 在该状态会关闭窗口，改用左侧「微信」导航"
                plan = {"type": "nav", "name": "微信", "reason": "esc_unsafe"}
            else:
                plan = {"type": "key", "key": key}

        # 白屏：点左侧「聊天」图标最有效，不必先浪费一轮 ESC
        if plan is None and state == "blank":
            plan = {
                "type": "click",
                "point": list(_left_nav_chat_point(win)),
                "reason": "blank_nav_chat",
            }

        sig = (state, atype)
        repeated = sig == last_sig
        if plan is None or repeated:
            # 模型没给动作，或同一状态下同样的建议已经试过没效果 → 从兜底候选中挑「还没试过的」
            candidates: list[dict] = []
            for step in _FALLBACK_PLANS:
                if step == "nav_chat":
                    candidates.append(
                        {"type": "nav", "name": "微信", "reason": "fallback_nav_chat"}
                    )
                elif step == "conv_top":
                    candidates.append(
                        {
                            "type": "click",
                            "point": list(_first_conversation_point(win)),
                            "reason": "fallback_conv_top",
                        }
                    )
                elif ":" in step:
                    candidates.append(
                        {
                            "type": "raw_key",
                            "seq": step.split(":", 1)[1],
                            "reason": "fallback",
                        }
                    )
            tried = res.setdefault("tried_plans", [])
            plan = candidates[0]
            for cand in candidates:
                key = json.dumps(cand, sort_keys=True, ensure_ascii=False)
                if key not in tried:
                    plan = cand
                    tried.append(key)
                    break
        last_sig = sig

        item["plan"] = plan
        ev("执行动作", round=rnd, plan=plan)

        if dry_run:
            res.update(result="dry_run", rounds_used=rnd, rect=rect)
            return res

        try:
            if plan["type"] == "nav":
                res.setdefault("nav_clicks", []).append(
                    click_left_nav(win, plan.get("name") or "微信")
                )
            elif plan["type"] == "key":
                _send_hotkey(str(plan.get("key")))
            elif plan["type"] == "raw_key":
                from pywinauto.keyboard import send_keys

                send_keys(str(plan["seq"]))
            else:
                click_screen(
                    plan["point"][0],
                    plan["point"][1],
                    double=plan["type"] == "double_click",
                )
        except Exception as exc:  # noqa: BLE001
            item["exec_error"] = f"{type(exc).__name__}: {exc}"
            ev("动作执行失败", error=item["exec_error"])
        time.sleep(1.5)

        # 动作后体检：万一误点把窗口关到托盘/最小化，立刻拉回来再继续
        after = find_main_hwnd()
        if after is None or not after["visible"] or after["iconic"]:
            ev("动作后窗口不可见，强制唤回", round=rnd)
            res.setdefault("recover", []).append(force_show_main_hwnd())
            win = find_main_window(2.0) or win

    # 兜底用完还没回去：最后点一次左侧「微信」导航
    win = find_main_window(2.0) or win
    if not dry_run:
        nav = click_left_nav(win, "微信")
        ev("兜底点击左侧「微信」导航", **nav)
        res.setdefault("nav_clicks", []).append(nav)
        time.sleep(1.5)
    res.setdefault("rounds", []).append({"round": "final", "plan": {"type": "nav", "name": "微信"}})

    win = find_main_window(2.0) or win
    rect = _chat_view_rect(win)
    img = grab(rect)
    path = save_shot(img, f"step3_final_{datetime.now():%H%M%S}")
    data = llm_json(img, CHAT_PROMPT)
    state = str(data.get("state") or "other").lower()
    res.setdefault("rounds", []).append(
        {"round": "verify", "state": state, "evidence": data.get("evidence"), "shot": path}
    )
    ev("最终校验", state=state, shot=path)

    res.update(
        result="ok" if state == "main_chat" else "partial",
        rounds_used=max_rounds + 1,
        rect=rect,
        windows_after=wechat_top_windows(),
        session_items_after=session_count(win),
    )
    if res["result"] != "ok":
        res["error"] = f"未能回到主聊天窗，最后判定 state={state}"
    return res


# --------------------------------------------------------------------------
# 步骤 4：在「公众号」会话里搜索指定公众号并取回结果
# --------------------------------------------------------------------------
RESULT_PROMPT = """你在看微信「搜一搜」的**搜索结果独立窗口**（窗口标题形如「xxx - 搜一搜」）。

只输出一个 JSON 对象，不要解释、不要 Markdown：
{
  "keyword": "顶部搜索框里的关键词",
  "state": "results|suggest|loading|empty|other",
  "active_tab": "当前选中的分类名（全部/账号/文章/视频…）",
  "tabs": [{"name": "分类名", "center_norm": [0.0, 0.0]}],
  "results": [{"name": "条目标题", "kind": "账号|文章|视频|其他", "desc": "一句话说明", "center_norm": [0.0, 0.0]}]
}

要点：
- tabs 是搜索框下面那一行分类（AI搜索/全部/账号/视频/划线/文章/直播…），按从左到右顺序给全；
- results 只列**左侧结果区**里的条目，绝对不要列右侧「相关搜索」的关键词；
- 公众号账号的条目形如「建约车评 - 账号」，name 取「建约车评」，kind 给「账号」；
- center_norm 是相对整张图 0~1 的比例坐标，指向该分类/条目可点击区域的中心；
- 看不清就给 null，不要编造。
"""


def _session_list(win):
    try:
        for lst in win.descendants(control_type="List"):
            if lst.element_info.automation_id == "session_list":
                return lst
    except Exception:  # noqa: BLE001
        return None
    return None


def find_session_item(win, name: str) -> tuple[object | None, list[str]]:
    """在会话列表里按会话名找条目（名字取每项第一行，如「公众号」）。"""
    lst = _session_list(win)
    if lst is None:
        return None, []
    names: list[str] = []
    try:
        for it in lst.children(control_type="ListItem"):
            text = (it.window_text() or "").replace("\r", "").strip()
            first = text.split("\n")[0].strip()
            names.append(first)
            if first == name or first.startswith(name):
                return it, names
    except Exception:  # noqa: BLE001
        return None, names
    return None, names


def find_session_item_deep(win, name: str, rounds: int = 7) -> tuple[object | None, list[str]]:
    """找会话条目；不在可视区就滚动会话列表继续找（先滚到顶，再往下翻）。

    坑（2026-11-02 实测）：会话多的时候「公众号」会被挤出可视区，UIA 只暴露当前可见的
    ListItem，于是步骤4 直接报"会话列表里没有「公众号」"并停摆——必须自己滚列表找。
    """
    item, names = find_session_item(win, name)
    if item is not None:
        return item, names
    lst = _session_list(win)
    if lst is None:
        return None, names
    try:
        r = lst.rectangle()
    except Exception:  # noqa: BLE001
        return None, names
    rect = [int(r.left), int(r.top), int(r.right), int(r.bottom)]
    if rect[2] - rect[0] < 40 or rect[3] - rect[1] < 40:
        return None, names
    seen = list(names)
    plan = [-3, -3, -3, 3, 3, 3, 3][: max(rounds, 1)]
    for notches in plan:
        # gate 每 13~25 秒被微信复位（见 `_GATE_WAKE` 注释），这个循环要 10~20 秒，
        # 中途复位会让 UIA 只剩空壳、后面几轮全读不到条目 → 每轮开头保活一次。
        maybe_wake_gate("步骤4 滚会话列表")
        scroll_page(rect, notches)
        item, names = find_session_item(win, name)
        seen += [n for n in names if n not in seen]
        if item is not None:
            return item, seen
    return None, seen


def click_top_button(win, label: str) -> dict:
    """点窗口顶部那一排的同名按钮（如「搜索」放大镜）。"""
    t = window_rect(win)[1]
    for btn in win.descendants(control_type="Button"):
        try:
            if (btn.window_text() or "").strip() != label:
                continue
            br = btn.rectangle()
            rect = [int(br.left), int(br.top), int(br.right), int(br.bottom)]
            if rect[1] - t > 130:  # 只认窗口顶部区域，别点到别处的同名按钮
                continue
            btn.click_input()
            return {"ok": True, "label": label, "rect": rect}
        except Exception:  # noqa: BLE001
            continue
    return {"ok": False, "error": f"窗口顶部没有找到「{label}」按钮"}


def type_text(text: str, submit: bool = False) -> dict:
    """往当前焦点控件输入文字，可选回车提交。

    中文必须走剪贴板：send_keys 走 IME 会吞字，set_edit_text 不产生真实输入事件
    （微信收不到，搜索下拉不弹）。
    实测：微信搜索框粘完必须再按回车才会真正出结果列表，只输入只会停在建议下拉。
    """
    import win32clipboard
    from pywinauto.keyboard import send_keys

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()
    send_keys("^a")
    time.sleep(0.2)
    send_keys("^v")
    time.sleep(1.2)
    res = {"ok": True, "text": text, "via": "clipboard"}
    if submit:
        send_keys("{ENTER}")
        res["submit"] = True
        time.sleep(2.0)  # 等结果列表加载
    return res


PAGE_BUTTONS = ("搜索", "通知", "个人中心", "独立窗口显示")


def gzh_page_buttons(win) -> list[dict]:
    """主窗右上角那组按钮（搜索/通知/个人中心/独立窗口显示）。

    实测：这组按钮**只在右侧真的打开了某个会话页时**才存在；
    主聊天窗空白（没选中会话）时只有 置顶/最小化/最大化/关闭。
    所以它是「公众号页面到底有没有打开」的确定性判据。
    """
    out: list[dict] = []
    try:
        top = window_rect(win)[1]
        for b in win.descendants(control_type="Button"):
            t = (b.window_text() or "").strip()
            if t not in PAGE_BUTTONS:
                continue
            r = b.rectangle()
            rect = [int(r.left), int(r.top), int(r.right), int(r.bottom)]
            if rect[1] - top > 130:  # 只看顶部那一排
                continue
            out.append({"t": t, "rect": rect})
    except Exception:  # noqa: BLE001
        return out
    return out


def gzh_page_open(win) -> dict:
    hits = gzh_page_buttons(win)
    return {"ok": len(hits) >= 2, "buttons": [h["t"] for h in hits], "hits": hits}


def ensure_gzh_page(win, max_clicks: int = 3, dry_run: bool = False) -> dict:
    """保证右侧真的打开「公众号」页面。

    坑（实测）：点会话列表里的「公众号」有时会把它点"收起"，右侧变空白、
    右上角那组按钮消失，此时去搜索只会搜到聊天记录里的人和会话，不是公众号。
    所以：已开就跳过（避免重复点把页面点没了）；没开就点一次 + 校验，最多 N 次。
    """
    out: dict = {"clicks": []}
    before = gzh_page_open(win)
    out["already_open"] = before["ok"]
    out["buttons_before"] = before["buttons"]
    if before["ok"]:
        out.update(ok=True, buttons=before["buttons"])
        return out
    if dry_run:
        out.update(ok=False, dry_run=True, error="dry_run：公众号页面没打开（未点击）")
        return out

    for i in range(1, max_clicks + 1):
        win = find_main_window(2.0) or win
        item, names = find_session_item(win, "公众号")
        if item is None:
            out.update(ok=False, error="会话列表里没有「公众号」会话", session_names=names[:10])
            return out
        try:
            item.scroll_into_view()
            time.sleep(0.3)
        except Exception:  # noqa: BLE001
            pass
        try:
            r = item.rectangle()
            rect = [int(r.left), int(r.top), int(r.right), int(r.bottom)]
        except Exception:  # noqa: BLE001
            rect = []
        via = "uia_click"
        try:
            item.click_input()
        except Exception as exc:  # noqa: BLE001
            via = f"coord({type(exc).__name__})"
            if rect:
                click_screen((rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2)
        time.sleep(1.8)
        chk = gzh_page_open(win)
        out["clicks"].append(
            {"attempt": i, "rect": rect, "via": via, "buttons": chk["buttons"], "ok": chk["ok"]}
        )
        log(f"点「公众号」会话 第{i}次 → {'页面已打开' if chk['ok'] else '仍没打开'} {chk['buttons']}")
        if chk["ok"]:
            out.update(ok=True, buttons=chk["buttons"])
            return out

    out.update(
        ok=False,
        error=f"点了 {max_clicks} 次「公众号」会话，右上角始终没有 搜索/通知 这组按钮（页面没打开）",
    )
    return out


def find_search_window(timeout: float = 6.0, main_hwnd: int = 0) -> dict | None:
    """找「搜一搜」的独立结果窗口。

    实测：回车后结果不在主窗里，而是另开一个 CEF 顶层窗口（进程 `WeChatAppEx.exe`）。
    该窗口在微信失焦后会自动关闭，必须在同一次运行里操作。

    坑（2026-11-02 真踩）：不要用「窗口类 + 标题含微信」去筛。VS Code / 豆包的 Electron
    窗口 class 同样是 `Chrome_WidgetWin_1`，当时 IDE 打开的是标题带"微信UI自动化技能"的文档，
    于是把 IDE 当成了搜索窗口（还因为面积更大排第一），视觉模型一路读出 page=other /
    title=CodeBuddy，白跑 4 轮。唯一可靠判据是**进程名 `WeChatAppEx.exe`**（见 `cef_windows()`）。
    """
    deadline = time.time() + max(timeout, 0)
    while True:
        cands = [
            w
            for w in cef_windows()
            if int(w["hwnd"]) != int(main_hwnd or 0)
            and (w["rect"][2] - w["rect"][0]) >= 420
            and (w["rect"][3] - w["rect"][1]) >= 320
        ]
        if cands:
            cands.sort(
                key=lambda c: (
                    0 if c["class"] == "Chrome_WidgetWin_0" else 1,
                    -((c["rect"][2] - c["rect"][0]) * (c["rect"][3] - c["rect"][1])),
                )
            )
            return cands[0]
        if time.time() >= deadline:
            return None
        time.sleep(0.4)


def page_search_box(win) -> dict | None:
    """公众号页面顶部的内嵌搜索框。

    实测：点右上角放大镜后，主窗里会多出一个 Edit，位置在右上角
    （例如 rect=[965,173,1125,196]），左边会话列表里那个「搜索」框是另一个。
    这里只认「窗口顶部 130px 内 + 位于窗口右半边」的 Edit。
    """
    try:
        rect = window_rect(win)
        mid = rect[0] + (rect[2] - rect[0]) * 0.5
        for e in win.descendants(control_type="Edit"):
            r = e.rectangle()
            er = [int(r.left), int(r.top), int(r.right), int(r.bottom)]
            if er[1] - rect[1] < 130 and er[0] > mid:
                return {"edit": e, "rect": er, "value": (e.window_text() or "").strip()}
    except Exception:  # noqa: BLE001
        return None
    return None


def search_in_gzh_page(win, keyword: str, dry_run: bool = False, retries: int = 2) -> dict:
    """在「公众号」页面顶部搜索：点放大镜 → 等内嵌搜索框 → 点进去粘贴 → 校验。

    注意方向：不是左侧会话列表的全局搜索框，而是公众号页面右上角那个放大镜。
    """
    from pywinauto.keyboard import send_keys

    out: dict = {"retries": []}
    if dry_run:
        out.update(ok=False, dry_run=True)
        return out

    for i in range(1, max(retries, 1) + 1):
        if page_search_box(win) is None:
            btn = click_top_button(win, "搜索")
            out["search_button"] = btn
            if not btn.get("ok"):
                out.update(ok=False, error=btn.get("error") or "没点到页面右上角搜索按钮")
                return out
            time.sleep(1.3)

        box = page_search_box(win)
        if box is None:
            out["retries"].append({"attempt": i, "error": "点放大镜后没出现内嵌搜索框"})
            log(f"公众号页搜索框 第{i}次 → 没出现搜索框")
            continue

        click_screen((box["rect"][0] + box["rect"][2]) // 2, (box["rect"][1] + box["rect"][3]) // 2)
        time.sleep(0.5)
        typed = type_text(keyword, submit=False)
        time.sleep(0.9)
        value = (page_search_box(win) or {}).get("value", "")
        ok = keyword in value
        out["retries"].append(
            {"attempt": i, "rect": box["rect"], "value": value, "ok": ok, "via": typed.get("via")}
        )
        log(f"公众号页搜索框 第{i}次 → 值='{value}' {'OK' if ok else '没收到'}")
        if ok:
            out.update(ok=True, value=value, rect=box["rect"])
            return out
        send_keys("^a")
        send_keys("{DELETE}")
        time.sleep(0.4)

    out.update(ok=False, error="公众号页面顶部搜索框始终没收到关键词")
    return out


WECHAT_CEF_PROCS = ("wechatappex.exe",)


def _proc_name(pid: int) -> str:
    import psutil

    try:
        return (psutil.Process(pid).name() or "").lower()
    except Exception:  # noqa: BLE001
        return ""


def cef_windows(title_filter: bool = True) -> list[dict]:
    """微信的 CEF 页面窗口（搜一搜 / 独立公众号页都开在这里）。

    实测（2026-11-02）：这类窗口 class=`Chrome_WidgetWin_0`、title=`微信`，
    但**进程是 `WeChatAppEx.exe`**，不是 `Weixin.exe`——所以按 wechat_pids() 过滤会漏。
    判据必须是「进程名 + 窗口类 + 标题」三者一起，否则会误伤编辑器/浏览器的 Chromium 窗口。

    `title_filter=False` 时不看标题：**文章详情窗口的标题是文章标题**（不带"微信"），
    步骤 6 要能认出它，所以另给一个不做标题过滤的入口。
    """
    import win32gui
    import win32process

    out: list[dict] = []

    def _cb(hwnd, _p):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            cls = win32gui.GetClassName(hwnd)
            if not cls.startswith("Chrome_WidgetWin"):
                return True
            title = win32gui.GetWindowText(hwnd) or ""
            if title_filter and "微信" not in title and "搜一搜" not in title:
                return True
            _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
            if _proc_name(int(pid)) not in WECHAT_CEF_PROCS:
                return True
            l, t, r, b = win32gui.GetWindowRect(hwnd)
            if (r - l) < 300 or (b - t) < 200:
                return True
            out.append(
                {
                    "hwnd": int(hwnd),
                    "class": cls,
                    "title": title[:20],
                    "proc": "WeChatAppEx.exe",
                    "rect": [int(l), int(t), int(r), int(b)],
                }
            )
        except Exception:  # noqa: BLE001
            pass
        return True

    win32gui.EnumWindows(_cb, None)
    return out


def close_residue_windows(main_hwnd: int, dry_run: bool = False) -> dict:
    """关掉历史残留的「搜一搜 / 独立公众号页」窗口，让搜索从干净状态开始。

    为什么需要：实测这类 CEF 窗口会**跨多次运行存活**（同一 hwnd 3999314 从 13:25 活到 16:03），
    下次搜索直接复用它，页面停在上次的 tab/公众号主页上，视觉模型读到的"当前状态"就带历史残留，
    跳数也随之漂移。清理后每次都是新窗口，状态可预期。
    """
    import win32con
    import win32gui

    out: dict = {"cef": [], "qt": [], "closed": 0, "dry_run": dry_run}

    for w in cef_windows():
        rec = {k: w[k] for k in ("hwnd", "class", "title", "rect")}
        rec["proc"] = "WeChatAppEx.exe"
        if dry_run:
            rec["dry_run"] = True
        else:
            try:
                win32gui.PostMessage(w["hwnd"], win32con.WM_CLOSE, 0, 0)
                rec["closed"] = True
            except Exception as exc:  # noqa: BLE001
                rec["error"] = f"{type(exc).__name__}: {exc}"
        out["cef"].append(rec)
    if out["cef"] and not dry_run:
        time.sleep(1.5)

    # 主窗旁的 Qt 附属窗口（设置/聊天文件/独立窗口显示的公众号页）复用步骤3 的白名单关窗
    sides = close_side_windows(main_hwnd, dry_run=dry_run)
    out["qt"] = [{k: s.get(k) for k in ("hwnd", "title", "closed", "dry_run", "error")} for s in sides]
    if sides and not dry_run:
        time.sleep(1.0)

    out["closed"] = len([c for c in out["cef"] if c.get("closed")]) + len(
        [s for s in out["qt"] if s.get("closed")]
    )
    return out


def norm_point(rect: list[int], norm) -> tuple[int, int] | None:
    """0~1 比例坐标 → 屏幕绝对坐标。

    action_point 要求 action.type 是 click，这里包一层，避免各处漏写 type 导致
    静默返回 None（踩过：结果全丢，日志还一片正常）。
    """
    return action_point({"action": {"type": "click", "center_norm": norm}}, rect)


def read_search_window(rect: list[int], tag: str = "search") -> dict:
    """截图搜索窗口 → 让模型结构化读：当前分类、分类行、结果条目。

    截图必须**裁到内容列**再喂模型（`content_crop`）：搜一搜窗口在副屏最大化时宽 1928px，
    两侧各一大片灰底留白，整窗喂进去模型给的 `center_norm` 系统性偏右/偏下
    （实测条目 y 偏 ~110px、x 偏 ~290px，点下去落到筛选行或右边的「上次使用」按钮上）。
    裁成 748px 的白色内容列后误差降到 ~50px，再叠加卡片带吸附即可稳点。
    """
    full = grab(rect)
    img, box = content_crop(full)
    crect = [rect[0] + box[0], rect[1] + box[1], rect[0] + box[2], rect[1] + box[3]]
    shot = save_shot(img, f"step4_{tag}_{datetime.now():%H%M%S}")
    data = llm_json(img, RESULT_PROMPT)
    data["_shot"] = shot
    data["_rect"] = crect
    data["_img"] = img
    return data


OPEN_PROMPT = """你在看微信「搜一搜」窗口里的页面。判断当前是哪种页面，只输出 JSON：
{
  "page": "search_results|account_list|account_home|article|other",
  "title": "页面上最显眼的主体名称（公众号名/文章标题）",
  "evidence": "判断依据，20 字内",
  "accounts": [{"name": "账号名", "center_norm": [0.0, 0.0]}],
  "articles": [{"title": "文章标题", "center_norm": [0.0, 0.0]}]
}

页面判据：
- search_results：顶部有搜索输入框 + 一排分类（全部/账号/视频/文章…）；
- account_list：顶部有「不限/公众号/服务号」筛选，下面是账号条目列表（头像+名称+公众号+简介+原创篇数）；
- account_home：公众号主页（大号公众号名 + 关注/私信按钮，下面是文章列表）；
- article：单篇图文正文。
center_norm 是相对整张图 0~1 的比例坐标；拿不到给 null。不要输出 JSON 之外的内容。
"""

ACCOUNT_KINDS = ("账号", "公众号", "服务号")


def classify_search_page(rect: list[int], tag: str = "page") -> dict:
    """给搜一搜窗口的当前页面归类（搜索结果/账号列表/公众号主页/文章）。

    同 `read_search_window`：截图先裁内容列，返回的 `rect`/`img` 都是裁剪后的，
    后续点击一律用 `page["rect"]` 换算坐标，别再用整窗矩形。
    """
    full = grab(rect)
    img, box = content_crop(full)
    crect = [rect[0] + box[0], rect[1] + box[1], rect[0] + box[2], rect[1] + box[3]]
    shot = save_shot(img, f"step4_{tag}_{datetime.now():%H%M%S}")
    data = llm_json(img, OPEN_PROMPT)
    accounts = [a for a in (data.get("accounts") or []) if isinstance(a, dict)]
    return {
        "shot": shot,
        "page": str(data.get("page") or "other"),
        "title": data.get("title"),
        "evidence": data.get("evidence"),
        "accounts": accounts,
        "articles": [a for a in (data.get("articles") or []) if isinstance(a, dict)],
        "rect": crect,
        "img": img,
        "error": data.get("_error"),
    }


def pick_match(entries: list[dict], keyword: str) -> dict | None:
    """从条目里挑和搜索词最匹配的那个：全等 > 包含 > 反过来包含。"""
    kw = keyword.strip()
    cands = [e for e in entries if isinstance(e, dict) and e.get("center_norm")]
    for e in cands:
        if str(e.get("name") or e.get("title") or "").strip() == kw:
            return e
    for e in cands:
        n = str(e.get("name") or e.get("title") or "").strip()
        if n and (kw in n or n in kw):
            return e
    return cands[0] if cands else None


def separator_rows(img, lo: int = 225, hi: int = 252, frac: float = 0.55) -> list[tuple[int, int]]:
    """找出「横向浅灰分隔线」的行（卡片之间的分割线，几乎贯穿整列）。"""
    import numpy as np

    a = np.asarray(img.convert("L")).astype(int)
    weight = ((a >= lo) & (a <= hi)).mean(axis=1)
    rows = np.where(weight > frac)[0]
    if len(rows) == 0:
        return []
    out: list[list[int]] = []
    start = prev = int(rows[0])
    for r in rows[1:]:
        r = int(r)
        if r - prev <= 3:
            prev = r
            continue
        out.append([start, prev])
        start = prev = r
    out.append([start, prev])
    return [(s, e) for s, e in out]


def card_bands(img, min_h: int = 90) -> list[tuple[int, int, int]]:
    """结果卡片在图像内的纵向分带：[(y0, y1, 中心 y)]，自上而下。

    实测（2026-11-02，副屏最大化）：搜一搜结果卡片的浅灰分隔线在裁剪图里位于
    y≈207 / 381 / 582 / 754 / 927，够稳，可以当"第几张卡片"的锚点。
    注意：服务号卡片里的灰色按钮（如「激活防伪/服务大厅」）也会被当成分隔线，
    只会把该卡片切短一点，不影响点到卡片内。
    """
    seps = separator_rows(img)
    bands: list[tuple[int, int, int]] = []
    for i in range(len(seps) - 1):
        y0 = seps[i][1] + 1
        y1 = seps[i + 1][0] - 1
        if y1 - y0 >= min_h:
            bands.append((y0, y1, (y0 + y1) // 2))
    return bands


def _card_center_x(img, y0: int, y1: int, crect: list[int], fallback: int) -> int:
    """卡片在这一行带里的**真实水平中心**（屏幕 x）。

    坑（2026-11-03 实测）：公众号主页的文章卡片是**内缩**的——左边有约 200px 白边，
    `crect[0] + 140` 落在卡片左边的空白槽里，点下去毫无反应（「底盘磕一下…」连点 3 次
    页面仍是 account_home）。搜一搜结果页的卡片贴左，所以那边继续用固定偏移。
    """
    import numpy as np

    a = np.asarray(img.convert("RGB")).astype(int)
    h = a.shape[0]
    y0 = max(0, min(h - 1, int(y0)))
    y1 = max(y0 + 1, min(h, int(y1)))
    nonwhite = (a[y0:y1].min(axis=2) < 245).sum(axis=0)
    cols = np.where(nonwhite >= 3)[0]
    if len(cols) < 20:
        return fallback
    runs: list[list[int]] = []
    start = prev = int(cols[0])
    for c in cols[1:]:
        c = int(c)
        if c - prev <= 30:  # 卡片内部的字间/图间空白不算断开
            prev = c
            continue
        runs.append([start, prev])
        start = prev = c
    runs.append([start, prev])
    lo, hi = max(runs, key=lambda r: r[1] - r[0])  # 最宽的一段 = 卡片本体
    if hi - lo < 60:
        return fallback
    return crect[0] + (lo + hi) // 2


def pick_click_point(
    img, crect: list[int], norm, rank: int = 0, strategy: str = "band", card_x: bool = False
) -> tuple[tuple[int, int] | None, str]:
    """给「第 rank 个结果卡片」算一个点击点，返回 (屏幕坐标, 策略说明)。

    三种策略（按可靠性排序）：
    - band：模型只负责"排序"，坐标由卡片分隔线算（不依赖模型的像素估计）；
    - near：按模型 y 找最近的那条卡片带；
    - vlm ：直接用模型的 center_norm（裁剪图的坐标，误差已缩到几十像素）。
    x 默认固定在内容列左侧 140px（头像/标题区），避开右侧的「上次使用 / N个朋友关注」；
    `card_x=True`（公众号主页的文章卡片）改用该带内卡片自身的水平中心。
    """
    x = crect[0] + 140
    bands = card_bands(img) if img is not None else []
    if strategy == "band" and 0 <= rank < len(bands):
        y0, y1, yc = bands[rank]
        xx = _card_center_x(img, y0, y1, crect, x) if card_x else x
        return (xx, crect[1] + yc), f"band#{rank}"
    if strategy in ("band", "near") and bands and isinstance(norm, list) and len(norm) == 2:
        try:
            y_img = float(norm[1]) * img.height
        except (TypeError, ValueError):
            y_img = None
        if y_img is not None:
            best = min(bands, key=lambda b: abs(b[2] - y_img))
            xx = _card_center_x(img, best[0], best[1], crect, x) if card_x else x
            return (xx, crect[1] + best[2]), f"near#{bands.index(best)}"
    pt = norm_point(crect, norm)
    return pt, "vlm"


def open_account_in_search_window(
    rect: list[int], keyword: str, max_rounds: int = 4, dry_run: bool = False
) -> dict:
    """在搜一搜窗口里找到与关键词对应的公众号并点开。

    实测跳转链：search_results（全部）→ 点账号卡片 → account_list（账号列表）
    → 点同名条目 → account_home（公众号主页：大名字 + 关注/私信 + 文章列表）。
    每跳都靠视觉判页面类型，点到不对的页面就再找再点，最多 max_rounds 次。
    """
    out: dict = {"hops": [], "keyword": keyword, "ok": False}
    if dry_run:
        out["dry_run"] = True
        return out

    page = classify_search_page(rect, "open_0")
    out["hops"].append(
        {"round": 0, "page": page["page"], "title": page["title"], "shot": page["shot"]}
    )

    for rnd in range(1, max_rounds + 1):
        cur = page.get("page")
        title = str(page.get("title") or "")

        if cur == "account_home" and (not title or keyword in title or title in keyword):
            out.update(ok=True, rounds=rnd - 1, page=page["page"], title=title, shot=page["shot"])
            return out

        if cur == "account_home":
            # 主页开错了公众号 → 退回上一页重找
            out["hops"].append({"round": rnd, "action": "back(开错了)", "title": title})
            _send_back()
            time.sleep(2.0)
            page = classify_search_page(rect, f"open_{rnd}")
            out["hops"].append(
                {"round": rnd, "page": page["page"], "title": page["title"], "shot": page["shot"]}
            )
            continue

        if cur in ("search_results", "account_list"):
            entries = page.get("accounts") or []
            best = pick_match(entries, keyword)
            if best is None:
                out["error"] = f"当前页（{cur}）没读到账号条目"
                out["hops"].append({"round": rnd, "action": "无账号条目", "page": cur})
                return out
            crect = page.get("rect") or rect
            img = page.get("img")
            rank = next((i for i, e in enumerate(entries) if e is best), 0)
            # 三种定位策略轮换：先按卡片带序号，再按模型 y 找最近卡片带，最后退回模型坐标
            strategy = ("band", "near", "vlm")[(rnd - 1) % 3]
            pt, how = pick_click_point(img, crect, best.get("center_norm"), rank, strategy)
            if pt is None:
                out["error"] = "账号条目坐标换算失败"
                return out
            click_screen(pt[0], pt[1])
            out["hops"].append(
                {
                    "round": rnd,
                    "action": f"click 账号 {best.get('name')}",
                    "page": cur,
                    "point": list(pt),
                    "how": how,
                    "rank": rank,
                    "crop_rect": crect,
                }
            )
            time.sleep(2.8)
        elif cur == "article":
            out["hops"].append({"round": rnd, "action": "back(误点到文章)", "title": title})
            _send_back()
            time.sleep(2.0)
        else:
            out["hops"].append({"round": rnd, "action": f"未知页面 {cur}", "title": title})
            time.sleep(1.0)

        page = classify_search_page(rect, f"open_{rnd}")
        out["hops"].append(
            {"round": rnd, "page": page["page"], "title": page["title"], "shot": page["shot"]}
        )

    out["error"] = out.get("error") or f"{max_rounds} 轮内没打开公众号主页"
    out["page"] = page.get("page")
    out["title"] = page.get("title")
    out["shot"] = page.get("shot")
    return out


def _send_back() -> None:
    from pywinauto.keyboard import send_keys

    send_keys("%{LEFT}")  # Alt+Left：CEF 页面返回上一页


def step4_search_account(
    keyword: str,
    open_first: bool = True,
    dry_run: bool = False,
) -> dict:
    """步骤4：主聊天窗 → 「公众号」页 → 右上角放大镜 → 搜关键词 → 取结果 → 点开对应公众号。

    open_first=True（默认）：在回车弹出的「搜一搜」独立窗口里，找到与该关键词对应的
    公众号条目并点开其主页（公众号主页：大名字 + 关注/私信 + 文章列表）。
    """
    res: dict = {"stage": "4_search", "keyword": keyword, "events": []}

    def ev(kind: str, **kw) -> None:
        item = {"t": f"{datetime.now():%H:%M:%S}", "event": kind}
        item.update(kw)
        res["events"].append(item)
        log(f"{kind} {kw if kw else ''}")

    win = find_main_window(3.0)
    if win is None:
        res.update(result="error", error="未找到微信主窗，请先跑 --stage launch")
        return res
    if wechat_hung():
        res.update(result="error", error="微信窗口无响应，先处理再跑")
        return res
    # 后面靠键盘输入关键词，先把焦点抢回微信（IDE 里打字会抢焦点；被遮挡还会截错图）
    res["foreground"] = bring_to_front(window_hwnd(win))
    time.sleep(0.6)

    # 4.1 会话列表里必须有「公众号」（被挤出可视区时自动滚列表找）
    _item, names = find_session_item_deep(win, "公众号")
    res["session_names"] = names[:20]
    if _item is None:
        diag = a11y_diagnosis()
        res["a11y"] = diag
        ev("未找到公众号会话", names=names[:10], verdict=diag.get("verdict"))
        if not names:
            # 一条会话都读不到 ≠ 被滚出可视区，是无障碍树本身不可用（§11.30）
            res.update(
                result="error",
                error=f"会话列表读不到任何条目：{diag.get('cause')}",
                fix=diag.get("fix"),
            )
        else:
            res.update(
                result="error",
                error="会话列表里没有「公众号」（可能被滚动出可视区，先手动滚一下再跑）",
            )
        return res

    # 4.2 先清掉历史残留的「搜一搜 / 独立公众号页」窗口，保证后面从干净状态开始
    residue = close_residue_windows(window_hwnd(win), dry_run=dry_run)
    res["residue"] = residue
    ev(
        "清理残留窗口",
        cef=[{"hwnd": c.get("hwnd"), "rect": c.get("rect"), "closed": c.get("closed")} for c in residue["cef"]],
        qt=[{"title": q.get("title"), "closed": q.get("closed")} for q in residue["qt"]],
        closed=residue.get("closed"),
        dry_run=dry_run,
    )
    if residue.get("cef") or residue.get("qt"):
        win = find_main_window(2.0) or win

    # 4.3 真正打开「公众号」页面（幂等 + 校验 + 重试；点不上就不许往下搜）
    page = ensure_gzh_page(win, dry_run=dry_run)
    res["gzh_page"] = page
    ev(
        "公众号页面",
        ok=page.get("ok"),
        already_open=page.get("already_open"),
        buttons=page.get("buttons") or page.get("buttons_before"),
        clicks=[c["attempt"] for c in page.get("clicks", [])],
    )
    if not page.get("ok") and not dry_run:
        res.update(result="error", error=page.get("error") or "公众号页面没打开")
        return res
    win = find_main_window(2.0) or win

    # 4.4 公众号页面右上角放大镜 → 内嵌搜索框 → 粘贴关键词 → 校验 → 回车
    typed = search_in_gzh_page(win, keyword, dry_run=dry_run)
    res["typed"] = typed
    res["search_box_value"] = typed.get("value", "")
    res["text_verified"] = bool(typed.get("ok"))
    ev("搜索框输入", ok=typed.get("ok"), value=typed.get("value"), dry_run=typed.get("dry_run"))
    if not typed.get("ok") and not dry_run:
        res.update(result="error", error=typed.get("error") or "搜索框没收到关键词")
        return res

    if dry_run:
        res.update(result="ok", note="dry_run：只校验到搜索框可定位，未真正搜索")
        return res

    from pywinauto.keyboard import send_keys

    send_keys("{ENTER}")

    # 4.5 回车后结果开在**独立窗口**里（不是主窗！主窗失焦它还会自己关）
    swin = find_search_window(6.0, main_hwnd=window_hwnd(win))
    res["search_window"] = swin
    ev("搜一搜窗口", found=swin is not None, **(swin or {}))
    if swin is None:
        res.update(
            result="error",
            error="回车后没等到「搜一搜」结果窗口（可能被别的窗口抢了焦点，或关键词没提交）",
        )
        return res

    rect = swin["rect"]
    time.sleep(1.5)

    # 4.6 先读一遍：拿到分类行 + 当前分类 + 结果
    data1 = read_search_window(rect, "all")
    if data1.get("_error"):
        res["llm_error"] = data1["_error"]
    res["read_all"] = {
        "shot": data1.get("_shot"),
        "state": data1.get("state"),
        "active_tab": data1.get("active_tab"),
        "tabs": [t.get("name") for t in (data1.get("tabs") or []) if isinstance(t, dict)],
        "results": len(data1.get("results") or []),
        "raw": (data1.get("_raw") or "")[:1500],
        "error": data1.get("_error"),
    }
    ev(
        "读搜索窗口",
        state=data1.get("state"),
        active_tab=data1.get("active_tab"),
        tabs=res["read_all"]["tabs"][:10],
        results=res["read_all"]["results"],
    )

    # 兜底：截图内容不像搜一搜就立刻收手，别拿着别的窗口瞎点 4 轮
    if str(data1.get("state") or "") not in ("results", "empty"):
        res.update(
            result="error",
            error=(
                f"搜索窗口内容异常（state={data1.get('state')}，title={swin.get('title')}），"
                "可能没开搜一搜窗口或被别的窗口顶到前面"
            ),
        )
        return res

    data = data1

    def _norm(rs, crect) -> list[dict]:
        out: list[dict] = []
        for it in rs or []:
            if not isinstance(it, dict):
                continue
            pt = norm_point(crect, it.get("center_norm"))
            if pt is None:
                continue
            out.append(
                {
                    "name": it.get("name"),
                    "kind": it.get("kind"),
                    "desc": it.get("desc"),
                    "center": list(pt),
                }
            )
        return out

    ACC_KINDS = ("账号", "公众号")
    results = _norm(data.get("results"), data.get("_rect") or rect)

    # 「全部」页里通常第一条就是公众号账号卡片（形如「xxx - 账号」）。
    # 只有拿不到账号条目时，才去点「账号」分类再读一遍（模型给 tab 坐标有几十像素误差，仅作兜底）。
    if not any(str(r.get("kind") or "") in ACC_KINDS for r in results):
        tabs = [t for t in (data.get("tabs") or []) if isinstance(t, dict)]
        active = str(data.get("active_tab") or "")
        tab = next((t for t in tabs if str(t.get("name") or "").strip() == "账号"), None)
        if tab is not None and active != "账号":
            pt = norm_point(data.get("_rect") or rect, tab.get("center_norm"))
            if pt is not None:
                click_screen(pt[0], pt[1])
                ev("切到「账号」分类", point=list(pt), norm=tab.get("center_norm"))
                time.sleep(2.2)
                data = read_search_window(rect, "account")
                res["read_account"] = {
                    "shot": data.get("_shot"),
                    "active_tab": data.get("active_tab"),
                    "results": len(data.get("results") or []),
                    "raw": (data.get("_raw") or "")[:1200],
                }
                acc = _norm(data.get("results"), data.get("_rect") or rect)
                ev(
                    "读账号分类",
                    active_tab=data.get("active_tab"),
                    results=len(acc),
                )
                if acc:
                    results = acc

    # 4.7 账号条目排前面
    results.sort(key=lambda r: 0 if str(r.get("kind") or "") in ACC_KINDS else 1)
    res["results"] = results
    res["results_source"] = "llm" if results else ""
    ev(
        "搜索结果",
        count=len(results),
        names=[(r.get("name") or "")[:20] for r in results[:6]],
        source=res["results_source"],
    )

    # 4.8 在搜一搜窗口里找到「与关键词对应的公众号」并点开（默认就做）
    if not open_first:
        res["open_account"] = {"ok": False, "skipped": "已用 --no-open 跳过"}
    else:
        opened = open_account_in_search_window(rect, keyword, dry_run=False)
        res["open_account"] = opened
        ev(
            "点开公众号",
            ok=opened.get("ok"),
            page=opened.get("page"),
            title=opened.get("title"),
            rounds=opened.get("rounds"),
            hops=[
                {k: h.get(k) for k in ("round", "action", "page", "title")}
                for h in opened.get("hops", [])
            ],
            error=opened.get("error"),
        )
        try:
            res["windows_after"] = [
                {k: v for k, v in w.items() if k in ("hwnd", "class", "title", "rect")}
                for w in wechat_top_windows()[:5]
            ]
        except Exception:  # noqa: BLE001
            pass

    acct_ok = bool((res.get("open_account") or {}).get("ok"))
    res["result"] = "ok" if (results and (acct_ok or not open_first)) else ("partial" if results else "empty")
    if not results:
        res["error"] = "搜索窗口里没读到结果条目（截图见 read_all.shot）"
    elif open_first and not acct_ok:
        res["error"] = (res.get("open_account") or {}).get("error") or "没能点开对应的公众号主页"
    return res


# --------------------------------------------------------------------------
# 步骤 5：在公众号主页采集「文章 / 视频号」列表（按最近 N 天过滤）
# --------------------------------------------------------------------------
FEED_PROMPT = """你在看某个微信公众号的主页（微信「搜一搜」里打开的公众号详情页）。

只输出一个 JSON 对象，不要解释、不要 Markdown：
{
  "page": "account_home|account_list|search_results|article|other",
  "title": "公众号名称",
  "active_tab": "当前选中的分项（全部/贴图/文章/视频号）",
  "tabs": [{"name": "全部|贴图|文章|视频号", "center_norm": [0.0, 0.0]}],
  "date_headers": [{"label": "今天|昨天|星期五|10月30日", "center_norm": [0.0, 0.0]}],
  "items": [{"title": "条目标题", "date_label": "它所属的日期分组标签", "kind": "文章|视频|贴图|其他", "extra": "阅读数/播放数/时长等", "center_norm": [0.0, 0.0]}],
  "bottom_reached": false
}

要点：
- **page 先判型**：如果画面是一篇文章的正文页（有大段正文/配图，底部有「阅读 982」「赞/在看/分享/写留言」，
  或顶部是文章标题+公众号名+发布时间），`page="article"`，此时 items 留空、不要编造卡片；
- 页面结构：顶部是公众号信息 + 关注/私信按钮，下面一排分项「全部 贴图 文章 视频号」，
  再下面是按日期分组的卡片流（分组标签形如「今天」「昨天」「星期五」「10月30日」）；
- items 只列**当前可见**的卡片，从上到下按出现顺序，每条都要给出它所在日期分组的标签；
- **不要**把「2个内容」这类**合集/专辑折叠条**、「查看全部 / 查看更多」控件、「置顶」这种分组标签本身当成条目；
- date_label 必须照抄页面上的分组文字，不要自己推算日期；
- bottom_reached：页面已经滑到底、下面没有更多内容时为 true；
- center_norm 是相对整张图 0~1 的比例坐标；看不清给 null，不要编造。
"""

TAB_ORDER = ("全部", "贴图", "文章", "视频号")
WEEKDAY_NAMES = ("星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日")

# 视频号卡片上没有日期分组标签（实测 2026-11-02），没法按天筛；按"当天更新一般不超过 4 条"
# 的经验直接取列表最前面 4 条，且不往下滚。
VIDEO_TAB_LIMIT = 4

# --feed-fill-window：时间窗口没闭合时最多在 --max-scroll 之外额外补几屏（防止无限下滚）
FILL_WINDOW_EXTRA = 6

# 列表里「不是文章」的条目：合集/专辑折叠条（实测「2个内容」）和「查看全部」这类控件文字。
# 视觉模型会把它们当条目读进清单，但步骤 6 点不开（报「没能在文章列表里点开这篇」）→ 清单阶段就丢。
FEED_CHROME_RE = re.compile(
    r"^(?:\d+\s*[个篇条]\s*(?:内容|文章|合集)|(?:查看全部|查看更多|展开全部|更多))$"
)

# 「置顶」分组：卡片上**没有发布时间**，只有「置顶」标签，所以日期解析不出来。
# 它只是被长期钉在列表最前，不等于「最近发的」——实测（2026-11-03 21:44 长安汽车）
# 今天没发文章，置顶那条实际是「星期六」的旧文，被当未知日期保留后步骤 6 白点一篇。
PINNED_LABELS = ("置顶",)


def _next_header_offset(headers: list, after: str) -> int | None:
    """`headers` 里 `after` **之后**第一个能解析出天数的分组标签，返回距今天数。

    置顶卡片上没有日期，只能借它下面第一个日期分组当日期上界（置顶文不会比它更新到
    超出那个分组还多的程度；判不出就返回 None，按原来的「未知日期」保留）。
    """
    try:
        k = headers.index(after) + 1
    except ValueError:
        return None
    for h in headers[k:]:
        off = parse_day_offset(h)
        if off is not None:
            return off
    return None


def content_crop(img):
    """把截图裁到「页面内容列」，返回 (裁剪图, [x0, y0, x1, y1])。

    坑（2026-11-02 实测）：搜一搜窗口在副屏最大化时宽 1928px，两侧是 #f7f7f7 留白、
    内容卡片区是纯白 #ffffff。整窗喂给模型时它把「全部」判在 x=0.63（实际 0.31），
    误差 600px，点 tab 必点飞；裁成 748 宽的内容列后误差降到 ~50px，
    再叠加像素聚类即可精确命中。
    """
    import numpy as np

    a = np.asarray(img.convert("RGB")).astype(int)
    white = a.min(axis=2) >= 252
    col_frac = white.mean(axis=0)
    cols = np.where(col_frac > 0.25)[0]
    if len(cols) < 100:
        return img, [0, 0, img.width, img.height]
    x0, x1 = int(cols.min()), int(cols.max())
    row_frac = white[:, x0 : x1 + 1].mean(axis=1)
    rows = np.where(row_frac > 0.15)[0]
    y0, y1 = (int(rows.min()), int(rows.max())) if len(rows) else (0, img.height - 1)
    if (x1 - x0) < img.width * 0.25 or (y1 - y0) < img.height * 0.3:
        return img, [0, 0, img.width, img.height]
    return img.crop((x0, y0, x1 + 1, y1 + 1)), [x0, y0, x1 + 1, y1 + 1]


def grab_content(rect: list[int]) -> tuple[object, list[int]]:
    """抓窗口截图并裁到内容列，返回 (裁剪图, 该图对应的屏幕矩形)。"""
    full = grab(rect)
    img, box = content_crop(full)
    screen_rect = [rect[0] + box[0], rect[1] + box[1], rect[0] + box[2], rect[1] + box[3]]
    return img, screen_rect


def _ensure_front(hwnd: int) -> None:
    """确保目标窗口在最前面。

    坑（2026-11-02 实测）：`grab(rect)` 抓的是**屏幕区域**，不是窗口内容。微信主窗
    (1955,23)-(2894,1008) 和搜一搜 CEF 窗口 (1916,-4)-(3844,1084) 在副屏上大面积重叠，
    只要主窗在前面，截出来就是主窗的「公众号」页，模型一路读成 page=other/title=公众号。
    """
    if not hwnd:
        return
    try:
        import win32gui

        if int(win32gui.GetForegroundWindow()) == int(hwnd):
            return
        bring_to_front(int(hwnd))
        time.sleep(0.4)
    except Exception:  # noqa: BLE001
        pass


def read_feed_page(rect: list[int], tag: str = "feed", hwnd: int = 0) -> dict:
    """截图（裁内容列）→ 让模型结构化读公众号主页。"""
    _ensure_front(hwnd)
    img, crect = grab_content(rect)
    shot = save_shot(img, f"step5_{tag}_{datetime.now():%H%M%S}")
    data = llm_json(img, FEED_PROMPT, max_tokens=1800)
    data["_shot"] = shot
    data["_rect"] = crect
    data["_img"] = img
    return data


def _dark_clusters(gray_row: object, thresh: int = 120, gap: int = 8, lo: int = 8, hi: int = 70) -> list[tuple[int, int]]:
    """在一行（或多行叠加）的灰度里，把「深色文字」聚成若干列簇。"""
    import numpy as np

    dark = (gray_row < thresh).sum(axis=0)
    cols = np.where(dark >= 1)[0]
    if len(cols) == 0:
        return []
    out: list[list[int]] = []
    start = prev = int(cols[0])
    for c in cols[1:]:
        c = int(c)
        if c - prev <= gap:
            prev = c
            continue
        out.append([start, prev])
        start = prev = c
    out.append([start, prev])
    return [(s, e) for s, e in out if lo <= (e - s) <= hi]


def _text_row_extent(a, centers: list[int], yc: int, half: int = 18, pad: int = 45) -> int:
    """在 yc 附近求这一行文字的**垂直中心**（点 tab 必须点在文字上，点空白不生效）。"""
    import numpy as np

    lo = max(0, yc - half)
    hi = min(a.shape[0], yc + half)
    if hi - lo < 3:
        return yc
    sub = a[lo:hi, :]
    mask = np.zeros(sub.shape[1], dtype=bool)
    for c in centers:
        mask[max(0, c - pad) : min(sub.shape[1], c + pad)] = True
    rows = np.where((sub[:, mask] < 120).sum(axis=1) >= 2)[0]
    if len(rows) == 0:
        return yc
    runs: list[tuple[int, int]] = []
    start = prev = int(rows[0])
    for r in rows[1:]:
        r = int(r)
        if r - prev <= 2:
            prev = r
            continue
        runs.append((start, prev))
        start = prev = r
    runs.append((start, prev))
    y_rel = yc - lo
    best = min(
        runs,
        key=lambda ab: 0 if ab[0] - 2 <= y_rel <= ab[1] + 2 else min(abs(ab[0] - y_rel), abs(ab[1] - y_rel)),
    )
    return lo + (best[0] + best[1]) // 2


def tab_row_centers(img, y_hint: int, expect: int = 4) -> tuple[list[int], int]:
    """在 y_hint 附近找「分项标签行」，返回 (各标签 x 中心, 该行文字的垂直中心)。

    首页有 4 个分项（全部/贴图/文章/视频号），所以要求正好 4 个等距文字簇；
    找不到就返回 ([], y_hint)，交给模型坐标兜底。

    坑（2026-11-02 实测）：早先直接返回命中的扫描线 y，落在 tab 文字**上方十几像素的空白**里，
    点下去毫无反应（日志连续 3 次"点了 3 次也没切到文章"）；必须用 `_text_row_extent`
    把 y 吸到文字行中心。
    """
    import numpy as np

    a = np.asarray(img.convert("L")).astype(int)
    h, w = a.shape
    cands: list[tuple[int, int, list[int], int]] = []
    # 不能只在 y_hint ±45px 里找：模型给的 y 经常差 100~200px（实测主页 tab 在 y=415，
    # 模型说 0.28*816≈228），窗口一改尺寸就整行错过。这里扫整个上半屏再按距 y_hint 取最近。
    lo = max(8, int(h * 0.08))
    hi = min(h - 8, int(h * 0.62))
    for yc in range(lo, hi, 3):
        band = a[yc - 8 : yc + 8, :]
        cl = _dark_clusters(band)
        # 贴边的深色带不是文字：左边是窗口边框/隔壁屏的暗边，右边是滚动条。
        # 实测（2026-11-03 21:25）它们各占一个簇 → 分项行被数成 5 簇 ≠ 4 → 返回空、
        # 退回模型坐标点偏（连点 3 次没切到「文章」）。真正的分项文字不会贴边。
        cl = [(s, e) for s, e in cl if s > 12 and e < w - 12]
        if len(cl) != expect:
            continue
        centers = [(s + e) // 2 for s, e in cl]
        gaps = [centers[i + 1] - centers[i] for i in range(len(centers) - 1)]
        if not gaps or min(gaps) < 20 or max(gaps) > 160:
            continue
        if max(gaps) > min(gaps) * 2.2:  # 分项之间间距基本均匀
            continue
        if centers[0] < w * 0.05 or centers[-1] > w * 0.9:
            continue
        if sum(e - s for s, e in cl) < 60:  # 4 个标签总得有字
            continue
        density = int((band < 120).sum())
        cands.append((abs(yc - y_hint), -density, centers, yc))
    if not cands:
        return ([], y_hint)
    _dist, _dens, centers, yc = min(cands)
    return (centers, _text_row_extent(a, centers, yc))


def click_feed_tab(rect: list[int], name: str, rows: int = 3, hwnd: int = 0) -> dict:
    """点公众号主页的分项 tab（全部/贴图/文章/视频号），点完校验 active。"""
    out: dict = {"tab": name, "attempts": []}
    for i in range(1, max(rows, 1) + 1):
        data = read_feed_page(rect, f"tab_{name}", hwnd=hwnd)
        if str(data.get("page") or "") != "account_home":
            out.update(ok=False, error=f"当前不是公众号主页（page={data.get('page')}）")
            return out
        if str(data.get("active_tab") or "").strip() == name:
            out.update(ok=True, already=True, attempt=i)
            return out

        img, crect = grab_content(rect)
        tabs = [t for t in (data.get("tabs") or []) if isinstance(t, dict)]
        vlm_tab = next((t for t in tabs if str(t.get("name") or "").strip() == name), None)

        y_hint = int(img.height * 0.35)
        if vlm_tab and isinstance(vlm_tab.get("center_norm"), list) and len(vlm_tab["center_norm"]) == 2:
            try:
                y_hint = int(float(vlm_tab["center_norm"][1]) * img.height)
            except (TypeError, ValueError):
                pass

        centers, y_row = tab_row_centers(img, y_hint)
        if len(centers) == len(TAB_ORDER) and name in TAB_ORDER:
            x = centers[TAB_ORDER.index(name)]
            via = "pixel"
        elif vlm_tab is not None:
            pt = norm_point(crect, vlm_tab.get("center_norm"))
            if pt is None:
                out["attempts"].append({"attempt": i, "error": "tab 坐标为空"})
                continue
            x, y_row = pt[0] - crect[0], pt[1] - crect[1]
            via = "vlm"
        else:
            out["attempts"].append({"attempt": i, "error": f"没读到 tab {name}"})
            continue

        click_screen(crect[0] + x, crect[1] + y_row)
        out["attempts"].append(
            {"attempt": i, "via": via, "point": [crect[0] + x, crect[1] + y_row], "centers": centers}
        )
        time.sleep(2.0)

        chk = read_feed_page(rect, f"tab_{name}_after", hwnd=hwnd)
        active = str(chk.get("active_tab") or "").strip()
        out["attempts"][-1]["active_after"] = active
        log(f"点分项「{name}」第{i}次（{via}）→ 当前={active or '未知'}")
        if active == name:
            out.update(ok=True, attempt=i, active=active)
            return out

    out.update(ok=False, error=f"点了 {rows} 次也没切到「{name}」")
    return out


def parse_day_offset(label: str, today=None) -> int | None:
    """把主页上的日期分组标签换算成「距今几天」（今天=0，昨天=1）。

    支持：今天 / 昨天 / 前天 / 星期X / M月D日 / YYYY年M月D日 / N小时前 / N分钟前。
    认不出返回 None（交由调用方决定是保留还是丢弃）。
    """
    from datetime import date as _date
    from datetime import timedelta

    if not label:
        return None
    s = str(label).strip().replace(" ", "")
    today = today or _date.today()
    if s in ("今天", "今日", "刚刚") or s.endswith("分钟前") or s.endswith("小时前"):
        return 0
    if s in ("昨天", "昨日"):
        return 1
    if s == "前天":
        return 2
    if s in WEEKDAY_NAMES:
        delta = (today.weekday() - WEEKDAY_NAMES.index(s)) % 7
        return delta or 7  # 当天会用"今天"，所以 0 其实是 7 天前
    m = re.fullmatch(r"(\d{4})年(\d{1,2})月(\d{1,2})日", s)
    if m:
        try:
            d = _date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            return (today - d).days
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{1,2})月(\d{1,2})日", s)
    if m:
        try:
            d = _date(today.year, int(m.group(1)), int(m.group(2)))
        except ValueError:
            return None
        if d > today:  # 跨年：显示的是去年
            try:
                d = d.replace(year=today.year - 1)
            except ValueError:
                return None
        return (today - d).days
    if s in ("", "—", "-"):
        return None
    del timedelta
    return None


def scroll_page(rect: list[int], notches: int = 3) -> None:
    """把鼠标放到页面中间滚轮翻页（负值向下）。"""
    import win32api
    import win32con

    cx = (rect[0] + rect[2]) // 2
    cy = (rect[1] + rect[3]) // 2
    win32api.SetCursorPos((cx, cy))
    time.sleep(0.2)
    for _ in range(abs(notches)):
        win32api.mouse_event(
            win32con.MOUSEEVENTF_WHEEL, 0, 0, -120 if notches > 0 else 120, 0
        )
        time.sleep(0.15)
    time.sleep(1.3)


def _img_diff(a, b) -> float:
    """两张图的平均像素差，用来判断"滚不动了"。"""
    import numpy as np

    try:
        x = np.asarray(a.convert("L").resize((160, 160))).astype(int)
        y = np.asarray(b.convert("L").resize((160, 160))).astype(int)
        return float(np.abs(x - y).mean())
    except Exception:  # noqa: BLE001
        return 999.0


def collect_feed_tab(
    rect: list[int],
    tab: str,
    days: int,
    max_scroll: int = 0,
    dry_run: bool = False,
    hwnd: int = 0,
    limit: int = 0,
    fill_window: bool = False,
) -> dict:
    """读某个分项（默认「文章」）当前屏的条目，按 days 过滤。

    `fill_window=True`：时间窗口还没闭合（这一屏最下面一条仍在窗口内）时**自动多滚一屏**，
    最多在 `max_scroll` 之外补 6 屏。实测 `--days 1` 只读当前屏会漏 —— 今天有 2 篇，
    当前屏只显示 1 篇（`seen=1`），补滚后 `seen=2`。

    默认 `max_scroll=0`：**只读当前屏幕上的列表，不往下滚**（2026-11-03 用户要求）。
    需要历史内容时才显式给 `max_scroll>0`，那时才按「滚到日期早于 days 或到底」停。

    `limit > 0` 时走「按条数取」模式（视频号这类没有日期分组的分项）：只取列表最前面
    的 limit 条，不做日期过滤。
    """
    out: dict = {"tab": tab, "rounds": [], "items": [], "stop": None}
    if limit:
        out["limit"] = limit
    seen: dict[str, dict] = {}
    prev_img = None

    cap = max_scroll  # fill_window 会在这个上限之上临时加屏
    window_open = False
    i = 0
    while i <= cap:
        data = read_feed_page(rect, f"{tab}_{i}", hwnd=hwnd)
        img = data.get("_img")
        items = [x for x in (data.get("items") or []) if isinstance(x, dict)]
        headers = [
            h.get("label")
            for h in (data.get("date_headers") or [])
            if isinstance(h, dict) and h.get("label")
        ]
        fresh = 0
        too_old = False
        last_off = None  # 这一屏**最下面**一条的距今天数：判断时间窗口有没有闭合
        for it in items:
            title = str(it.get("title") or "").strip()
            if not title:
                continue
            if FEED_CHROME_RE.match(title):
                # 主页上的合集/专辑折叠条（如「2个内容」）不是文章：步骤6 点不开它
                # （实测 21:28 那轮 index0 报「没能在文章列表里点开这篇」）
                out["dropped_chrome"] = int(out.get("dropped_chrome") or 0) + 1
                continue
            label = str(it.get("date_label") or "").strip()
            pinned = label in PINNED_LABELS
            off = parse_day_offset(label)
            if off is None and pinned:
                off = _next_header_offset(headers, label)  # 置顶卡片没日期，借下面第一个分组
            if off is not None and off >= days:
                too_old = True
                if pinned:
                    # 置顶的旧文不算「这段时间的内容」，不进清单（否则步骤 6 去点一篇窗口外的文章）
                    out["dropped_pinned_older"] = int(out.get("dropped_pinned_older") or 0) + 1
                    continue
            if off is not None:
                last_off = off
            if title in seen:
                continue
            fresh += 1
            seen[title] = {
                "title": title,
                "date_label": it.get("date_label"),
                "day_offset": off,
                "pinned": pinned or None,
                "kind": it.get("kind"),
                "extra": it.get("extra"),
                "batch": i,
            }
        out["rounds"].append(
            {
                "round": i,
                "shot": data.get("_shot"),
                "items": len(items),
                "fresh": fresh,
                "headers": headers,
                "bottom_reached": bool(data.get("bottom_reached")),
                "state": data.get("page"),
            }
        )
        log(f"[{tab}] 第{i}轮：读到 {len(items)} 条（新增 {fresh}），日期分组 {headers}")

        # 分组标签本身也参与「窗口闭合」判定：屏幕上已经出现超出窗口的分组（如 --days 1 时看到「昨天」），
        # 说明窗口内的内容已经过完，不用再补滚（实测少判这一条会白多滚一屏）
        for label in headers:
            ho = parse_day_offset(label)
            if ho is not None and ho >= days:
                too_old = True

        # 时间窗口有没有闭合：必须在下面的 break 之前算，否则「已到更早内容」那轮的残留值会误报提醒
        window_open = bool(not too_old and last_off is not None and last_off < days)

        if limit and len(seen) >= limit:
            out["stop"] = f"已取到前 {limit} 条（该分项无日期分组）"
            break
        if too_old and not limit:
            out["stop"] = "已到比目标天数更早的内容"
            break
        if data.get("bottom_reached"):
            out["stop"] = "到底了"
            break
        if dry_run:
            out["stop"] = "dry_run"
            break
        if prev_img is not None and img is not None and _img_diff(prev_img, img) < 1.0:
            out["stop"] = "滚不动（内容没变化）"
            break
        if (
            fill_window
            and not limit
            and window_open
            and i >= cap
            and cap < max_scroll + FILL_WINDOW_EXTRA
        ):
            cap += 1
            out["auto_extend"] = int(out.get("auto_extend") or 0) + 1
            log(f"[{tab}] 窗口未闭合（最后一条距今 {last_off} 天）→ 自动补滚 1 屏（累计 {cap} 屏）")
        if i >= cap:
            break  # 默认 max_scroll=0 且不开 fill_window：读完当前屏就收
        prev_img = img
        scroll_page(rect)
        i += 1

    out["window_open"] = window_open  # True=窗口边界还没读到，清单可能不完整
    if not out["stop"]:
        ext = int(out.get("auto_extend") or 0)
        out["stop"] = (
            f"滚到 {cap} 屏上限（max_scroll={max_scroll}，自动补滚 {ext} 屏）"
            if (cap or ext)
            else "只取当前屏（max_scroll=0）"
        )

    all_items = list(seen.values())
    if limit:
        # 无日期分组的分项（视频号）：按列表顺序取前 limit 条，不做日期过滤
        kept = all_items[:limit]
        dropped = 0
    else:
        kept = [x for x in all_items if x["day_offset"] is None or x["day_offset"] < days]
        dropped = len(all_items) - len(kept)
    out["items"] = kept
    out["total_seen"] = len(all_items)
    out["dropped_older_than_days"] = dropped
    out["unknown_date"] = len([x for x in kept if x["day_offset"] is None])
    return out


def step5_collect_feed(
    days: int = 1,
    tabs: tuple[str, ...] = ("文章",),
    max_scroll: int = 0,
    dry_run: bool = False,
    video_limit: int = VIDEO_TAB_LIMIT,
    fill_window: bool = False,
) -> dict:
    """步骤5：在已打开的公众号主页里，读「文章」分项当前屏的列表，按最近 N 天过滤。

    前置：步骤 4 已经把某个公众号主页（account_home）打开在搜一搜窗口里。

    默认只采「文章」且**不滚动**（2026-11-03 要求：不一直往下滚、不再采视频号）。
    要历史内容就显式给 `--max-scroll N`；要视频号就显式 `--tabs 文章,视频号`
    （视频号卡片没有日期分组，按 `video_limit` 取前几条）。
    """
    res: dict = {
        "stage": "5_feed",
        "days": days,
        "tabs": list(tabs),
        "max_scroll": max_scroll,  # 步骤6 要靠它决定「清单跨了几屏、找文章要往下滚几屏」
        "video_limit": video_limit,
        "events": [],
        "collected": {},
    }

    def ev(kind: str, **kw) -> None:
        item = {"t": f"{datetime.now():%H:%M:%S}", "event": kind}
        item.update(kw)
        res["events"].append(item)
        log(f"{kind} {kw if kw else ''}")

    # 主窗只用来把「搜一搜」CEF 窗口排除掉；微信缩在托盘时 UIA 找不到主窗，
    # 但 CEF 页面照常在，可以继续采集（实测 23:31 那轮就是这样卡在"找不到微信主窗"）。
    win = find_main_window(3.0)
    main_hwnd = window_hwnd(win) if win is not None else 0
    if win is None:
        res["warning"] = "没找到微信主窗（可能缩在托盘），只按搜一搜窗口继续"

    swin = find_search_window(3.0, main_hwnd=main_hwnd)
    res["search_window"] = swin
    if swin is None:
        res.update(result="error", error="没有搜一搜窗口（先跑步骤 4 打开公众号主页）")
        return res
    rect = swin["rect"]
    fh = int(swin["hwnd"])
    # 必须把「搜一搜」窗口自己抬到最前：主窗在副屏上大面积盖住它，抬主窗=截到主窗
    _ensure_front(fh)

    page = read_feed_page(rect, "pre", hwnd=fh)
    res["account"] = page.get("title")
    res["page"] = page.get("page")
    ev("读数", page=page.get("page"), title=page.get("title"), shot=page.get("_shot"))
    if str(page.get("page") or "") != "account_home":
        res.update(
            result="error",
            error=f"当前不是公众号主页（page={page.get('page')}），请先跑步骤 4",
        )
        return res
    if dry_run:
        res.update(result="ok", dry_run=True, note="只确认在公众号主页，未采集")
        return res

    for tab in tabs:
        clk = click_feed_tab(rect, tab, hwnd=fh)
        res["collected"].setdefault(tab, {})
        res["collected"][tab]["click"] = {
            k: clk.get(k)
            for k in ("ok", "already", "attempt", "active", "error", "attempts")
        }
        ev("切分项", tab=tab, ok=clk.get("ok"), error=clk.get("error"))
        if not clk.get("ok"):
            res["collected"][tab]["error"] = clk.get("error")
            continue
        got = collect_feed_tab(
            rect,
            tab,
            days,
            max_scroll=max_scroll,
            hwnd=fh,
            limit=video_limit if tab == "视频号" else 0,
            fill_window=fill_window,
        )
        res["collected"][tab].update(got)
        ev(
            "采集完成",
            tab=tab,
            kept=len(got["items"]),
            seen=got.get("total_seen"),
            dropped=got.get("dropped_older_than_days"),
            pinned_dropped=got.get("dropped_pinned_older"),
            stop=got.get("stop"),
        )

    # 窗口没闭合（最后一屏最下面一条仍在窗口内）= 清单可能不完整，提醒但不擅自改行为
    open_tabs = [t for t, v in res["collected"].items() if v.get("window_open")]
    if open_tabs:
        res["window_warning"] = (
            f"{'、'.join(open_tabs)}：时间窗口没闭合（最后一屏最下面一条仍在窗口内），清单可能不完整"
            " → 加 --max-scroll N 或 --feed-fill-window 重跑"
        )
        ev("窗口提醒", tabs=open_tabs, days=days, max_scroll=max_scroll, fill_window=fill_window)

    # 清单实际跨了几屏（含 fill_window 自动补滚的）→ 步骤 6 找文章要按这个值决定滚几屏
    res["scrolled"] = max(
        [max_scroll + int(v.get("auto_extend") or 0) for v in res["collected"].values()] + [0]
    )
    ok_tabs = [t for t, v in res["collected"].items() if v.get("click", {}).get("ok")]
    res["total"] = sum(len(v.get("items") or []) for v in res["collected"].values())
    res["result"] = "ok" if ok_tabs else "error"
    if not ok_tabs:
        res["error"] = "所有分项都没切进去"
    return res


# --------------------------------------------------------------------------
# 步骤 6：按步骤 5 的文章清单逐篇打开，抓正文（注意关掉弹出的新窗口）
# --------------------------------------------------------------------------
ARTICLE_PROMPT = """你在看一篇微信公众号的图文文章（可能开在微信的独立窗口里）。

只输出一个 JSON 对象，不要解释、不要 Markdown：
{
  "page": "article|account_home|account_list|search_results|other",
  "title": "文章标题",
  "author": "公众号名称/作者",
  "publish_time": "页面上显示的发布时间原文",
  "paragraphs": ["正文段落或小标题，按页面从上到下"],
  "bottom_reached": false
}

要点：
- paragraphs 只抄**当前屏幕上看得清**的正文文字，不要总结、不要改写、不要补全；
- 「在看/点赞/分享/写留言/推荐阅读/相关文章」这类按钮和推荐位不算正文；
- **bottom_reached 只看正文本身有没有结束**：右下角那排「赞 / 在看 / 分享 / 写留言」是**吸底工具栏，
  文章任何位置都看得见**，看见它**不算**到底（实测这样会误判、正文只抄到一半就停）；
  只有正文下面出现「二维码 + 关注卡片」「合集·目录」「上一篇/下一篇」「阅读 982」这类**文末区块**，
  或最后一段下面确实没有更多正文了，才 bottom_reached=true；
- 拿不准就填 false —— 宁可多滚一轮，别提前停；
- 看不清就给空数组或 null，不要编造。
"""

_TITLE_NOISE = re.compile(r"[\s\u3000「」『』【】\[\]（）()《》<>\"'“”‘’·…—\-|,，。.、:：;；!！?？~～]+")


def _norm_title(s: str) -> str:
    """标题归一化：去空白与标点，用于"列表标题 ↔ 打开后的文章标题"模糊比对。"""
    return _TITLE_NOISE.sub("", str(s or "")).lower()


def _title_match(a: str, b: str, min_len: int = 8) -> bool:
    na, nb = _norm_title(a), _norm_title(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    n = min(len(na), len(nb), 20)
    if n >= min_len and na[:n] == nb[:n]:
        return True
    return na in nb or nb in na


def cef_hwnds(title_filter: bool = False) -> set[int]:
    return {int(w["hwnd"]) for w in cef_windows(title_filter=title_filter)}


def _pick_new_cef(before: set[int]) -> dict | None:
    """点文章后新冒出来的 CEF 窗口（文章详情常常是独立窗口，标题就是文章标题）。"""
    cands = [w for w in cef_windows(title_filter=False) if int(w["hwnd"]) not in before]
    if not cands:
        return None
    return max(
        cands,
        key=lambda w: (w["rect"][2] - w["rect"][0]) * (w["rect"][3] - w["rect"][1]),
    )


def close_cef_window(hwnd: int) -> dict:
    """按 HWND 关掉一个 CEF 窗口（只发 WM_CLOSE，绝不碰微信主窗）。"""
    import win32con
    import win32gui

    out: dict = {"hwnd": int(hwnd)}
    try:
        if not win32gui.IsWindow(int(hwnd)):
            out.update(closed=True, note="窗口已不存在")
            return out
        win32gui.PostMessage(int(hwnd), win32con.WM_CLOSE, 0, 0)
        time.sleep(1.2)
        out["closed"] = not win32gui.IsWindow(int(hwnd))
    except Exception as exc:  # noqa: BLE001
        out.update(closed=False, error=f"{type(exc).__name__}: {exc}")
    return out


def read_article_page(rect: list[int], tag: str = "article", hwnd: int = 0) -> dict:
    """截图（裁内容列）→ 让模型抄当前屏幕的正文。"""
    _ensure_front(hwnd)
    img, crect = grab_content(rect)
    shot = save_shot(img, f"step6_{tag}_{datetime.now():%H%M%S}")
    data = llm_json(img, ARTICLE_PROMPT, max_tokens=2400)
    data["_shot"] = shot
    data["_rect"] = crect
    data["_img"] = img
    return data


ARTICLE_CLIP_PROMPT = """下面是一篇微信公众号文章用「Ctrl+A 全选 → Ctrl+C 复制」拿到的原始文本，
里面混着页面外壳（公众号简介、关注/在看/点赞/分享、阅读量、留言区、上一篇/下一篇、合集目录、广告等）。

只输出一个 JSON 对象，不要解释、不要 Markdown：
{
  "title": "文章标题",
  "author": "公众号名称",
  "publish_time": "页面上显示的发布时间原文，没有就 null",
  "paragraphs": ["正文段落，照抄原文原句，按原文顺序"],
  "is_image_article": false,
  "dropped": ["去掉的外壳片段类型，简短列出"]
}

要点：
- **正文原句照抄**：不总结、不改写、不补全、不把多条合并成一条；
- 小标题和编号条目（「车型·车市」「1.10月销量：…」）都算正文，要保留；
- 去掉这些外壳：公众号简介/口号（如「只做有价值的汽车科技报道」）、「试用申请/登录官网/网址」、
  单独的「图片」占位行、图片文件名（形如「主图.jpg」「第一版logo.gif」「xx.png」——复制时图片只会
  留下文件名，不是正文）、「赞/在看/分享/写留言/留言/还没有留言，来写首评/暂无评论」、阅读量数字行、
  上一篇/下一篇标题、「合集·目录」、举报投诉；
- 若整篇几乎没有连续正文文字（只有标题 + 简介 + 「图片」占位），说明正文是一张长图 →
  `is_image_article=true`，paragraphs 只留能确认属于正文的文字。
"""


def _clip_text() -> str:
    """读剪贴板纯文本（没有则空串）。"""
    import win32clipboard

    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                return str(win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT) or "")
        finally:
            win32clipboard.CloseClipboard()
    except Exception:  # noqa: BLE001
        return ""


def _set_clip_text(s: str) -> None:
    import win32clipboard

    try:
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, s)
        finally:
            win32clipboard.CloseClipboard()
    except Exception:  # noqa: BLE001
        pass


def grab_article_clipboard(rect: list[int], hwnd: int = 0, tag: str = "clip") -> dict:
    """`Ctrl+A` → `Ctrl+C` → 读剪贴板，一次拿全文。

    实测（2026-11-03）：比逐屏截图抄更好 —— 一次 5584 字全文，而逐屏 15 轮只有 4222 字，
    还得处理跨屏重叠去重。

    **默认不点击**（2026-11-03 改）：原来固定点正文区 45% 高度处，那个位置很容易落在**配图**上，
    微信文章里点图片会弹出大图预览、页面离开文章页，`Ctrl+A` 就废了（「汽车公社」那篇实测踩到）。
    而 CEF 标签页抬到前台时文档本来就有键盘焦点 —— **完全不点也能复制**（实测一次 12628 字）。
    所以策略按「代价从小到大」排：不点 → 按 Esc 清掉残留再复制 → 点标题行（纯文本，不会触发预览）。
    """
    from pywinauto.keyboard import send_keys

    out: dict = {"ok": False, "chars": 0, "text": "", "attempts": []}
    old_clip = _clip_text()  # 用完还原，别长期占着用户剪贴板
    title_pt = (
        (rect[0] + rect[2]) // 2,
        int(rect[1] + (rect[3] - rect[1]) * 0.10),  # 标题行：一定是文本，点不出图片预览
    )
    strategies = [("no_click", None), ("esc_no_click", None), ("title_click", title_pt)]

    for name, pt in strategies:
        sentinel = f"__WXCLIP__{name}"
        _set_clip_text(sentinel)
        _ensure_front(hwnd)
        time.sleep(0.5)
        if name == "esc_no_click":
            send_keys("{ESC}")  # 清掉可能残留的选区/预览
            time.sleep(0.6)
        if pt is not None:
            click_screen(pt[0], pt[1])
            time.sleep(1.0)
        send_keys("^a")
        time.sleep(1.1)
        send_keys("^c")
        time.sleep(1.4)
        text = _clip_text()
        got = 0 if text == sentinel else len(text)
        out["attempts"].append({"try": name, "click": list(pt) if pt else None, "chars": got})
        if text and text != sentinel:
            out.update(ok=True, chars=len(text), text=text, via=name)
            break
        # 上一步可能点开了图片预览/大图 → 关掉，免得把后面的策略带偏
        send_keys("{ESC}")
        time.sleep(0.8)

    out["shot"] = save_shot(grab(rect), f"{tag}_clip")
    if old_clip:
        _set_clip_text(old_clip)
    if not out["ok"]:
        out["error"] = f"三种复制方式都没拿到内容：{[a['try'] for a in out['attempts']]}"
    return out


def clean_article_text(text: str, max_chars: int = 12000) -> dict:
    """让 LLM 从剪贴板原文里整理出正文段落（去外壳、保原句）。"""
    data = llm_text_json(ARTICLE_CLIP_PROMPT + "\n原始文本：\n" + text[:max_chars], max_tokens=8000)
    if data.get("_error"):
        return {"ok": False, "error": data["_error"]}
    paras = [str(p).strip() for p in (data.get("paragraphs") or []) if str(p or "").strip()]
    raw = str(data.get("_raw") or "")
    salvaged = False
    if not paras and raw:
        paras = _salvage_strings(raw, "paragraphs")  # JSON 被 token 截断也要把正文救回来
        salvaged = bool(paras)
    return {
        "ok": bool(paras),
        "title": data.get("title") or (_salvage_strings(raw, "title") or [None])[0],
        "author": data.get("author") or (_salvage_strings(raw, "author") or [None])[0],
        "publish_time": data.get("publish_time"),
        "paragraphs": paras,
        "chars": sum(len(p) for p in paras),
        "is_image_article": bool(data.get("is_image_article")),
        "dropped": data.get("dropped") or _salvage_strings(raw, "dropped"),
        "salvaged": salvaged,
        "truncated": len(text) > max_chars,
        "raw_chars": len(raw),
    }


def _scroll_to_top(rect: list[int], hwnd: int = 0, max_up: int = 12) -> dict:
    """把文章页滚回顶部，再开始抄。

    坑（2026-11-03 实测）：CEF **记得同一篇文章上次的滚动位置**——上一轮在文章页里滚到底，
    这次点开同一篇直接落在文末，第 0 轮就 `bottom_reached=true`，只抄到文末的二维码/合集目录/
    上一篇（81 字）。所以抄正文之前必须先归位：先 Ctrl+Home，再用滚轮往上滚到画面不再变。
    """
    from pywinauto.keyboard import send_keys

    _ensure_front(hwnd)
    home_ok = False
    try:
        send_keys("^({HOME})")
        home_ok = True
    except Exception:  # noqa: BLE001
        pass
    time.sleep(1.0)

    prev = None
    ups = 0
    for _ in range(max_up):
        img = grab(rect)
        if prev is not None and _img_diff(prev, img) < 1.0:
            break
        prev = img
        scroll_page(rect, notches=-4)
        ups += 1
    return {"home_key": home_ok, "up_scrolls": ups}


_PUNCT = set(
    " \t\r\n，。、；：！？“”‘’（）()《》【】…—-·,.!?:;\'\"|/\\「」『』「％%"
)


def _cmp_chars(s: str) -> list[tuple[str, int]]:
    """列出原文里「可比较的字符」及其在原串中的下标（忽略空白与标点）。"""
    return [(ch.lower(), i) for i, ch in enumerate(s) if ch not in _PUNCT]


def _similar(a: str, b: str) -> bool:
    """两段是不是同一段：归一化相等，或相似度 ≥ 0.85（VLM 两次抄法常差几个字）。"""
    na, nb = _norm_title(a), _norm_title(b)
    if not na or not nb or len(na) < 12 or len(nb) < 12:
        return False
    if na == nb:
        return True
    from difflib import SequenceMatcher

    return SequenceMatcher(None, na, nb, autojunk=False).ratio() >= 0.85


def _overlap_cut(a: str, b: str, min_len: int = 10) -> int:
    """a 的末尾与 b 的开头重叠多少 → 返回 b 里「不重叠部分」的起始下标。

    - `len(b)`：b 已被 a 覆盖（整段重复，丢掉）；
    - `0`：找不到重叠。
    比较时忽略空白与标点，所以两次抄写标点/空格不同也能对上。
    """
    ka, kb = _cmp_chars(a), _cmp_chars(b)
    if not kb:
        return len(b)
    if not ka:
        return 0
    na = "".join(c for c, _ in ka)
    nb = "".join(c for c, _ in kb)
    if nb in na:
        return len(b)
    for k in range(min(len(ka), len(kb), 300), min_len - 1, -1):
        if na[-k:] == nb[:k]:
            return kb[k - 1][1] + 1
    return 0


def _prefix_dupe(a: str, b: str, min_len: int = 12, min_ratio: float = 0.6) -> bool:
    """两段是不是「同一段的两种抄法」：开头一大截一样（共同前缀 ≥ 短段的 60%）。

    实测重复就是这种：跨屏时 VLM 对同一段给了长度不同的两种抄法
    （「…天神之眼B供应中的角色更核心」 vs 「…天神之眼B中供应的不再只是芯片。知情人士称…」）。
    阈值 60% 是为了不误伤「10月销量：小鹏…」/「10月销量：蔚来…」这种同前缀的不同条目
    （它们共同前缀只有 5 个字，远小于短段的 60%）。
    """
    ka, kb = _cmp_chars(a), _cmp_chars(b)
    n = min(len(ka), len(kb))
    if n < min_len:
        return False
    na = "".join(c for c, _ in ka)
    nb = "".join(c for c, _ in kb)
    i = 0
    while i < n and na[i] == nb[i]:
        i += 1
    return i >= max(min_len, int(n * min_ratio))


def _merge_paragraphs(acc: list[str], cur: list[str]) -> dict:
    """把一屏段落并进累计正文：先按段落对齐，再在边界做**字符级**去重。

    旧做法是「整段归一化文本做 key 去重」，跨屏同一句只要 VLM 措辞差一两个字就挡不住
    （实测「与此前双方在天神之眼C上的合作模式不同…」出现两条），所以改成字符级合并。
    """
    cur = [str(p or "").strip() for p in (cur or []) if str(p or "").strip()]
    info: dict = {"paras": len(cur), "matched": 0, "added": 0, "mode": "append"}
    if not cur:
        info["mode"] = "empty"
        return info
    if not acc:
        acc.extend(cur)
        info.update(added=len(cur), mode="first")
        return info

    # 1) 段落级对齐：acc 末尾 s 段 ≈ cur 开头 s 段，取最大的 s
    s = 0
    for cand in range(min(len(acc), len(cur)), 0, -1):
        if all(_similar(acc[len(acc) - cand + i], cur[i]) for i in range(cand)):
            s = cand
            break
    if s:
        head, tail = acc[-1], cur[s - 1]
        cut = _overlap_cut(head, tail)
        if cut >= len(tail):
            pass  # 边界段整段重复 → 保留 acc 里那份
        elif cut:
            acc[-1] = head + tail[cut:]  # 字符级拼接：只接上不重叠的尾巴
        else:
            acc[-1] = max(head, tail, key=len)  # 同一段两种抄法 → 留更长的那份
        keys = {_norm_title(x)[:160] for x in acc}
        rest: list[str] = []
        for p in cur[s:]:
            key = _norm_title(p)[:160]
            if not key or key in keys:
                continue
            keys.add(key)
            rest.append(p)
        acc.extend(rest)
        info.update(matched=s, added=len(rest), mode="paragraph")
        return info

    # 2) 没有整段重叠 → 用累计正文最后一段和这屏前几段做字符级对齐
    tail = acc[-1]
    for j in range(min(3, len(cur))):
        # 先判「同一段两种抄法」：否则整段被覆盖时会把多出来的尾巴拆成单独一段（单测抓到过）
        if _prefix_dupe(tail, cur[j]):
            if len(cur[j]) > len(tail):
                acc[-1] = cur[j]  # 同一段两种抄法 → 留更长的那份
            acc.extend(cur[j + 1 :])
            info.update(matched=j, added=len(cur) - j - 1, mode="same_para")
            return info
        cut = _overlap_cut(tail, cur[j])
        if cut:
            if cut < len(cur[j]):
                acc.append(cur[j][cut:])
            acc.extend(cur[j + 1 :])
            info.update(matched=j, added=len(cur) - j, mode="char")
            return info

    acc.extend(cur)
    info.update(added=len(cur), mode="no_overlap")
    return info


def _try_merge(acc: list[str], cur: list[str]) -> tuple[list[str], dict]:
    """在**副本**上合并，返回 `(结果, 说明)`：接不上时调用方可以先不采纳。"""
    trial = list(acc)
    return trial, _merge_paragraphs(trial, cur)


def collect_article_body(
    rect: list[int],
    hwnd: int = 0,
    max_scroll: int = 16,
    tag: str = "body",
    step_notches: int = 6,
    mode: str = "auto",
) -> dict:
    """取一篇文章的正文。

    `mode`：
    - `auto`（默认）：先走**剪贴板路线**（点正文区空白 → Ctrl+A → Ctrl+C → LLM 整理），
      拿不到可用正文（图片长文、复制没生效）再回退逐屏截图抄；
    - `clipboard`：只走剪贴板；
    - `scroll`：只逐屏截图抄。

    逐屏路线：每屏往下滚 `step_notches` 档，屏与屏之间用 `_merge_paragraphs()` 做**字符级重叠合并**
    （接不上就回滚 3 档重读再试，每轮最多 2 次；实在接不上才整屏收下，计入 `gaps`）。

    逐屏路线只在剪贴板路线不可用时才走（`to_top` 归位、`rounds`、`gaps` 都只属于这条路线）。
    """
    out: dict = {"paragraphs": [], "rounds": [], "stop": None, "source": None}
    acc: list[str] = []
    prev = None
    meta: dict = {}

    # 路线一：剪贴板（点正文区空白 → Ctrl+A → Ctrl+C → LLM 整理）。
    # 一次拿全文：不用跨屏去重、不怕滚动位置、不会误判到底。
    if mode in ("auto", "clipboard"):
        clip = grab_article_clipboard(rect, hwnd=hwnd, tag=tag)
        out["clipboard"] = {
            "ok": clip.get("ok"),
            "chars": clip.get("chars"),
            "via": clip.get("via"),          # no_click / esc_no_click / title_click
            "attempts": clip.get("attempts"),  # 每种策略各拿了几字，排查焦点问题靠它
            "error": clip.get("error"),
        }
        if clip.get("ok"):
            clean = clean_article_text(str(clip.get("text") or ""))
            out["clipboard"]["cleaned"] = {
                "ok": clean.get("ok"),
                "chars": clean.get("chars"),
                "is_image_article": clean.get("is_image_article"),
                "truncated": clean.get("truncated"),
                "dropped": clean.get("dropped"),
                "salvaged": clean.get("salvaged"),
                "raw_chars": clean.get("raw_chars"),
                "error": clean.get("error"),
            }
            if clean.get("ok") and int(clean.get("chars") or 0) >= 120:
                out.update(
                    source="clipboard",
                    paragraphs=clean["paragraphs"],
                    title=clean.get("title"),
                    author=clean.get("author"),
                    publish_time=clean.get("publish_time"),
                    chars=clean["chars"],
                    stop="剪贴板一次取全（Ctrl+A / Ctrl+C）",
                )
                log(f"[正文] 剪贴路线：原文 {clip.get('chars')} 字 → 整理后 {clean['chars']} 字")
                return out
        if mode == "clipboard":
            out["stop"] = "剪贴板路线没拿到可用正文"
            out["chars"] = 0
            return out
        log("[正文] 剪贴路线不可用 → 回退逐屏截图抄")

    # 路线二：逐屏截图抄（图片长文 / 复制被禁用时）
    out["source"] = "scroll"
    # 先归位：不这么做，页面可能停在上一次阅读的位置（实测直接落在文末）
    out["to_top"] = _scroll_to_top(rect, hwnd=hwnd)

    for i in range(0, max_scroll + 1):
        data = read_article_page(rect, f"{tag}_{i}", hwnd=hwnd)
        img = data.get("_img")
        if i == 0:
            meta = {
                "title": data.get("title"),
                "author": data.get("author"),
                "publish_time": data.get("publish_time"),
            }
        # 先在副本上合并：接不上就先不采纳，避免"整屏硬追加"造出重复段落
        trial, info = _try_merge(acc, data.get("paragraphs") or [])
        back_total = 0
        for back in range(1, 3):
            if info["mode"] != "no_overlap" or i == 0:
                break
            # 接不上 = 这一下滚过头、中间漏了一段 → 回滚 3 档重读，再试一次对齐（每轮最多 2 次）
            scroll_page(rect, notches=-3)
            back_total += 3
            again = read_article_page(rect, f"{tag}_{i}r{back}", hwnd=hwnd)
            trial, info = _try_merge(acc, again.get("paragraphs") or [])
            info["retry"] = back
            # 注意：不要用重读到的图去更新 img —— 那是回滚后的位置，
            # 会让下一轮的「滚不动」比对拿错位图当基线（实测第 4 轮被误判成到底）
        if back_total:
            # 把回滚的量补回来：否则下一轮读回同一屏，净进度为 0（实测第 4/5 轮就是这样卡住）
            scroll_page(rect, notches=back_total)
        if info["mode"] == "no_overlap":
            clean = [str(p).strip() for p in (data.get("paragraphs") or []) if str(p or "").strip()]
            acc.extend(clean)  # 回滚两次还是接不上 → 宁可重复也别丢内容
            info["added"] = len(clean)
            info["mode"] = "no_overlap_forced"
        else:
            acc[:] = trial

        out["rounds"].append(
            {
                "round": i,
                "shot": data.get("_shot"),
                "paras": len(data.get("paragraphs") or []),
                "matched": info["matched"],
                "added": info["added"],
                "mode": info["mode"],
                "retry": info.get("retry"),
                "bottom": bool(data.get("bottom_reached")),
                "page": data.get("page"),
            }
        )
        log(
            f"[正文] 第{i}轮：{len(data.get('paragraphs') or [])} 段"
            f"（对齐 {info['matched']}、新增 {info['added']}、{info['mode']}）"
        )

        if data.get("bottom_reached"):
            out["stop"] = "到底了"
            break
        if i > 0 and str(data.get("page") or "") not in ("", "article"):
            out["stop"] = f"页面已不是文章（page={data.get('page')}）"
            break
        if prev is not None and img is not None and _img_diff(prev, img) < 1.0:
            out["stop"] = "滚不动（内容没变化）"
            break
        prev = img
        if i == max_scroll:
            out["stop"] = f"达到 max_scroll={max_scroll}"
            break
        scroll_page(rect, notches=step_notches)

    out["paragraphs"] = acc
    out.update(meta)
    out["chars"] = sum(len(p) for p in acc)
    # 只有"回滚两次还是接不上、硬收下"才算真缺口（no_overlap_forced）
    out["gaps"] = sum(1 for r in out["rounds"] if r.get("mode") == "no_overlap_forced")
    return out


def close_article_tab(hwnd: int, rect: list[int], tag: str = "tabclose") -> dict:
    """关掉「文章」那个标签页，回到公众号主页标签。

    坑（2026-11-03 实测）：搜一搜窗口是**浏览器式多标签**（顶部一排标签 + ×），
    点开文章其实是**新开一个标签页**（同一个 HWND，所以 `classify_search_page` 只看到
    page 从 account_home 变 article，被误判成"同窗口跳转"）。这种标签页里没有上一页，
    `Alt+Left` 按 3 次仍停在文章里；**`Ctrl+W` 一次就回到 account_home**。
    """
    from pywinauto.keyboard import send_keys

    _ensure_front(hwnd)
    time.sleep(0.4)
    cur = read_feed_page(rect, f"{tag}_pre", hwnd=hwnd)
    if str(cur.get("page")) != "article":
        # 已经不在文章页（中途被带走/自己退回了）→ 千万别按 Ctrl+W，那会关掉主页标签
        return {"via": "skip(不在文章页)", "page": cur.get("page")}
    send_keys("^w")
    time.sleep(2.2)
    _ensure_front(hwnd)
    page = read_feed_page(rect, f"{tag}", hwnd=hwnd)
    return {"via": "ctrl+w", "page": page.get("page")}


def click_article(
    rect: list[int], hwnd: int, title: str, max_scroll: int = 0, allow_scroll: bool = False
) -> dict:
    """在公众号主页的文章列表里找到标题对应的那篇并点开。

    定位沿用「卡片带序号优先、模型 y 次之、模型坐标兜底」三种策略轮换；
    点开后立刻判：冒出新的 CEF 窗口（文章独立窗口）或当前页变 article 才算成功。

    **默认不滚屏**（2026-11-03 改）：步骤 5 默认只取当前屏（`--max-scroll 0`），清单里那篇就在
    当前屏；找不到是识别问题，往下滚只会把列表推走、下一篇又找不到（实测第六步因此
    「一步步滚屏」连滚 10 轮）。要翻历史内容显式给 `allow_scroll=True` + `max_scroll=N`。
    """
    out: dict = {"title": title, "attempts": [], "ok": False}
    before = cef_hwnds(title_filter=False)
    if allow_scroll:
        # 清单是跨屏收集的（步骤 5 滚过 N 屏），列表此刻停在最后一屏 → 先归位，再逐屏往下找
        out["to_top"] = _scroll_to_top(rect, hwnd=hwnd)

    for i in range(0, max_scroll + 1):
        data = read_feed_page(rect, f"find_{i}", hwnd=hwnd)
        items = [x for x in (data.get("items") or []) if isinstance(x, dict)]
        img = data.get("_img")
        crect = data.get("_rect") or rect
        rank = next(
            (k for k, it in enumerate(items) if _title_match(it.get("title"), title)), None
        )
        if rank is None:
            cur_page = str(data.get("page") or "")
            out["attempts"].append(
                {
                    "round": i,
                    "found": False,
                    "page": cur_page,
                    "visible": [str(x.get("title") or "")[:24] for x in items[:4]],
                }
            )
            if cur_page == "article":
                # 其实已经点开了（只是标题没对上/模型没列卡片）→ **绝不能再往下滚**：
                # 那会把文章滚到底，CEF 记住位置后下次点开直接落在文末（§8.8 实测坑）
                out.update(
                    ok=True,
                    via="already_article",
                    hwnd=int(hwnd),
                    rect=rect,
                    page="article",
                    page_title=data.get("title"),
                    title_matched=False,
                    shot=data.get("_shot"),
                )
                return out
            if not allow_scroll:
                out["error"] = (
                    "当前屏没找到这篇（步骤5 只取当前屏，不滚屏找；"
                    "要翻历史内容加 --article-find-scroll N）"
                )
                return out
            scroll_page(rect)
            continue

        for strat in ("band", "near", "vlm"):
            # card_x=True：主页的文章卡片是内缩的，x 取卡片自身中心（固定左偏移会点到空白槽）
            pt, how = pick_click_point(
                img, crect, items[rank].get("center_norm"), rank, strat, card_x=True
            )
            if pt is None:
                continue
            click_screen(pt[0], pt[1])
            time.sleep(2.6)
            new = _pick_new_cef(before)
            if new is not None:
                _ensure_front(new["hwnd"])
                time.sleep(0.4)
                page = classify_search_page(new["rect"], f"art_{i}_{how}")
                ok = str(page.get("page")) == "article" or _title_match(page.get("title"), title)
                out["attempts"].append(
                    {
                        "round": i,
                        "rank": rank,
                        "how": how,
                        "point": list(pt),
                        "page": page.get("page"),
                        "page_title": page.get("title"),
                        "new_window": new["hwnd"],
                    }
                )
                if ok:
                    out.update(
                        ok=True,
                        via="new_window",
                        hwnd=int(new["hwnd"]),
                        rect=new["rect"],
                        page=page.get("page"),
                        page_title=page.get("title"),
                        title_matched=bool(_title_match(page.get("title"), title)),
                        shot=page.get("shot"),
                    )
                    return out
                # 认错窗口（比如点了别的卡片）→ 关掉新窗口，接着试下一种策略
                close_cef_window(int(new["hwnd"]))
                _ensure_front(hwnd)
                continue

            page = classify_search_page(rect, f"art_{i}_{how}")
            out["attempts"].append(
                {
                    "round": i,
                    "rank": rank,
                    "how": how,
                    "point": list(pt),
                    "page": page.get("page"),
                    "page_title": page.get("title"),
                }
            )
            if str(page.get("page")) == "article":
                # 页面已是文章页就算点开成功：**不能用窗口标题判定是不是这篇** ——
                # 实测微信文章页的窗口标题常常是公众号名（"建约车评"）或"合集·目录"，
                # 拿它做硬匹配会把成功当失败（然后一路滚着找，越找越偏）。
                # 标题对不对交给正文抓取阶段核对（collect_article_body 会读自己的 title）。
                out.update(
                    ok=True,
                    via="same_window",
                    hwnd=int(hwnd),
                    rect=rect,
                    page="article",
                    page_title=page.get("title"),
                    title_matched=bool(_title_match(page.get("title"), title)),
                    shot=page.get("shot"),
                )
                return out

    out["error"] = "没能在文章列表里点开这篇"
    return out


def load_step5_items(feed_json: str = "") -> list[dict]:
    """读步骤 5 落盘的文章清单（默认 out/step5_feed.json）。"""
    path = Path(feed_json) if feed_json else (OUT_DIR / "step5_feed.json")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    collected = data.get("collected") or {}
    if "文章" in collected:
        # 「文章」分项存在但为空 = 这段时间没有文章，绝不能拿别的分项（视频号）顶替
        return [x for x in ((collected.get("文章") or {}).get("items") or []) if isinstance(x, dict)]
    for val in collected.values():
        items = [x for x in ((val or {}).get("items") or []) if isinstance(x, dict)]
        if items:
            return items
    return []


def load_step5_meta(feed_json: str = "") -> dict:
    """步骤 5 落盘时的窗口参数（`days` / `max_scroll`）。

    步骤 6 要靠 `max_scroll` 决定「清单是跨几屏收集的、找文章最多要往下滚几屏」——
    实测 `--days 2 --max-scroll 4` 收了 4 篇（列表已滚到第 4 屏），步骤 6 不滚屏就只
    能点到当前屏那 2 篇，后 2 篇报「当前屏没找到这篇」。
    """
    path = Path(feed_json) if feed_json else (OUT_DIR / "step5_feed.json")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return {k: data.get(k) for k in ("days", "max_scroll", "scrolled", "tabs") if k in data}


def step6_collect_articles(
    items: list[dict] | None = None,
    limit: int = 5,
    max_scroll: int = 16,
    dry_run: bool = False,
    feed_json: str = "",
    step_notches: int = 6,
    article_mode: str = "clipboard",
    find_scroll: int = -1,
) -> dict:
    """步骤6：拿步骤5的文章清单，逐篇点开抄正文，抄完把文章关掉/回到主页。

    打开方式实测有两种，收尾手段不同：
    - 搜一搜窗口内**新开标签页**（同一个 HWND，page 从 account_home 变 article）
      → 抓完用 `Ctrl+W` 关掉该标签页（`Alt+Left` 在新标签里没有上一页，退不回去）；
    - 另开独立 CEF 窗口（标题=文章标题）→ 抓完 `WM_CLOSE` 关掉，主页窗口留在原处。

    抄之前一律先 `_scroll_to_top()` 归位（CEF 记得上次的滚动位置，不归位会从文末开始抄）。
    """
    if items is None:
        items = load_step5_items(feed_json)
    res: dict = {
        "stage": "6_article",
        "limit": limit,
        "max_scroll": max_scroll,
        "articles": [],
        "events": [],
    }
    if find_scroll < 0:
        # 没显式指定就跟着步骤 5 的滚动深度走：清单是跨几屏收的，找文章最多也滚几屏。
        # 实测 --days 2 --max-scroll 4 收了 4 篇，步骤6 不滚屏只能点到当前屏那 2 篇。
        meta5 = load_step5_meta(feed_json)
        # scrolled = 步骤5 实际滚了几屏（含 --feed-fill-window 自动补滚），优先用它
        find_scroll = int(
            meta5.get("scrolled")
            if meta5.get("scrolled") is not None
            else meta5.get("max_scroll") or 0
        )
        res["feed_meta"] = meta5
        res["find_scroll"] = find_scroll

    def ev(kind: str, **kw) -> None:
        item = {"t": f"{datetime.now():%H:%M:%S}", "event": kind}
        item.update(kw)
        res["events"].append(item)
        log(f"{kind} {kw if kw else ''}")

    titles = [str(x.get("title") or "") for x in (items or []) if str(x.get("title") or "").strip()]
    res["titles"] = titles[:limit]
    if not titles:
        res.update(
            result="error",
            error=(
                "没有文章清单：步骤5 的「文章」分项在当前 --days 范围内没内容"
                "（加大 --days 重跑步骤5，或用 --feed-json 指定别的清单）"
            ),
        )
        return res

    win = find_main_window(3.0)
    main_hwnd = window_hwnd(win) if win is not None else 0
    if win is None:
        res["warning"] = "没找到微信主窗（可能缩在托盘），只按搜一搜窗口继续"
    swin = find_search_window(3.0, main_hwnd=main_hwnd)
    res["search_window"] = swin
    if swin is None:
        res.update(result="error", error="没有搜一搜窗口（先跑步骤 4/5 打开公众号主页）")
        return res
    rect = swin["rect"]
    fh = int(swin["hwnd"])
    _ensure_front(fh)

    page = read_feed_page(rect, "art_pre", hwnd=fh)
    res["account"] = page.get("title")
    if str(page.get("page") or "") != "account_home":
        res.update(result="error", error=f"当前不是公众号主页（page={page.get('page')}）")
        return res
    if dry_run:
        res.update(result="ok", dry_run=True, note="只确认在公众号主页，未点开文章")
        return res

    clk = click_feed_tab(rect, "文章", hwnd=fh)
    res["tab"] = {k: clk.get(k) for k in ("ok", "already", "attempt", "active", "error")}
    ev("切到文章列表", ok=clk.get("ok"), active=clk.get("active"))
    # 回到列表顶部，保证"前几篇"就是最新的几篇
    scroll_page(rect, -8)
    time.sleep(0.8)

    ok_count = 0
    for idx, title in enumerate(titles[:limit]):
        rec: dict = {"index": idx, "title": title}
        if idx:
            # 上一篇没退回主页就别硬点（会在文章页里乱点/乱滚，还会污染滚动位置）
            cur = read_feed_page(rect, f"pre_{idx}", hwnd=fh)
            if str(cur.get("page")) != "account_home":
                res["error"] = (
                    f"第 {idx} 篇开始前不在公众号主页（page={cur.get('page')}），"
                    "已中断；重跑步骤 4 打开主页后再来"
                )
                ev("中断", index=idx, page=cur.get("page"))
                break
        ev("打开文章", index=idx, title=title[:30])
        hit = click_article(
            rect, fh, title, max_scroll=find_scroll, allow_scroll=find_scroll > 0
        )
        rec["open"] = {
            k: hit.get(k)
            for k in ("ok", "via", "page", "page_title", "title_matched", "hwnd", "attempts")
        }
        if not hit.get("ok"):
            rec["ok"] = False
            rec["error"] = hit.get("error") or "没点开"
            res["articles"].append(rec)
            ev("打开失败", index=idx, error=rec["error"])
            # 兜底：点歪了也可能停在文章页 → 一样要把那个标签页关掉，别留着
            _ensure_front(fh)
            cur = read_feed_page(rect, f"recover_{idx}", hwnd=fh)
            if str(cur.get("page")) == "article":
                rec["close"] = close_article_tab(fh, rect, tag=f"tabclose_fail_{idx}")
            elif str(cur.get("page")) != "account_home":
                _send_back()
                time.sleep(1.8)
            continue

        read_rect = hit.get("rect") or rect
        read_hwnd = int(hit.get("hwnd") or fh)
        body = collect_article_body(
            read_rect,
            hwnd=read_hwnd,
            max_scroll=max_scroll,
            tag=f"a{idx}",
            step_notches=step_notches,
            mode=article_mode,
        )
        rec.update(
            ok=True,
            via=hit.get("via"),
            page_title=body.get("title") or hit.get("page_title"),
            author=body.get("author"),
            publish_time=body.get("publish_time"),
            chars=body.get("chars"),
            paragraphs=body.get("paragraphs"),
            rounds=body.get("rounds"),
            stop=body.get("stop"),
            to_top=body.get("to_top"),
            gaps=body.get("gaps"),
            source=body.get("source"),
            clipboard=body.get("clipboard"),
        )

        # 关掉/退回：独立窗口直接关；搜一搜里点开文章是**新开标签页** → 用 Ctrl+W 关标签
        if hit.get("via") == "new_window":
            rec["close"] = close_cef_window(read_hwnd)
        else:
            rec["close"] = close_article_tab(fh, rect, tag=f"tabclose_{idx}")
            if str(rec["close"].get("page")) != "account_home":
                _send_back()
                time.sleep(1.8)
                rec["close"]["alt+left"] = read_feed_page(rect, f"back_{idx}_alt", hwnd=fh).get(
                    "page"
                )
        _ensure_front(fh)
        back = read_feed_page(rect, f"back_{idx}", hwnd=fh)
        rec["back_page"] = back.get("page")
        rec["back_tries"] = 0
        for t in range(1, 4):  # 最多再退 3 次：Alt+Left 有时退到搜索页/中间页（实测 page=other）
            if str(back.get("page")) == "account_home":
                break
            _send_back()
            time.sleep(1.8)
            _ensure_front(fh)
            back = read_feed_page(rect, f"back_{idx}_{t}", hwnd=fh)
            rec["back_page"] = back.get("page")
            rec["back_tries"] = t
        if str(rec["back_page"]) != "account_home":
            rec["needs_reopen"] = True
            ev("退回主页失败", index=idx, page=rec["back_page"], tries=rec["back_tries"])
        ok_count += 1
        res["articles"].append(rec)
        ev(
            "文章已抓",
            index=idx,
            chars=rec.get("chars"),
            stop=rec.get("stop"),
            back=rec.get("back_page"),
        )

    res["total"] = len(res["articles"])
    res["ok_count"] = ok_count
    res["result"] = "ok" if ok_count else ("partial" if res["articles"] else "empty")
    if not ok_count:
        res["error"] = res.get("error") or "一篇文章都没抓成功"
    return res


# --------------------------------------------------------------------------
# 流程编排：每步先做前置判定，不满足就回退上一步重做
# --------------------------------------------------------------------------
def precheck_screen() -> dict:
    """步骤2 前置：主窗存在、可见、没最小化、完整落在屏幕内（哪块屏都行）。"""
    found = find_main_hwnd()
    if found is None:
        return {"ok": False, "why": "没找到微信主窗（未启动或还停在登录窗）", "recoverable": False}
    out = {
        "hwnd": found["hwnd"],
        "rect": found["rect"],
        "visible": found["visible"],
        "iconic": found["iconic"],
    }
    if not found["visible"]:
        return dict(out, ok=False, why="主窗不可见（可能在托盘或被隐藏）")
    if found["iconic"]:
        return dict(out, ok=False, why="主窗最小化")
    ratio = onscreen_ratio(found["rect"])
    out["onscreen_ratio"] = round(ratio, 3)
    out["on_primary_ratio"] = round(intersects_ratio(found["rect"], primary_monitor_rect()), 3)
    if ratio < 0.98:
        return dict(out, ok=False, why=f"主窗只有 {ratio:.0%} 在屏幕内（按窗口坐标点会落空）")
    return dict(out, ok=True)


def precheck_chat() -> dict:
    """步骤3 前置：主窗就绪 + 会话列表有条目 + 没有被附属窗口盖住。"""
    base = precheck_screen()
    if not base["ok"]:
        return base
    if wechat_hung():
        return dict(base, ok=False, why="微信窗口无响应", recoverable=False)
    win = find_main_window(2.0)
    if win is None:
        diag = a11y_diagnosis()
        base["a11y"] = diag
        return dict(
            base,
            ok=False,
            recoverable=False,
            why=f"UIA 拿不到主窗：{diag.get('cause')}",
            fix=diag.get("fix"),
        )
    n = session_count(win)
    base["session_items"] = n
    if n <= 0:
        diag = a11y_diagnosis()
        base["a11y"] = diag
        return dict(
            base,
            ok=False,
            recoverable=False,
            why=f"会话列表为空：{diag.get('cause')}",
            fix=diag.get("fix"),
        )
    sides = close_side_windows(base["hwnd"], dry_run=True)
    if sides:
        base["side_windows"] = [s["title"] for s in sides]
        return dict(base, ok=False, why="有附属窗口盖在前面（设置/聊天文件/文章详情等）")
    return dict(base, ok=True)


def precheck_search() -> dict:
    """步骤4 前置：聊天页就绪 + 会话列表里能看到「公众号」。"""
    base = precheck_chat()
    if not base["ok"]:
        return base
    win = find_main_window(2.0)
    item, names = find_session_item_deep(win, "公众号")
    base["has_gzh"] = item is not None
    base["session_names"] = names[:10]
    if item is None:
        # 回退上一步也造不出这个会话，属于硬阻塞，直接停
        return dict(base, ok=False, recoverable=False, why="会话列表里没有「公众号」会话")
    return dict(base, ok=True)


def precheck_feed(days: int = 1, tabs: tuple[str, ...] = ("文章",)) -> dict:
    """步骤5 前置：搜一搜窗口还开着，且当前页面就是公众号主页。

    主页样式（顶部公众号信息 + 一排分项 + 按日期分组卡片流）是步骤4 点开的结果，
    所以这里只读一次页面，不满足就回退步骤4 重开（recoverable=True）。
    """
    base: dict = {"days": days, "wanted_tabs": list(tabs)}
    win = find_main_window(3.0)
    if win is not None and wechat_hung():
        return dict(base, ok=False, recoverable=False, why="微信窗口无响应")
    main_hwnd = window_hwnd(win) if win is not None else 0
    base["main_window"] = main_hwnd or None
    if win is None:
        base["warning"] = "没找到微信主窗（可能缩在托盘），只按搜一搜窗口判定"
    swin = find_search_window(3.0, main_hwnd=main_hwnd)
    if swin is None:
        return dict(base, ok=False, why="搜一搜窗口已关闭（公众号主页没打开）")
    base["search_window"] = {"title": swin.get("title"), "rect": swin.get("rect")}
    page = read_feed_page(swin["rect"], "precheck", hwnd=int(swin["hwnd"]))
    base["page"] = page.get("page")
    base["title"] = page.get("title")
    base["shot"] = page.get("_shot")
    if str(page.get("page") or "") != "account_home":
        return dict(base, ok=False, why=f"当前不是公众号主页（page={page.get('page')}）")
    names = {str(t.get("name") or "") for t in (page.get("tabs") or [])}
    base["tabs_visible"] = sorted(n for n in names if n in TAB_ORDER)
    if not names & set(TAB_ORDER):
        return dict(base, ok=False, why="主页上没找到分项标签（全部/贴图/文章/视频号）")
    base["days"] = days
    base["wanted_tabs"] = list(tabs)
    return dict(base, ok=True)


def run_flow(
    keyword: str,
    open_first: bool = True,
    dry_run: bool = False,
    max_runs: int = 2,
    days: int = 1,
    tabs: tuple[str, ...] = ("文章",),
    max_scroll: int = 0,
    video_limit: int = VIDEO_TAB_LIMIT,
    fill_window: bool = False,
) -> dict:
    """步骤2→3→4→5 连跑。

    规则：每一步执行前先跑前置判定；前置不满足 → 回退上一步重做
    （2 的前置不满足或属于硬阻塞 → 直接停）。每步最多重试 max_runs 次，
    整个流程最多回退 10 次，防止在两个步骤之间来回横跳。
    """
    res: dict = {"stage": "flow", "keyword": keyword, "events": [], "steps": []}

    def ev(kind: str, **kw) -> None:
        item = {"t": f"{datetime.now():%H:%M:%S}", "event": kind}
        item.update(kw)
        res["events"].append(item)
        log(f"{kind} {kw if kw else ''}")

    plan = [
        {
            "name": "2_window_ready",
            "pre": precheck_screen,
            "run": lambda: step2_window_ready(dry_run=dry_run),
            "ok": lambda r: r.get("result") == "ok",
        },
        {
            "name": "3_chat",
            "pre": precheck_chat,
            "run": lambda: step3_main_chat(dry_run=dry_run),
            "ok": lambda r: r.get("result") == "ok",
        },
        {
            "name": "4_search",
            "pre": precheck_search,
            "run": lambda: step4_search_account(
                keyword, open_first=open_first, dry_run=dry_run
            ),
            "ok": lambda r: r.get("result") == "ok",
        },
        {
            "name": "5_feed",
            "pre": lambda: precheck_feed(days, tabs),
            "run": lambda: step5_collect_feed(
                days=days,
                tabs=tabs,
                max_scroll=max_scroll,
                dry_run=dry_run,
                video_limit=video_limit,
                fill_window=fill_window,
            ),
            # 采集本身失败（某分项没切进去）不算流程失败，只要读到主页就算通过
            "ok": lambda r: r.get("result") in ("ok", "partial"),
        },
    ]

    idx = 0
    runs = [0] * len(plan)
    guard = 0
    while idx < len(plan):
        guard += 1
        if guard > 20:
            res.update(result="error", error="流程反复回退超过 20 次，已中止")
            return res

        st = plan[idx]
        pre = st["pre"]()
        if not pre.get("ok"):
            ev(
                "前置判定不满足",
                step=st["name"],
                why=pre.get("why"),
                recoverable=pre.get("recoverable", True),
            )
            if idx == 0 or not pre.get("recoverable", True):
                res.update(
                    result="blocked",
                    error=f"{st['name']} 前置不满足：{pre.get('why')}",
                    precheck=pre,
                )
                return res
            idx -= 1
            continue

        runs[idx] += 1
        if runs[idx] > max_runs:
            res.update(result="error", error=f"{st['name']} 已重试 {max_runs} 次仍未通过")
            return res

        ev("执行步骤", step=st["name"], attempt=runs[idx])
        out = st["run"]()
        res["steps"].append(
            {
                "step": st["name"],
                "attempt": runs[idx],
                "result": out.get("result"),
                "error": out.get("error"),
                "precheck": pre,
                "detail": out,
            }
        )
        if st["ok"](out):
            ev("步骤通过", step=st["name"])
            idx += 1
            continue

        ev("步骤失败，回退上一步", step=st["name"], error=out.get("error"))
        if idx == 0:
            res.update(
                result="error",
                error=f"{st['name']} 失败且没有上一步可退：{out.get('error')}",
            )
            return res
        idx -= 1

    res["result"] = "ok"
    res["summary"] = [
        {"step": s["step"], "attempt": s["attempt"], "result": s["result"], "error": s["error"]}
        for s in res["steps"]
    ]
    return res


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def dump(name: str, obj: dict) -> None:
    path = OUT_DIR / f"{name}.json"
    try:
        path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
        log(f"结果已写入 {path}")
    except OSError as exc:
        log(f"写结果失败: {exc}")


def main() -> int:
    _setup_console()
    ap = argparse.ArgumentParser(description="微信桌面端主流程（步骤化，LLM 视觉决策）")
    ap.add_argument(
        "--stage",
        default="all",
        choices=[
            "all",
            "flow",
            "launch",
            "ready",
            "screen",
            "chat",
            "search",
            "feed",
            "article",
            "a11y",
            "probe",
        ],
    )
    ap.add_argument("--chat-rounds", type=int, default=4, help="步骤3 最多尝试几轮回到主聊天窗")
    ap.add_argument("--keyword", default="", help="步骤4 要搜索的公众号关键词")
    ap.add_argument("--days", type=int, default=1, help="步骤5 只保留最近 N 天的内容")
    ap.add_argument(
        "--tabs",
        default="文章",
        help="步骤5 要采集的分项，逗号分隔（默认只采「文章」；视频号要显式写 --tabs 文章,视频号）",
    )
    ap.add_argument(
        "--max-scroll",
        type=int,
        default=0,
        help="步骤5 每个分项最多下滚几次（默认 0=只读当前屏，不滚动）",
    )
    ap.add_argument(
        "--feed-fill-window",
        action="store_true",
        help="步骤5 时间窗口没闭合（最后一屏最下面一条仍在窗口内）时自动多滚一屏，"
        f"最多在 --max-scroll 之外补 {FILL_WINDOW_EXTRA} 屏。实测 --days 1 只读当前屏会漏："
        "今天有 2 篇、当前屏只显示 1 篇",
    )
    ap.add_argument(
        "--video-limit",
        type=int,
        default=VIDEO_TAB_LIMIT,
        help="步骤5 采「视频号」时只取最前面几条（该分项卡片没有日期分组）",
    )
    ap.add_argument(
        "--fix-a11y",
        action="store_true",
        help="--stage a11y 时顺手把系统读屏标志置 1（对照用；11-04 复测确认不是根因，见 §11.30）",
    )
    ap.add_argument(
        "--hot-gate",
        action="store_true",
        help="--stage a11y 时热激活 Qt accessibility gate（写 Weixin.dll 的 gate 字节=1），"
        "不重启微信就让 mmui 控件树物化——实测有效，优先用这个",
    )
    ap.add_argument(
        "--repair-a11y",
        action="store_true",
        help="--stage a11y 时先热激活，失败才退化成「清残留插件进程 → 杀微信 → 重启 → 复查」"
        "（会关掉你正在用的微信；11-05 实测重启本身并不能让树出现，仅作最后手段）",
    )
    ap.add_argument(
        "--no-open",
        action="store_false",
        dest="open_first",
        default=True,
        help="步骤4 只搜不点开（默认会在搜一搜窗口里找到对应公众号并点开主页）",
    )
    ap.add_argument(
        "--article-limit",
        type=int,
        default=5,
        help="步骤6 最多抓几篇正文（按步骤5清单的顺序）",
    )
    ap.add_argument(
        "--article-max-scroll",
        type=int,
        default=16,
        help="步骤6 每篇文章最多往下滚几次（抄正文用）",
    )
    ap.add_argument(
        "--article-scroll-notches",
        type=int,
        default=6,
        help="步骤6 每屏往下滚几档（默认 6；越大越快，但两屏间接不上要靠重叠合并补）",
    )
    ap.add_argument(
        "--article-mode",
        choices=("clipboard", "auto", "scroll"),
        default="clipboard",
        help="步骤6 取正文方式：clipboard（默认）=点正文区空白 → Ctrl+A → Ctrl+C → LLM 整理，"
        "全程不滚屏；auto=剪贴板失败再逐屏截图（图片长文用）；scroll=只逐屏截图",
    )
    ap.add_argument(
        "--article-find-scroll",
        type=int,
        default=-1,
        help="步骤6 找不到文章标题时最多往下滚几屏。默认 -1=自动跟随步骤5 的 --max-scroll"
        "（清单跨几屏收的，找之前就先把列表归位、最多滚同样几屏）；0=强制不滚",
    )
    ap.add_argument(
        "--feed-json",
        default="",
        help="步骤6 使用的清单文件，默认 out/step5_feed.json",
    )
    ap.add_argument(
        "--with-article",
        action="store_true",
        help="all 模式下步骤5 之后继续跑步骤6（默认不跑，避免自动点开文章）",
    )
    ap.add_argument("--scan-timeout", type=float, default=180.0, help="等待扫码登录的秒数")
    ap.add_argument("--poll", type=float, default=3.0, help="扫码轮询间隔秒数")
    ap.add_argument("--relogin", type=int, default=1, help="超时/报错后关闭登录窗重来的次数")
    ap.add_argument("--wechat-exe", default="", help="手动指定 Weixin.exe 路径")
    ap.add_argument("--dry-run", action="store_true", help="只判定，不启动不点击不移动")
    args = ap.parse_args()

    summary: dict = {"started_at": f"{datetime.now():%Y-%m-%d %H:%M:%S}"}

    if args.stage == "flow" and not args.keyword:
        print(
            json.dumps(
                {"result": "error", "error": "--stage flow 需要 --keyword（步骤2→3→4→5 连跑）"},
                ensure_ascii=False,
                indent=2,
            )
        )
        return 1

    tabs = tuple(t.strip() for t in args.tabs.split(",") if t.strip())

    if args.stage == "flow":
        f = run_flow(
            keyword=args.keyword,
            open_first=args.open_first,
            dry_run=args.dry_run,
            days=args.days,
            tabs=tabs,
            max_scroll=args.max_scroll,
            video_limit=args.video_limit,
            fill_window=args.feed_fill_window,
        )
        dump("flow", f)
        summary["flow"] = {
            "result": f.get("result"),
            "error": f.get("error"),
            "steps": f.get("summary") or [],
        }
        summary["finished_at"] = f"{datetime.now():%Y-%m-%d %H:%M:%S}"
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return 0 if f.get("result") == "ok" else 1

    if args.stage == "a11y":
        # 无障碍体检：读屏标志 / MMUI 子窗第一层子节点 / 会话条目
        wake_info: dict | None = None
        if args.hot_gate:
            wake_info = wake_a11y(force=True)
            log(
                f"热激活 gate：ok={wake_info.get('ok')} ready={wake_info.get('ready')} "
                f"attempts={wake_info.get('attempts')}"
            )
        repair_info: dict | None = None
        if args.repair_a11y:
            # 重启是最后手段：先热激活（不关微信）。11-05 实测「清残留插件进程 + 重启」完全没用
            # （清干净后连续 3 次重启，45 秒连读 0/45 轮可读），别再拿它当修复动作。
            if wake_info is None:
                wake_info = wake_a11y()
                log(f"先试热激活：ok={wake_info.get('ok')} ready={wake_info.get('ready')}")
            if wake_info.get("ready"):
                log("热激活已让 mmui 树可用，跳过重启")
            else:
                repair_info = repair_a11y_tree(exe=args.wechat_exe, dry_run=args.dry_run)
                log(
                    f"恢复动作完成 ok={repair_info.get('ok')} "
                    f"plugins={repair_info.get('plugins')} killed_weixin={repair_info.get('killed_weixin')} "
                    f"top_windows={repair_info.get('top_windows')}"
                )
                if repair_info.get("ok"):
                    st = step1_launch(
                        scan_timeout=args.scan_timeout,
                        poll=args.poll,
                        relogin=args.relogin,
                        exe=args.wechat_exe,
                    )
                    repair_info["step1"] = {
                        k: st.get(k) for k in ("result", "state", "session_items", "rect")
                    }
        diag = a11y_diagnosis()
        if wake_info is not None:
            diag["wake_a11y"] = wake_info
        if repair_info is not None:
            diag["repair"] = repair_info
        if args.fix_a11y:
            diag["flag_fix"] = ensure_screen_reader_flag()
            diag["after"] = a11y_diagnosis()
            diag["note"] = (
                "读屏标志不是根因（11-04 复测：启动前置 1 仍没树）。verdict 是 a11y_disabled 就跑 "
                "--stage a11y --hot-gate（热写 Qt accessibility gate 字节，不用重启微信）"
            )
        print(json.dumps(diag, ensure_ascii=False, indent=2))
        return 0 if diag.get("verdict") == "ok" else 1

    if args.stage == "probe":
        # 环境体检：exe 能否定位、进程/窗口/显示器实况 + 一次视觉判定（不点不动）
        env = {
            "exe": find_wechat_exe(),
            "pids": wechat_pids(),
            "windows": wechat_top_windows(),
            "monitors": monitors(),
            "primary_work": primary_work_area(),
            "primary_monitor": primary_monitor_rect(),
            "virtual": virtual_screen_rect(),
            "vl": {
                "base": VL_BASE,
                "model": VL_MODEL,
                "thinking": VL_THINKING,
                "key_set": bool(VL_KEY),
                "env_file": str(ENV_FILE) if ENV_FILE.exists() else None,
                "env_keys_loaded": LOADED_ENV_KEYS,
            },
        }
        area = primary_monitor_rect()
        img = grab(area)
        path = save_shot(img, f"probe_{datetime.now():%H%M%S}")
        data = llm_json(img, REVEAL_PROMPT)
        data["_area"] = area
        data["_point"] = action_point(data, area)
        print(json.dumps({"env": env, "shot": path, "llm": data}, ensure_ascii=False, indent=2))
        return 0

    if args.stage in ("all", "launch"):
        s1 = step1_launch(
            scan_timeout=args.scan_timeout,
            poll=args.poll,
            relogin=args.relogin,
            exe=args.wechat_exe,
            dry_run=args.dry_run,
        )
        dump("step1_launch", s1)
        summary["step1"] = {
            "result": s1.get("result"),
            "state": s1.get("state"),
            "rect": s1.get("rect"),
            "error": s1.get("error"),
            "fix": s1.get("fix"),
        }
        if args.stage == "launch":
            print(json.dumps(summary, ensure_ascii=False, indent=2))
            return 0 if s1.get("result") in ("already_logged_in", "logged_in") else 1
        if s1.get("result") not in ("already_logged_in", "logged_in"):
            print(json.dumps(summary, ensure_ascii=False, indent=2))
            return 1

    if args.stage in ("all", "ready", "screen"):
        s2 = step2_window_ready(dry_run=args.dry_run)
        dump("step2_ready", s2)
        summary["step2"] = {
            "result": s2.get("result"),
            "rect_before": s2.get("rect_before"),
            "rect_after": s2.get("rect_after"),
            "onscreen_after": s2.get("onscreen_after"),
            "on_primary_after": s2.get("on_primary_after"),
            "session_items": s2.get("session_items"),
            "session_items_after": s2.get("session_items_after"),
            "error": s2.get("error"),
            "fix": s2.get("fix"),
        }
        if s2.get("result") != "ok":
            print(json.dumps(summary, ensure_ascii=False, indent=2))
            return 1

    if args.stage in ("all", "chat"):
        s3 = step3_main_chat(max_rounds=args.chat_rounds, dry_run=args.dry_run)
        dump("step3_chat", s3)
        summary["step3"] = {
            "result": s3.get("result"),
            "rounds_used": s3.get("rounds_used"),
            "session_items_after": s3.get("session_items_after"),
            "error": s3.get("error"),
        }
        if s3.get("result") != "ok":
            print(json.dumps(summary, ensure_ascii=False, indent=2))
            return 1

    run_search = args.stage == "search" or (args.stage == "all" and bool(args.keyword))
    if run_search:
        if not args.keyword:
            print(
                json.dumps(
                    {"result": "error", "error": "步骤4 需要 --keyword，例如 --keyword 建约车评"},
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 1
        s4 = step4_search_account(
            keyword=args.keyword,
            open_first=args.open_first,
            dry_run=args.dry_run,
        )
        dump("step4_search", s4)
        summary["step4"] = {
            "result": s4.get("result"),
            "keyword": s4.get("keyword"),
            "text_verified": s4.get("text_verified"),
            "search_box_value": s4.get("search_box_value"),
            "results_source": s4.get("results_source"),
            "results": [
                {"name": r.get("name"), "center": r.get("center")}
                for r in (s4.get("results") or [])[:10]
            ],
            "opened": {
                k: (s4.get("open_account") or {}).get(k)
                for k in ("ok", "page", "title", "rounds", "error")
            },
            "error": s4.get("error"),
        }
        if s4.get("result") != "ok" or args.stage == "search":
            summary["finished_at"] = f"{datetime.now():%Y-%m-%d %H:%M:%S}"
            print(json.dumps(summary, ensure_ascii=False, indent=2))
            return 0 if s4.get("result") == "ok" else 1

    # 步骤5：只有在公众号主页已经打开时才自动跟跑（all 模式下 --no-open 不做采集）
    run_feed = args.stage == "feed" or (
        args.stage == "all" and bool(args.keyword) and args.open_first and not args.dry_run
    )
    if run_feed:
        s5 = step5_collect_feed(
            days=args.days,
            tabs=tabs,
            max_scroll=args.max_scroll,
            dry_run=args.dry_run,
            video_limit=args.video_limit,
            fill_window=args.feed_fill_window,
        )
        dump("step5_feed", s5)
        summary["step5"] = {
            "result": s5.get("result"),
            "account": s5.get("account"),
            "page": s5.get("page"),
            "total": s5.get("total"),
            "days": s5.get("days"),
            "tabs": {
                tab: {
                    "click": (v.get("click") or {}).get("ok"),
                    "kept": len(v.get("items") or []),
                    "seen": v.get("total_seen"),
                    "dropped": v.get("dropped_older_than_days"),
                    "stop": v.get("stop"),
                    "error": v.get("error"),
                }
                for tab, v in (s5.get("collected") or {}).items()
            },
            "error": s5.get("error"),
        }
        summary["finished_at"] = f"{datetime.now():%Y-%m-%d %H:%M:%S}"
        if args.stage == "feed" or s5.get("result") != "ok":
            summary["result"] = s5.get("result")
            print(json.dumps(summary, ensure_ascii=False, indent=2))
            return 0 if s5.get("result") == "ok" else 1

    # 步骤6：按步骤5 的清单逐篇抓正文（默认只在 --stage article 时跑）
    run_article = args.stage == "article" or (
        args.stage == "all" and args.with_article and bool(args.keyword) and not args.dry_run
    )
    if run_article:
        s6 = step6_collect_articles(
            limit=args.article_limit,
            max_scroll=args.article_max_scroll,
            dry_run=args.dry_run,
            feed_json=args.feed_json,
            step_notches=args.article_scroll_notches,
            article_mode=args.article_mode,
            find_scroll=args.article_find_scroll,
        )
        dump("step6_articles", s6)
        summary["step6"] = {
            "result": s6.get("result"),
            "account": s6.get("account"),
            "total": s6.get("total"),
            "ok_count": s6.get("ok_count"),
            "articles": [
                {
                    "title": (a.get("page_title") or a.get("title") or "")[:40],
                    "ok": a.get("ok"),
                    "via": a.get("via"),
                    "chars": a.get("chars"),
                    "stop": a.get("stop"),
                    "back": a.get("back_page"),
                    "error": a.get("error"),
                }
                for a in (s6.get("articles") or [])
            ],
            "error": s6.get("error"),
        }
        summary["finished_at"] = f"{datetime.now():%Y-%m-%d %H:%M:%S}"
        if args.stage == "article" or not s6.get("ok_count"):
            summary["result"] = s6.get("result")
            print(json.dumps(summary, ensure_ascii=False, indent=2))
            return 0 if s6.get("result") in ("ok", "partial") else 1

    summary["finished_at"] = f"{datetime.now():%Y-%m-%d %H:%M:%S}"
    summary["result"] = "ok"
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
