"""
Close_ExWin - 關閉異常視窗
使用 WinEvent Hook 監聽視窗建立事件，即時處理 Excel 異常彈窗
不輪詢、不切換前景、視窗一出現立即處理
"""

import sys
import os
import json
import time
import threading
import logging
import ctypes
import ctypes.wintypes
import tkinter as tk
from tkinter import ttk, messagebox
import pystray
from PIL import Image, ImageDraw

# ── 路徑 ────────────────────────────────────────────────────
def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR   = get_base_dir()
CONFIG_FILE = os.path.join(BASE_DIR, "Close_ExWin_config.json")
LOG_DIR        = os.path.join(BASE_DIR, "logs")
SCREENSHOT_DIR = os.path.join(LOG_DIR, "screenshots")
os.makedirs(LOG_DIR,        exist_ok=True)
os.makedirs(SCREENSHOT_DIR, exist_ok=True)
LOG_FILE    = os.path.join(LOG_DIR, f"Close_ExWin_{time.strftime('%Y%m%d_%H%M%S')}.log")

# ── 預設設定 ────────────────────────────────────────────────
DEFAULT_CONFIG = {
    "enable_log": True,
    "enable_screenshot": True,
    "action_delay": 15,
    "rules": [
        {"title": "Microsoft Excel",        "match": "exact",    "keyword": "正在等候", "action": "enter",         "buttons": "確定",       "enabled": True,  "screenshot": False},
        {"title": "Microsoft Excel",        "match": "exact",    "keyword": "",         "action": "enter",         "buttons": "不儲存,取消", "enabled": True,  "screenshot": False},
        {"title": "檔案使用中",              "match": "exact",    "keyword": "",         "action": "tab_tab_enter", "buttons": "",           "enabled": True,  "screenshot": False},
        {"title": "Microsoft Visual Basic", "match": "exact",    "keyword": "",         "action": "click_end",     "buttons": "",           "enabled": True,  "screenshot": True},
        {"title": "Excel",                  "match": "exact",    "keyword": "",         "action": "close",         "buttons": "",           "enabled": True,  "screenshot": False},
        {"title": "活頁簿1 - Excel",         "match": "exact",    "keyword": "",         "action": "close",         "buttons": "",           "enabled": True,  "screenshot": False},
    ]
}

ACTION_LABELS = {
    "enter":         "按 Enter",
    "tab_tab_enter": "按 Tab+Tab+Enter",
    "close":         "強制關閉",
    "click_end":     "按結束(E)",
}

MATCH_LABELS = {
    "exact":    "完全符合",
    "contains": "包含",
}

MATCH_KEYS  = {v: k for k, v in MATCH_LABELS.items()}
ACTION_KEYS = {v: k for k, v in ACTION_LABELS.items()}

# ── 設定讀寫 ────────────────────────────────────────────────
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            return cfg
        except Exception:
            pass
    return json.loads(json.dumps(DEFAULT_CONFIG))

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ── 日誌 ────────────────────────────────────────────────────
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    datefmt="%m/%d %H:%M:%S",
    encoding="utf-8"
)
logging.getLogger("comtypes").setLevel(logging.WARNING)

_enable_log      = True
_enable_log_lock = threading.Lock()

def log(msg):
    with _enable_log_lock:
        if _enable_log:
            logging.info(msg)

def set_log_enabled(val: bool):
    global _enable_log
    with _enable_log_lock:
        _enable_log = val

_enable_screenshot      = True
_enable_screenshot_lock = threading.Lock()

def set_screenshot_enabled(val: bool):
    global _enable_screenshot
    with _enable_screenshot_lock:
        _enable_screenshot = val

def is_screenshot_enabled():
    with _enable_screenshot_lock:
        return _enable_screenshot

_action_delay      = 3
_action_delay_lock = threading.Lock()

def set_action_delay(val: int):
    global _action_delay
    with _action_delay_lock:
        _action_delay = max(0, int(val))

def get_action_delay():
    with _action_delay_lock:
        return _action_delay

# ── Win32 常數與型別 ────────────────────────────────────────
user32   = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

WM_CLOSE   = 0x0010
WM_KEYDOWN = 0x0100
WM_KEYUP   = 0x0101
WM_COMMAND = 0x0111
BM_CLICK   = 0x00F5
VK_RETURN  = 0x0D
VK_TAB     = 0x09
IDOK       = 1
IDNO       = 7

EVENT_SYSTEM_FOREGROUND   = 0x0003
EVENT_OBJECT_SHOW         = 0x8002
WINEVENT_OUTOFCONTEXT     = 0x0000

WinEventProc = ctypes.WINFUNCTYPE(
    None,
    ctypes.wintypes.HANDLE,   # hWinEventHook
    ctypes.wintypes.DWORD,    # event
    ctypes.wintypes.HWND,     # hwnd
    ctypes.wintypes.LONG,     # idObject
    ctypes.wintypes.LONG,     # idChild
    ctypes.wintypes.DWORD,    # dwEventThread
    ctypes.wintypes.DWORD,    # dwmsEventTime
)

# ── 設定（全域，附鎖）─────────────────────────────────────
_config      = load_config()
_config_lock = threading.Lock()

# ── 視窗比對 ────────────────────────────────────────────────
def get_title(hwnd):
    buf = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(hwnd, buf, 512)
    return buf.value

def match_title(window_title, rule_title, match_type):
    if match_type == "exact":
        return window_title == rule_title
    elif match_type == "contains":
        return rule_title in window_title
    return False

def find_rule(title: str, content: str = "") -> dict | None:
    with _config_lock:
        rules = _config.get("rules", [])
    # 先比對有 keyword 的規則（更精確），再比對無 keyword 的規則
    for has_kw in (True, False):
        for r in rules:
            if not r.get("enabled", True):
                continue
            kw = r.get("keyword", "")
            if bool(kw) != has_kw:
                continue
            if kw and kw not in content:
                continue
            if match_title(title, r["title"], r.get("match", "exact")):
                return r
    return None

# ── 視窗動作（不需切換前景）─────────────────────────────────
def press_key(hwnd, vk):
    scan    = user32.MapVirtualKeyW(vk, 0)
    lp_down = (scan << 16) | 1
    lp_up   = (scan << 16) | 0xC0000001
    user32.PostMessageW(hwnd, WM_KEYDOWN, vk, lp_down)
    time.sleep(0.05)
    user32.PostMessageW(hwnd, WM_KEYUP,   vk, lp_up)
    time.sleep(0.05)

GWL_ID       = -12
DM_GETDEFID  = 0x0400
DC_HASDEFID  = 0x5344
GW_OWNER     = 4

# UI Automation 按鈕優先序：先試「不儲存」，讀不到再試「取消」
_UIA_PRIORITY = ["不儲存", "取消"]

def _uia_preload() -> None:
    """啟動時預載 UIAutomationClient 型別庫，避免首次使用時延遲。"""
    try:
        import tempfile
        import comtypes.client
        import comtypes
        if getattr(sys, "frozen", False):
            gd = os.path.join(tempfile.gettempdir(), "comtypes_gen")
            os.makedirs(gd, exist_ok=True)
            comtypes.client.gen_dir = gd
        comtypes.client.GetModule((
            comtypes.GUID("{944DE083-8FB8-45CF-BCB7-C477ACB2F897}"), 1, 0))
    except Exception as e:
        log(f"UIA 預載失敗：{e}")

def _uia_click_button(hwnd: int, priority: list[str]) -> tuple[bool, list[str]]:
    """用 IUIAutomation 列出 NUIDialog 按鈕並依優先序點擊。回傳 (點擊的按鈕名或 None, 按鈕名列表)。"""
    names: list[str] = []
    try:
        import comtypes.client
        import comtypes

        comtypes.CoInitialize()  # 每個 worker thread 各自初始化
        from comtypes.gen import UIAutomationClient as UIA

        uia  = comtypes.client.CreateObject(UIA.CUIAutomation, interface=UIA.IUIAutomation)
        root = uia.ElementFromHandle(hwnd)

        # 對話框訊息文字
        txt_cond = uia.CreatePropertyCondition(
            UIA.UIA_ControlTypePropertyId, UIA.UIA_TextControlTypeId)
        texts = root.FindAll(UIA.TreeScope_Descendants, txt_cond)
        text_parts = [texts.GetElement(i).CurrentName for i in range(texts.Length)
                      if texts.GetElement(i).CurrentName.strip()]
        if text_parts:
            log(f"  → NUIDialog 訊息：{'｜'.join(text_parts)}")
        # 依規則 keyword 比對內文，取得對應的 buttons 優先序
        full_text = " ".join(text_parts)
        nui_title = get_title(hwnd)
        matched = find_rule(nui_title, full_text)
        buttons_str = matched.get("buttons", "") if matched else ""
        if buttons_str:
            priority = [b.strip() for b in buttons_str.split(",") if b.strip()]
            log(f"  → 規則比對「{matched.get('keyword','')}」，按鈕順序：{priority}")

        # 按鈕
        btn_cond = uia.CreatePropertyCondition(
            UIA.UIA_ControlTypePropertyId, UIA.UIA_ButtonControlTypeId)
        btns = root.FindAll(UIA.TreeScope_Descendants, btn_cond)

        for i in range(btns.Length):
            names.append(btns.GetElement(i).CurrentName)

        for target in priority:
            for i in range(btns.Length):
                btn = btns.GetElement(i)
                if btn.CurrentName == target:
                    pat = btn.GetCurrentPattern(UIA.UIA_InvokePatternId)
                    pat.QueryInterface(UIA.IUIAutomationInvokePattern).Invoke()
                    return target, names
    except Exception as e:
        log(f"  → UIA 錯誤：{e}")
    return None, names

def click_ok_button(hwnd) -> bool:
    """對「確定」按鈕送出 BM_CLICK（SendMessage，跨執行緒同步）。"""
    # NUIDialog 是 Office 現代 UI 框架，無標準 Win32 子視窗按鈕，BM_CLICK 無效
    cls_buf = ctypes.create_unicode_buffer(64)
    user32.GetClassNameW(hwnd, cls_buf, 64)
    if cls_buf.value == "NUIDialog":
        clicked_name, btn_names = _uia_click_button(hwnd, _UIA_PRIORITY)
        log(f"  → NUIDialog UIA 按鈕：{btn_names}")
        if clicked_name:
            time.sleep(0.5)
            if not user32.IsWindowVisible(hwnd):
                log(f"  → UIA 點擊「{clicked_name}」後視窗已關閉（hwnd={hwnd:#010x}）")
                return True
            log(f"  → UIA 點擊「{clicked_name}」後視窗仍存在（hwnd={hwnd:#010x}）")
        # UIA 失敗時 fallback
        for cmd_id, label in ((IDNO, "IDNO"), (IDOK, "IDOK")):
            log(f"  → NUIDialog：嘗試 WM_COMMAND({label})  hwnd={hwnd:#010x}")
            user32.SendMessageW(hwnd, WM_COMMAND, cmd_id, 0)
            time.sleep(0.5)
            if not user32.IsWindowVisible(hwnd):
                log(f"  → WM_COMMAND({label}) 後視窗已關閉（hwnd={hwnd:#010x}）")
                return True
            log(f"  → WM_COMMAND({label}) 後視窗仍存在（hwnd={hwnd:#010x}）")
        log(f"  → NUIDialog：嘗試 WM_CLOSE  hwnd={hwnd:#010x}")
        user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
        time.sleep(0.5)
        still = user32.IsWindowVisible(hwnd)
        log(f"  → WM_CLOSE 後視窗{'仍存在' if still else '已關閉'}（hwnd={hwnd:#010x}）")
        return not still  # 關閉成功回 True；失敗回 False 讓 caller fallback Enter

    log("  → enter 動作：優先嘗試 BM_CLICK 點擊確定鈕，失敗才 fallback 按 Enter")
    # 方法 1：DM_GETDEFID 取得預設按鈕 ID（最準確）
    dm = user32.SendMessageW(hwnd, DM_GETDEFID, 0, 0)
    if (dm >> 16) == DC_HASDEFID:
        def_id = dm & 0xFFFF
        def_hwnd = user32.GetDlgItem(hwnd, def_id)
        if def_hwnd:
            log(f"  → DM_GETDEFID id={def_id}，SendMessage BM_CLICK")
            user32.SendMessageW(def_hwnd, BM_CLICK, 0, 0)
            return True

    # 方法 2：標準 IDOK（ID=1）子視窗
    ok_hwnd = user32.GetDlgItem(hwnd, IDOK)
    if ok_hwnd:
        log("  → 找到 IDOK 子視窗，SendMessage BM_CLICK")
        user32.SendMessageW(ok_hwnd, BM_CLICK, 0, 0)
        return True

    # 方法 3：列舉子視窗，找文字為「確定」或「OK」的按鈕
    found: list[int] = [0]
    children: list[str] = []
    def _cb(child: int, _: int) -> bool:
        buf = ctypes.create_unicode_buffer(64)
        user32.GetWindowTextW(child, buf, 64)
        cls = ctypes.create_unicode_buffer(64)
        user32.GetClassNameW(child, cls, 64)
        cid = user32.GetWindowLongW(child, GWL_ID)
        children.append(f"{buf.value!r}(cls={cls.value},id={cid})")
        if buf.value in ("確定", "OK"):
            found[0] = child
            return False
        return True
    user32.EnumChildWindows(hwnd, _ChildEnumProc(_cb), 0)
    if found[0]:
        log("  → 列舉找到確定鈕，SendMessage BM_CLICK")
        user32.SendMessageW(found[0], BM_CLICK, 0, 0)
        return True

    # 找不到時記錄所有子視窗、視窗類別與父視窗，幫助診斷
    parent = user32.GetParent(hwnd)
    log(f"  → 未找到確定鈕，子視窗清單：{children}  cls={cls_buf.value}  parent={parent:#010x}")
    return False

def click_end_button(hwnd) -> bool:
    """找到「結束」按鈕並送出 BM_CLICK；找不到則 do nothing。"""
    found: list[int] = [0]
    def _cb(child: int, _: int) -> bool:
        buf = ctypes.create_unicode_buffer(64)
        user32.GetWindowTextW(child, buf, 64)
        if "結束" in buf.value:
            found[0] = child
            return False
        return True
    user32.EnumChildWindows(hwnd, _ChildEnumProc(_cb), 0)
    if found[0]:
        log("  → 找到結束鈕，SendMessage BM_CLICK")
        user32.SendMessageW(found[0], BM_CLICK, 0, 0)
        return True
    log("  → 未找到結束鈕（子視窗列舉為空），不動作")
    return False

def _handle_dependents(parent_hwnd: int) -> bool:
    """關閉父視窗前，先處理其子視窗及 owned 視窗中符合規則的項目。
    回傳 True 表示所有子視窗已處理（或無子視窗），False 表示有子視窗未能關閉。"""
    found: list[tuple[int, str, dict]] = []

    def _check(hwnd: int) -> None:
        if user32.IsWindowVisible(hwnd):
            t = get_title(hwnd)
            if t:
                r = find_rule(t)
                if r:
                    found.append((hwnd, t, r))

    def _child_cb(child: int, _: int) -> bool:
        _check(child)
        return True

    def _top_cb(hwnd: int, _: int) -> bool:
        if hwnd != parent_hwnd and user32.GetWindow(hwnd, GW_OWNER) == parent_hwnd:
            _check(hwnd)
        return True

    user32.EnumChildWindows(parent_hwnd, _ChildEnumProc(_child_cb), 0)
    user32.EnumWindows(_EnumWindowsProc(_top_cb), 0)

    all_closed = True
    for ch, t, r in found:
        ch_closed = True
        try:
            act = r["action"]
            log(f"先處理子/owned 視窗：{t!r}  hwnd={ch:#010x}  動作：{act}")
            if act == "close":
                user32.PostMessageW(ch, WM_CLOSE, 0, 0)
            elif act == "enter":
                ch_closed = click_ok_button(ch)
                if not ch_closed:
                    press_key(ch, VK_RETURN)
                    time.sleep(0.5)
                    ch_closed = not user32.IsWindowVisible(ch)
                    log(f"  → Enter 後視窗{'仍存在' if not ch_closed else '已關閉'}（hwnd={ch:#010x}）")
            elif act == "click_end":
                click_end_button(ch)
            elif act == "tab_tab_enter":
                press_key(ch, VK_TAB)
                press_key(ch, VK_TAB)
                press_key(ch, VK_RETURN)
        except Exception as e:
            log(f"子/owned 視窗動作失敗：{e}")
            ch_closed = False
        if not ch_closed:
            all_closed = False
        with _scan_cooldown_lock:
            _scan_cooldown[ch] = time.time()
        time.sleep(0.2)
    return all_closed


class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize",          ctypes.c_uint32),
        ("biWidth",         ctypes.c_int32),
        ("biHeight",        ctypes.c_int32),
        ("biPlanes",        ctypes.c_uint16),
        ("biBitCount",      ctypes.c_uint16),
        ("biCompression",   ctypes.c_uint32),
        ("biSizeImage",     ctypes.c_uint32),
        ("biXPelsPerMeter", ctypes.c_int32),
        ("biYPelsPerMeter", ctypes.c_int32),
        ("biClrUsed",       ctypes.c_uint32),
        ("biClrImportant",  ctypes.c_uint32),
    ]

def capture_window(hwnd: int, title: str) -> None:
    if not is_screenshot_enabled():
        return
    try:
        rect = ctypes.wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return
        w = rect.right - rect.left
        h = rect.bottom - rect.top
        if w <= 0 or h <= 0:
            return

        gdi32   = ctypes.windll.gdi32
        hwnd_dc = user32.GetWindowDC(hwnd)
        mem_dc  = gdi32.CreateCompatibleDC(hwnd_dc)
        bitmap  = gdi32.CreateCompatibleBitmap(hwnd_dc, w, h)
        gdi32.SelectObject(mem_dc, bitmap)

        PW_RENDERFULLCONTENT = 0x2
        user32.PrintWindow(hwnd, mem_dc, PW_RENDERFULLCONTENT)

        bmi = _BITMAPINFOHEADER()
        bmi.biSize        = ctypes.sizeof(_BITMAPINFOHEADER)
        bmi.biWidth       = w
        bmi.biHeight      = -h   # 負值 = top-down
        bmi.biPlanes      = 1
        bmi.biBitCount    = 32
        bmi.biCompression = 0    # BI_RGB

        buf = (ctypes.c_char * (w * h * 4))()
        gdi32.GetDIBits(mem_dc, bitmap, 0, h, buf, ctypes.byref(bmi), 0)

        img = Image.frombytes("RGBA", (w, h), bytes(buf), "raw", "BGRA").convert("RGB")

        gdi32.DeleteObject(bitmap)
        gdi32.DeleteDC(mem_dc)
        user32.ReleaseDC(hwnd, hwnd_dc)

        safe = "".join(c if c.isalnum() or c in " _-" else "_" for c in title)[:40]
        path = os.path.join(SCREENSHOT_DIR, f"{time.strftime('%Y%m%d_%H%M%S')}_{safe}.png")
        img.save(path)
        log(f"截圖：{path}")
    except Exception as e:
        log(f"截圖失敗：{e}")

def do_action(hwnd, title, action, screenshot: bool = False):
    parent_hwnd = user32.GetParent(hwnd)
    parent_title = get_title(parent_hwnd) if parent_hwnd else ""
    cls_buf = ctypes.create_unicode_buffer(64)
    user32.GetClassNameW(hwnd, cls_buf, 64)
    parent_info = f"  父視窗：{parent_title!r}(hwnd={parent_hwnd:#010x})" if parent_hwnd else ""
    log(f"偵測到：{title!r}{parent_info}  cls={cls_buf.value}  hwnd={hwnd:#010x}")
    delay = get_action_delay()
    if delay > 0:
        time.sleep(delay)
        cur_title = get_title(hwnd)
        if not user32.IsWindowVisible(hwnd):
            log(f"不動作（視窗已消失）：{title!r}  hwnd={hwnd:#010x}")
            return
        if cur_title != title:
            log(f"不動作（標題已變）：{title!r} → {cur_title!r}  hwnd={hwnd:#010x}")
            return
    if screenshot:
        capture_window(hwnd, title)
    log(f"處理：{title!r}  動作：{action}  hwnd={hwnd:#010x}")
    try:
        if action == "close":
            if _handle_dependents(hwnd):
                user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
            else:
                log(f"  → 子視窗未全部關閉，暫緩關閉父視窗  hwnd={hwnd:#010x}")
        elif action == "enter":
            # 優先用 BM_CLICK 直點按鈕（對 OLE 等待對話框更可靠）
            closed = click_ok_button(hwnd)
            if not closed:
                press_key(hwnd, VK_RETURN)
                time.sleep(0.5)
                closed = not user32.IsWindowVisible(hwnd)
                log(f"  → Enter 後視窗{'仍存在' if not closed else '已關閉'}（hwnd={hwnd:#010x}）")
        elif action == "click_end":
            click_end_button(hwnd)
        elif action == "tab_tab_enter":
            press_key(hwnd, VK_TAB)
            press_key(hwnd, VK_TAB)
            press_key(hwnd, VK_RETURN)
    except Exception as e:
        log(f"動作失敗：{e}")
    # 動作送出後重置 cooldown，避免同一視窗立即被再次偵測
    with _scan_cooldown_lock:
        _scan_cooldown[hwnd] = time.time()

# ── WinEvent Hook 回呼 ──────────────────────────────────────
def _on_win_event(hHook, event, hwnd, idObject, idChild, dwThread, dwTime):
    if not hwnd:
        return
    if is_paused():
        return
    try:
        title = get_title(hwnd)
        if not title:
            return
        rule = find_rule(title)
        if rule and not _parent_has_close_rule(hwnd) and _scan_check(hwnd):
            threading.Thread(
                target=do_action,
                args=(hwnd, title, rule["action"], rule.get("screenshot", False)),
                daemon=True
            ).start()
    except Exception:
        pass

_hook_proc  = WinEventProc(_on_win_event)
_hook_handles = []

def install_hooks():
    """安裝兩個事件鉤子：前景切換 + 視窗顯示"""
    global _hook_handles
    for event_id in (EVENT_SYSTEM_FOREGROUND, EVENT_OBJECT_SHOW):
        h = user32.SetWinEventHook(
            event_id, event_id,
            None, _hook_proc,
            0, 0,
            WINEVENT_OUTOFCONTEXT
        )
        if h:
            _hook_handles.append(h)
    log(f"Hook 已安裝（{len(_hook_handles)} 個）")
    return len(_hook_handles)

_EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
_ChildEnumProc   = ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)

_scan_cooldown      = {}   # hwnd -> last_handled_time
_scan_cooldown_lock = threading.Lock()

def _scan_check(hwnd):
    """回傳 True 表示可以處理（並記錄時間）；False 表示冷卻中，跳過"""
    now = time.time()
    cooldown = get_action_delay() + 10  # 必須大於 action_delay，否則 delay 期間 cooldown 會提早失效
    with _scan_cooldown_lock:
        if now - _scan_cooldown.get(hwnd, 0) < cooldown:
            return False
        _scan_cooldown[hwnd] = now
        # 清除已消失視窗的記錄
        dead = [h for h, t in _scan_cooldown.items()
                if now - t > cooldown * 2 and not user32.IsWindow(h)]
        for h in dead:
            del _scan_cooldown[h]
        return True

def _parent_has_close_rule(hwnd: int) -> bool:
    """父/owner 視窗有 close 規則時回傳 True，由 _handle_dependents 統一處理。"""
    p = user32.GetParent(hwnd)
    if not p:
        return False
    r = find_rule(get_title(p))
    return r is not None and r.get("action") == "close"

def _do_enum_scan():
    def _cb(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            title = get_title(hwnd)
            if title:
                rule = find_rule(title)
                if rule and not _parent_has_close_rule(hwnd) and _scan_check(hwnd):
                    threading.Thread(
                        target=do_action,
                        args=(hwnd, title, rule["action"], rule.get("screenshot", False)),
                        daemon=True
                    ).start()
        return True
    user32.EnumWindows(_EnumWindowsProc(_cb), 0)

def scan_existing_windows():
    """掃描並處理啟動前已存在的視窗"""
    _do_enum_scan()
    log("已掃描現有視窗")

def _periodic_scan():
    """每 2 秒掃描一次，補抓 Hook 漏掉的視窗（如 Excel OLE 等待對話框）"""
    while True:
        time.sleep(2)
        if not is_paused():
            try:
                _do_enum_scan()
            except Exception:
                pass

def uninstall_hooks():
    for h in _hook_handles:
        user32.UnhookWinEvent(h)
    _hook_handles.clear()
    log("Hook 已移除")

def run_message_loop():
    """Windows 訊息迴圈，讓 Hook 能收到事件"""
    msg = ctypes.wintypes.MSG()
    while True:
        ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
        if ret == 0 or ret == -1:
            break
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))

# ── 暫停狀態 ────────────────────────────────────────────────
_paused = False
_paused_lock = threading.Lock()

def is_paused():
    with _paused_lock:
        return _paused

# ── Tray 圖示 ───────────────────────────────────────────────
def make_icon_image(paused=False):
    img  = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    if paused:
        # 暫停：灰色圓 + 兩條直線（暫停符號）
        draw.ellipse([4, 4, 60, 60], fill=(150, 150, 150))
        draw.rectangle([18, 18, 27, 46], fill="white")
        draw.rectangle([37, 18, 46, 46], fill="white")
    else:
        # 執行中：藍色圓 + X
        draw.ellipse([4, 4, 60, 60], fill=(34, 139, 230))
        draw.line([20, 20, 44, 44], fill="white", width=5)
        draw.line([44, 20, 20, 44], fill="white", width=5)
    return img

_settings_open = False

def open_settings(icon, item):
    global _settings_open
    if _settings_open:
        return
    _settings_open = True
    threading.Thread(target=_settings_window, daemon=True).start()

def toggle_pause(icon, item):
    global _paused
    with _paused_lock:
        _paused = not _paused
        paused = _paused
    if paused:
        log("已暫停")
        icon.icon  = make_icon_image(paused=True)
        icon.title = "Close_ExWin（暫停中）"
    else:
        log("已恢復")
        icon.icon  = make_icon_image(paused=False)
        icon.title = "Close_ExWin"
    icon.update_menu()

def pause_label(item):
    with _paused_lock:
        return "▶ 恢復" if _paused else "⏸ 暫停"

def quit_app(icon, item):
    uninstall_hooks()
    icon.stop()
    os._exit(0)

def view_log(icon, item):
    if os.path.exists(LOG_FILE):
        os.startfile(LOG_FILE)

def view_screenshots(icon, item):
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    os.startfile(SCREENSHOT_DIR)

def run_tray():
    img  = make_icon_image(paused=False)
    menu = pystray.Menu(
        pystray.MenuItem(pause_label, toggle_pause),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("設定", open_settings),
        pystray.MenuItem("檢視 Log", view_log),
        pystray.MenuItem("查看截圖", view_screenshots),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("結束", quit_app),
    )
    icon = pystray.Icon("Close_ExWin", img, "Close_ExWin", menu)
    icon.run()

# ── 設定視窗 ────────────────────────────────────────────────
def _settings_window():
    global _settings_open, _config

    root = tk.Tk()
    root.title("Close_ExWin 設定")
    root.geometry("660x560")
    root.resizable(True, True)

    def on_close():
        global _settings_open
        _settings_open = False
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    with _config_lock:
        cfg = json.loads(json.dumps(_config))

    # ── 一般選項 ──
    opt_frame = ttk.LabelFrame(root, text="一般選項", padding=8)
    opt_frame.pack(fill="x", padx=8, pady=(6,2))
    log_var = tk.BooleanVar(value=cfg.get("enable_log", True))
    ttk.Checkbutton(opt_frame, text="記錄 Log（Close_ExWin.log）", variable=log_var).pack(anchor="w")
    screenshot_var = tk.BooleanVar(value=cfg.get("enable_screenshot", True))
    ttk.Checkbutton(opt_frame, text="截圖開關（關閉時全不截圖；開啟時依各規則設定）", variable=screenshot_var).pack(anchor="w")

    delay_row = ttk.Frame(opt_frame)
    delay_row.pack(anchor="w", pady=(4,0))
    ttk.Label(delay_row, text="動作延遲（秒）：").pack(side="left")
    delay_var = tk.IntVar(value=cfg.get("action_delay", 3))
    ttk.Spinbox(delay_row, from_=0, to=30, textvariable=delay_var, width=5).pack(side="left")
    ttk.Label(delay_row, text="（視窗出現後等待幾秒再動作，0 = 立即）").pack(side="left", padx=(6,0))

    # ── 規則清單 ──
    list_frame = ttk.LabelFrame(root, text="視窗規則", padding=8)
    list_frame.pack(fill="both", expand=True, padx=8, pady=6)

    cols = ("enabled", "title", "keyword", "match", "action", "buttons", "screenshot")
    tree = ttk.Treeview(list_frame, columns=cols, show="headings", selectmode="browse")
    tree.heading("enabled",    text="啟用")
    tree.heading("title",      text="視窗標題")
    tree.heading("keyword",    text="關鍵字")
    tree.heading("match",      text="比對")
    tree.heading("action",     text="動作")
    tree.heading("buttons",    text="按鈕")
    tree.heading("screenshot", text="截圖")
    tree.column("enabled",    width=48,  anchor="center")
    tree.column("title",      width=180)
    tree.column("keyword",    width=120)
    tree.column("match",      width=72,  anchor="center")
    tree.column("action",     width=120, anchor="center")
    tree.column("buttons",    width=100)
    tree.column("screenshot", width=48,  anchor="center")

    sb = ttk.Scrollbar(list_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    def refresh_tree():
        tree.delete(*tree.get_children())
        for i, r in enumerate(cfg["rules"]):
            tree.insert("", "end", iid=str(i), values=(
                "✔" if r.get("enabled", True) else "✘",
                r["title"],
                r.get("keyword", ""),
                MATCH_LABELS.get(r.get("match","exact"), r.get("match","exact")),
                ACTION_LABELS.get(r["action"], r["action"]),
                r.get("buttons", ""),
                "✔" if r.get("screenshot", False) else "✘",
            ))

    refresh_tree()

    # ── 編輯區 ──
    ef = ttk.LabelFrame(root, text="新增 / 編輯", padding=8)
    ef.pack(fill="x", padx=8, pady=2)

    ttk.Label(ef, text="視窗標題：").grid(row=0, column=0, sticky="w")
    title_var = tk.StringVar()
    ttk.Entry(ef, textvariable=title_var, width=30).grid(row=0, column=1, sticky="w", padx=4)

    ttk.Label(ef, text="比對：").grid(row=0, column=2, sticky="w", padx=(10,0))
    match_var = tk.StringVar(value=MATCH_LABELS["exact"])
    ttk.Combobox(ef, textvariable=match_var, width=10,
                 values=list(MATCH_LABELS.values()), state="readonly").grid(row=0, column=3, padx=4)

    ttk.Label(ef, text="關鍵字：").grid(row=1, column=0, sticky="w", pady=4)
    keyword_var = tk.StringVar()
    ttk.Entry(ef, textvariable=keyword_var, width=30).grid(row=1, column=1, sticky="w", padx=4)

    ttk.Label(ef, text="按鈕：").grid(row=1, column=2, sticky="w", padx=(10,0))
    buttons_rule_var = tk.StringVar()
    ttk.Entry(ef, textvariable=buttons_rule_var, width=20).grid(row=1, column=3, sticky="w", padx=4)

    ttk.Label(ef, text="動作：").grid(row=2, column=0, sticky="w", pady=4)
    action_var = tk.StringVar(value=ACTION_LABELS["enter"])
    ttk.Combobox(ef, textvariable=action_var, width=18,
                 values=list(ACTION_LABELS.values()), state="readonly").grid(row=2, column=1, sticky="w", padx=4)

    enabled_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(ef, text="啟用", variable=enabled_var).grid(row=2, column=2, sticky="w", padx=(10,0))
    screenshot_rule_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(ef, text="截圖", variable=screenshot_rule_var).grid(row=2, column=3, sticky="w")

    def on_select(event):
        sel = tree.selection()
        if not sel:
            return
        r = cfg["rules"][int(sel[0])]
        title_var.set(r["title"])
        keyword_var.set(r.get("keyword", ""))
        buttons_rule_var.set(r.get("buttons", ""))
        match_var.set(MATCH_LABELS.get(r.get("match", "exact"), r.get("match", "exact")))
        action_var.set(ACTION_LABELS.get(r["action"], r["action"]))
        enabled_var.set(r.get("enabled", True))
        screenshot_rule_var.set(r.get("screenshot", False))

    tree.bind("<<TreeviewSelect>>", on_select)

    bf = ttk.Frame(ef)
    bf.grid(row=3, column=0, columnspan=4, sticky="w", pady=4)

    def add_rule():
        t = title_var.get().strip()
        if not t:
            messagebox.showwarning("提示", "請輸入視窗標題", parent=root)
            return
        cfg["rules"].append({"title":   t,
                              "match":   MATCH_KEYS.get(match_var.get(), match_var.get()),
                              "keyword": keyword_var.get().strip(),
                              "action":  ACTION_KEYS.get(action_var.get(), action_var.get()),
                              "buttons": buttons_rule_var.get().strip(),
                              "enabled": enabled_var.get(),
                              "screenshot": screenshot_rule_var.get()})
        refresh_tree()

    def update_rule():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("提示", "請先選取規則", parent=root)
            return
        t = title_var.get().strip()
        if not t:
            messagebox.showwarning("提示", "標題不可空白", parent=root)
            return
        cfg["rules"][int(sel[0])] = {"title":   t,
                                       "match":   MATCH_KEYS.get(match_var.get(), match_var.get()),
                                       "keyword": keyword_var.get().strip(),
                                       "action":  ACTION_KEYS.get(action_var.get(), action_var.get()),
                                       "buttons": buttons_rule_var.get().strip(),
                                       "enabled": enabled_var.get(),
                                       "screenshot": screenshot_rule_var.get()}
        refresh_tree()

    def delete_rule():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("提示", "請先選取規則", parent=root)
            return
        idx = int(sel[0])
        if messagebox.askyesno("確認", f"確定刪除「{cfg['rules'][idx]['title']}」？", parent=root):
            cfg["rules"].pop(idx)
            refresh_tree()

    ttk.Button(bf, text="新增",     command=add_rule).pack(side="left", padx=2)
    ttk.Button(bf, text="更新選取", command=update_rule).pack(side="left", padx=2)
    ttk.Button(bf, text="刪除選取", command=delete_rule).pack(side="left", padx=2)

    def save_and_close():
        cfg["enable_log"] = log_var.get()
        set_log_enabled(cfg["enable_log"])
        cfg["enable_screenshot"] = screenshot_var.get()
        set_screenshot_enabled(cfg["enable_screenshot"])
        cfg["action_delay"] = delay_var.get()
        set_action_delay(cfg["action_delay"])
        with _config_lock:
            _config.clear()
            _config.update(cfg)
        save_config(cfg)
        messagebox.showinfo("已儲存", "設定已套用。", parent=root)
        on_close()

    ttk.Button(root, text="儲存並關閉", command=save_and_close).pack(pady=6)
    root.mainloop()

# ── 主程式 ──────────────────────────────────────────────────
if __name__ == "__main__":
    # 防止重複執行
    mutex = kernel32.CreateMutexW(None, False, "Close_ExWin_Mutex_v1")
    if kernel32.GetLastError() == 183:
        user32.MessageBoxW(0, "Close_ExWin 已在執行中", "Close_ExWin", 0)
        sys.exit(0)

    set_log_enabled(_config.get("enable_log", True))
    set_screenshot_enabled(_config.get("enable_screenshot", True))
    set_action_delay(_config.get("action_delay", 3))
    log("Close_ExWin 啟動（WinEvent Hook 模式）")
    _uia_preload()
    log(f"設定檔：{'存在' if os.path.exists(CONFIG_FILE) else '不存在（使用預設）'}  {CONFIG_FILE}")
    log(f"action_delay={_config.get('action_delay')}s  log={_config.get('enable_log')}  screenshot={_config.get('enable_screenshot')}")
    for i, r in enumerate(_config.get("rules", []), 1):
        status = "✓" if r.get("enabled", True) else "✗"
        log(f"  規則{i:02d} [{status}] {r['match']:8s} {r['title']!r:40s} → {r['action']}")
    if install_hooks() == 0:
        user32.MessageBoxW(0,
            "WinEvent Hook 安裝失敗，程式無法監聽視窗事件。\n請確認以系統管理員身份執行。",
            "Close_ExWin 警告", 0x30)
        log("Hook 安裝失敗，程式終止")

    scan_existing_windows()

    threading.Thread(target=_periodic_scan, daemon=True).start()

    # 訊息迴圈跑在子執行緒
    msg_thread = threading.Thread(target=run_message_loop, daemon=True)
    msg_thread.start()

    # Tray 在主執行緒
    run_tray()
