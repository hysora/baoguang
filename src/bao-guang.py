"""
bao-guang.py  —  多平台播放量/浏览量抓取
支持：X (Twitter) · Instagram · Facebook · TikTok
四个平台并行运行，各自使用独立浏览器 profile 保存登录状态。
"""

from playwright.sync_api import sync_playwright, BrowserContext, Page
from concurrent.futures import ThreadPoolExecutor, as_completed
import sys
import time
import os
import re
import traceback
import threading
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ──────────────────────────────────────────────
# 配置：在此填写所有账号 URL（自动识别平台）
# ──────────────────────────────────────────────
URLS = [
]

# Google Sheets 配置（留空则跳过，直接用上面的 URLS）
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1FYmDs2lKsrKQ_cye7xpUguuYAY2TsMK57GsFtYXTBhE/edit?gid=0#gid=0"

# URL → 元数据 的映射，由 fetch_google_sheet_urls() 填充
# 每条值形如 {"seq": "1", "name": "王永洋", "tab": "Sheet1"}
URL_META: dict[str, dict] = {}

# 平台识别规则（域名关键词 → 平台名）
_PLATFORM_RULES = [
    ("x.com",        "X"),
    ("twitter.com",  "X"),
    ("instagram.com","Instagram"),
    ("facebook.com", "Facebook"),
    ("tiktok.com",   "TikTok"),
]

def detect_platform(url: str) -> str | None:
    for keyword, platform in _PLATFORM_RULES:
        if keyword in url:
            return platform
    return None

def group_urls_by_platform(urls: list[str]) -> dict[str, list[str]]:
    """按平台分组，忽略无法识别的 URL"""
    grouped: dict[str, list[str]] = {}
    for url in urls:
        platform = detect_platform(url)
        if platform is None:
            print(f"[警告] 无法识别平台，已跳过：{url}")
            continue
        grouped.setdefault(platform, []).append(url)
    return grouped

# ──────────────────────────────────────────────
# 浏览器 Profile 目录（各平台独立，互不干扰）
# ──────────────────────────────────────────────
# 打包为 exe 后 __file__ 指向临时解压目录，需改用 exe 所在目录
if getattr(sys, "frozen", False):
    _BASE = os.path.dirname(sys.executable)
else:
    _BASE = os.path.dirname(os.path.abspath(__file__))


def _find_chrome() -> str:
    """查找系统 Chrome 可执行文件路径，找不到则抛出异常"""
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe")
        path, _ = winreg.QueryValueEx(key, "")
        if os.path.exists(path):
            return path
    except Exception:
        pass
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe"),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    raise RuntimeError(
        "未找到系统 Chrome，请先安装 Google Chrome：https://www.google.com/chrome/"
    )


CHROME_PATH = _find_chrome()


PROFILE = {
    "x":        os.path.join(_BASE, ".profile_x"),
    "instagram": os.path.join(_BASE, ".profile_instagram"),
    "facebook": os.path.join(_BASE, ".profile_facebook"),
    # TikTok 不使用持久化 profile（反爬较严，用全新浏览器）
}

# ── 持久化浏览器上下文（跨多次 main() 调用保持不关闭）──────────────
# platform -> (playwright, BrowserContext, thread_id)
_CTX_STORE: dict[str, tuple] = {}
# TikTok 专用：[playwright, browser, BrowserContext, thread_id]
_TT_STORE:  list = []


def _is_context_alive(ctx) -> bool:
    try:
        _ = ctx.pages
        return True
    except Exception:
        return False


def _close_platform_context(platform: str):
    """关闭指定平台的浏览器上下文，必须在创建该 context 的同一线程内调用"""
    entry = _CTX_STORE.pop(platform, None)
    if entry is None:
        return
    p, ctx, _ = entry
    try:
        ctx.close()
    except Exception:
        pass
    try:
        p.stop()
    except Exception:
        pass


def _ensure_context(platform: str) -> tuple:
    """返回平台对应的 (playwright, context)；context 失效时才重建，存活则直接复用。"""
    current_tid = threading.get_ident()
    if platform in _CTX_STORE:
        p, ctx, _ = _CTX_STORE[platform]
        if _is_context_alive(ctx):
            # context 仍然存活（无论哪个线程创建的），直接复用并更新线程记录
            _CTX_STORE[platform] = (p, ctx, current_tid)
            return p, ctx
        # context 已失效，关闭后重建
        try:
            ctx.close()
        except Exception:
            pass
        try:
            p.stop()
        except Exception:
            pass
        del _CTX_STORE[platform]

    p = sync_playwright().start()
    ctx = make_context(p, platform)
    _CTX_STORE[platform] = (p, ctx, current_tid)
    return p, ctx


def _ensure_tiktok() -> tuple:
    """返回 TikTok 的 (playwright, browser, context)；context 失效时才重建，存活则直接复用。"""
    current_tid = threading.get_ident()
    if _TT_STORE:
        p, browser, ctx, tid = _TT_STORE
        if _is_context_alive(ctx):
            _TT_STORE[3] = current_tid
            return p, browser, ctx
        try:
            ctx.close()
        except Exception:
            pass
        try:
            p.stop()
        except Exception:
            pass
        _TT_STORE.clear()

    p = sync_playwright().start()
    browser = p.chromium.launch(
        executable_path=CHROME_PATH,
        headless=False,
        args=["--disable-blink-features=AutomationControlled"],
    )
    ctx = browser.new_context(
        viewport={"width": 1280, "height": 900},
        locale="zh-CN",
    )
    ctx.add_init_script(_STEALTH_SCRIPT)
    _TT_STORE.extend([p, browser, ctx, current_tid])
    return p, browser, ctx


def _get_or_new_page(ctx):
    """复用 context 中第一个已有页面，没有则新建"""
    pages = ctx.pages
    if pages:
        pages[0].bring_to_front()
        return pages[0]
    return ctx.new_page()


def open_login_page(platform: str):
    """
    直接用 Chrome（非 Playwright 控制）打开登录页，避免被平台识别为自动化。
    登录成功后 cookie 保存在与 Playwright 共用的持久化 profile 目录中，
    后续定时运行时 Playwright 会自动读取这些 cookie。
    """
    import subprocess
    import webbrowser

    _PLATFORM_HOME = {
        "x":         "https://x.com",
        "instagram": "https://www.instagram.com",
        "facebook":  "https://www.facebook.com",
        "tiktok":    "https://www.tiktok.com",
    }
    url = _PLATFORM_HOME.get(platform)
    if not url:
        return

    profile_dir = PROFILE.get(platform)
    if platform != "tiktok" and profile_dir and CHROME_PATH:
        # 先释放 Playwright 对该 profile 目录的文件锁，否则普通 Chrome 会回退到临时 profile，
        # 导致登录 cookie 写不进共用目录
        _close_platform_context(platform)
        # 用普通 Chrome（不带 Playwright 自动化标志）打开，cookie 写入同一 profile
        try:
            subprocess.Popen([CHROME_PATH, f"--user-data-dir={profile_dir}", url])
            print(f"[{platform.upper()}] Chrome 已打开 {url}")
            print(f"[{platform.upper()}] 请完成登录，关闭浏览器后 cookie 会自动保存，下次自动化运行即可使用。")
            return
        except Exception as e:
            print(f"[{platform.upper()}] 直接启动 Chrome 失败，回退到系统浏览器: {e}")

    # TikTok 或回退：用系统默认浏览器打开
    webbrowser.open(url)
    print(f"[{platform.upper()}] 已用系统浏览器打开 {url}，请手动完成登录。")
    if platform == "tiktok":
        print("[TIKTOK] 注意：TikTok 每次抓取使用全新浏览器，登录状态无法自动传递给自动化。")

# Excel 写入锁（多平台并行时保护文件）
_excel_lock = threading.Lock()

# ──────────────────────────────────────────────
# 通用工具
# ──────────────────────────────────────────────

def parse_view_count(text: str) -> int | None:
    """解析播放量文本，支持：3,926 / 3.9K / 1.2M / 1.2万/萬 / 3亿/億"""
    # \xa0 是不换行空格（&nbsp;），Facebook 繁体页面常用
    text = text.strip().replace(",", "").replace(" ", "").replace("\xa0", "")
    if re.fullmatch(r'\d+', text):
        return int(text)
    m = re.fullmatch(r'([\d.]+)([KkMmBb])', text)
    if m:
        multiplier = {"K": 1_000, "M": 1_000_000, "B": 1_000_000_000}
        return int(float(m.group(1)) * multiplier[m.group(2).upper()])
    m = re.fullmatch(r'([\d.]+)[万萬]', text)  # 简体万 + 繁体萬
    if m:
        return int(float(m.group(1)) * 10_000)
    m = re.fullmatch(r'([\d.]+)[亿億]', text)  # 简体亿 + 繁体億
    if m:
        return int(float(m.group(1)) * 100_000_000)
    return None


def safe_goto(page: Page, url: str):
    """跳转页面，忽略内部重定向引发的导航中断异常；网络错误时每秒重试最多60次"""
    deadline = time.time() + 60
    attempt = 0
    while True:
        try:
            page.goto(url, wait_until="load", timeout=60000)
            return
        except Exception as e:
            err_str = str(e)
            if "interrupted by another navigation" in err_str:
                page.wait_for_load_state("domcontentloaded", timeout=30000)
                return
            if "net::ERR_" in err_str and time.time() < deadline:
                attempt += 1
                print(f"[网络错误] 第 {attempt} 次重试（{url[:60]}）...")
                time.sleep(1)
                continue
            raise


_STEALTH_SCRIPT = """
    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
    Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
    Object.defineProperty(navigator, 'languages', {get: () => ['zh-CN','zh','en']});
    window.chrome = {runtime: {}};
    const _query = window.navigator.permissions.query;
    window.navigator.permissions.query = (p) =>
        p.name === 'notifications'
            ? Promise.resolve({state: Notification.permission})
            : _query(p);
"""


def make_context(p, platform: str) -> BrowserContext:
    """创建持久化浏览器上下文"""
    print(f"[{platform.upper()}] CHROME_PATH={CHROME_PATH}")
    print(f"[{platform.upper()}] exists={os.path.exists(CHROME_PATH) if CHROME_PATH else False}")
    ctx = p.chromium.launch_persistent_context(
        user_data_dir=PROFILE[platform],
        executable_path=CHROME_PATH,
        headless=False,
        args=["--disable-blink-features=AutomationControlled"],
        viewport={"width": 1280, "height": 900},  # type: ignore[arg-type]
        locale="zh-CN",
    )
    print(f"[{platform.upper()}] Chrome 启动成功")
    ctx.add_init_script(_STEALTH_SCRIPT)
    return ctx


def scroll_until_stable(page: Page, delay: float = 2.0):
    """滚动到底部触发懒加载，等待网络静默后再继续（用于非 Facebook 平台）"""
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    time.sleep(delay)
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except Exception:
        pass  # 社交平台后台轮询可能导致 networkidle 永远不触发，忽略超时


def scroll_and_wait_new_elements(page: Page, selector: str, current_dom_count: int,
                                  poll: float = 0.4, timeout: float = 3.0) -> int:
    """
    滚动后等待：新容器数量 > current_dom_count 即立即返回。
    只要求"出现新容器"，不再要求"所有容器都已填充文字"——后者在 FB 通用类名下
    几乎永远不成立，会导致每次滚动都空耗满 timeout。骨架未渲染的元素外层
    parse_view_count 会自动跳过，下一轮滚动会再次抓到。
    超时后直接返回当前数量，交给外层 no_new 计数判断是否到底。
    """
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(poll)
        total = page.evaluate(
            f"() => document.querySelectorAll('{selector}').length")
        # 有新容器即返回
        if total > current_dom_count:
            return total
    return len(page.query_selector_all(selector))


def _wait_for_content(page: Page, selector: str, timeout: int = 15000):
    """等待指定选择器出现，超时则静默继续"""
    try:
        page.wait_for_selector(selector, timeout=timeout)
    except Exception:
        pass


# ──────────────────────────────────────────────
# X (Twitter)
# ──────────────────────────────────────────────

def _x_is_logged_in(page: Page) -> bool:
    for selector in [
        '[data-testid="SideNav_AccountSwitcher_Button"]',
        '[data-testid="AppTabBar_Home_Link"]',
        'a[href="/home"]',
    ]:
        try:
            page.wait_for_selector(selector, timeout=3000)
            return True
        except Exception:
            pass
    return False


def _x_ensure_login(page: Page):
    safe_goto(page, "https://x.com/")
    if _x_is_logged_in(page):
        print("[X] 检测到已登录，直接开始抓取。")
    else:
        print("[X] 尚未登录，请在浏览器中完成登录后按 Enter 继续...")
        input()
        page.wait_for_load_state("load", timeout=60000)


def _x_extract_views(aria_label: str) -> int | None:
    m = re.search(r'([\d,]+)\s+views', aria_label)
    if m:
        return int(m.group(1).replace(",", ""))
    return None


def scrape_x(urls: list[str], check_login: bool = True) -> dict[str, list[dict]]:
    results: dict[str, list[dict]] = {}
    _, context = _ensure_context("x")
    page = _get_or_new_page(context)
    if check_login:
        _x_ensure_login(page)
    for url in urls:
        print(f"\n[X] 正在访问：{url}")
        try:
            safe_goto(page, url)
        except Exception as e:
            print(f"[X] 网络错误，跳过当前账号，已保留 {len(results)} 个结果：{e}")
            results[url] = []
            continue
        time.sleep(3)
        _wait_for_content(page, '[data-testid="tweet"]')

        seen: set[str] = set()
        items: list[dict] = []
        last_count = 0
        no_new = 0

        while True:
            for tweet in page.query_selector_all('[data-testid="tweet"]'):
                link_el = tweet.query_selector('a[href*="/status/"]')
                tid = link_el.get_attribute("href") if link_el else None
                if not tid or tid in seen:
                    continue
                seen.add(tid)

                # 过滤转发
                ctx = tweet.query_selector('[data-testid="socialContext"]')
                if ctx:
                    t = ctx.inner_text().strip().lower()
                    if "reposted" in t or "转发" in t:
                        continue

                # 获取浏览量
                views_el = tweet.query_selector('[aria-label*="views. View post analytics"]')
                if not views_el:
                    views_el = tweet.query_selector('[role="group"][aria-label*="views"]')
                if not views_el:
                    continue

                vc = _x_extract_views(views_el.get_attribute("aria-label") or "")
                if vc is None:
                    continue

                items.append({"index": len(items) + 1, "views": vc,
                              "url": "https://x.com" + tid})

            cur = len(items)
            print(f"[X]   已收集 {cur} 条原创推文...")
            if cur == last_count:
                no_new += 1
                if no_new >= 3:
                    break
            else:
                no_new = 0
            last_count = cur
            scroll_until_stable(page)
            _wait_for_content(page, '[data-testid="tweet"]')

        results[url] = items
    _close_platform_context("x")
    return results


# ──────────────────────────────────────────────
# Instagram
# ──────────────────────────────────────────────

IG_VIEW_SELECTOR = (
    "span.xdj266r.x14z9mp.xat24cr.x1lziwak.xexx8yu"
    ".xyri2b.x18d9i69.x1c1uobl.x1hl2dhg.x16tdsg8.x1vvkbs"
)
IG_REEL_LINK_SELECTOR = 'a[href*="/reel/"]'


def _ig_is_logged_in(page: Page) -> bool:
    try:
        page.wait_for_selector('nav[role="navigation"]', timeout=5000)
        return page.query_selector('a[href="/accounts/login/"]') is None
    except Exception:
        return False


def _ig_has_login_dialog(page: Page) -> bool:
    """检测页面是否弹出了登录对话框（浏览时触发的登录墙）"""
    try:
        dialog = page.query_selector('[role="dialog"]')
        if not dialog:
            return False
        # 对话框内含登录链接
        if dialog.query_selector('a[href*="/accounts/login/"]'):
            return True
        # 对话框内文字含登录/注册关键词
        text = dialog.inner_text()
        return any(kw in text for kw in ["Log in", "登录", "Sign up", "注册"])
    except Exception:
        return False


def _ig_wait_for_login(page: Page):
    """检测到登录弹窗时，自动轮询直到真正登录完成才继续"""
    print("[IG] ⚠ 检测到登录弹窗，请在浏览器中完成登录，登录完成后将自动继续...")
    # 最多等待 10 分钟，每 2 秒检查一次
    for _ in range(300):
        time.sleep(2)
        try:
            cur_url = page.url
        except Exception:
            continue
        # 还在登录页或登录表单可见 → 继续等待
        if "login" in cur_url or "accounts/login" in cur_url:
            continue
        if page.query_selector('input[name="username"], input[name="password"]'):
            continue
        if _ig_has_login_dialog(page):
            continue
        print("[IG] ✅ 已确认登录完成，继续抓取。")
        return
    print("[IG] 等待登录超时，继续尝试抓取...")


def _ig_ensure_login(page: Page):
    safe_goto(page, "https://www.instagram.com/")
    if _ig_is_logged_in(page):
        print("[IG] 检测到已登录，直接开始抓取。")
    else:
        print("[IG] 尚未登录，请在浏览器中完成登录后按 Enter 继续...")
        input()


def scrape_instagram(urls: list[str], check_login: bool = True) -> dict[str, list[dict]]:
    results: dict[str, list[dict]] = {}
    _, context = _ensure_context("instagram")
    page = _get_or_new_page(context)
    if check_login:
        _ig_ensure_login(page)

    _js_extract = """
    (sel) => {
        const links = document.querySelectorAll('a[href*="/reel/"]');
        const out = [];
        for (const a of links) {
            const spans = a.querySelectorAll(sel);
            if (spans.length === 0) continue;
            const last = spans[spans.length - 1];
            out.push({href: a.href, text: (last.innerText || '').trim()});
        }
        return out;
    }
    """

    for url in urls:
        print(f"\n[IG] 正在访问：{url}")
        try:
            safe_goto(page, url)
        except Exception as e:
            print(f"[IG] 网络错误，跳过当前账号，已保留 {len(results)} 个结果：{e}")
            results[url] = []
            continue
        time.sleep(3)

        # 页面加载后立即检查是否弹出登录框
        if _ig_has_login_dialog(page):
            _ig_wait_for_login(page)

        _wait_for_content(page, IG_REEL_LINK_SELECTOR)

        items: list[dict] = []
        seen_hrefs: set[str] = set()
        last_new_time = time.time()
        IDLE_TIMEOUT = 10.0  # 连续 10 秒无新内容则认为到底

        while True:
            # 每次滚动前检查登录弹窗
            if _ig_has_login_dialog(page):
                _ig_wait_for_login(page)
                last_new_time = time.time()  # 登录等待期间不计入空闲

            try:
                reel_stats = page.evaluate(_js_extract, IG_VIEW_SELECTOR) or []
            except Exception:
                reel_stats = []

            for entry in reel_stats:
                href = entry.get("href", "")
                text = entry.get("text", "")
                if not href or href in seen_hrefs:
                    continue
                vc = parse_view_count(text)
                if vc is None:
                    continue
                seen_hrefs.add(href)
                items.append({"index": len(items) + 1, "views": vc, "url": url, "href": href})
                last_new_time = time.time()

            print(f"[IG]   已加载 {len(items)} 个播放量...")

            if time.time() - last_new_time > IDLE_TIMEOUT:
                print(f"[IG]   {IDLE_TIMEOUT:.0f}s 内无新内容，判断已到底。")
                break

            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(1)

        results[url] = items
    _close_platform_context("instagram")
    return results


# ──────────────────────────────────────────────
# Facebook
# ──────────────────────────────────────────────

FB_VIEW_SELECTOR = "span.x1lliihq.x6ikm8r.x10wlt62.x1n2onr6.xlyipyv.xuxw1ft"


def _fb_is_logged_in(page: Page) -> bool:
    for label in ["你的主页", "Your profile", "Profile"]:
        try:
            page.wait_for_selector(f'[aria-label="{label}"]', timeout=2000)
            return True
        except Exception:
            pass
    return page.query_selector('a[href*="login"]') is None


def _fb_page_has_login_wall(page: Page) -> bool:
    """检测当前页面是否出现登录墙（含"登录"/"Log in"文字的遮罩）"""
    result = page.evaluate("""
        () => {
            // 检查页面是否含有登录遮罩：查找含"登录"或"Log in"的按钮/链接
            const texts = ["登录", "Log in", "Login"];
            for (const el of document.querySelectorAll('a[href*="login"], button')) {
                const t = (el.innerText || "").trim();
                if (texts.includes(t)) return true;
            }
            // 检查用户提供的具体 span 类名
            for (const span of document.querySelectorAll(
                'span.x1lliihq.x6ikm8r.x10wlt62.x1n2onr6.xlyipyv.xuxw1ft')) {
                const t = (span.innerText || "").trim();
                if (texts.includes(t)) return true;
            }
            return false;
        }
    """)
    return bool(result)


def _fb_ensure_login(page: Page):
    safe_goto(page, "https://www.facebook.com/")
    time.sleep(3)
    if _fb_is_logged_in(page):
        print("[FB] 检测到已登录，直接开始抓取。")
    else:
        print("[FB] 尚未登录，请在浏览器中完成登录后按 Enter 继续...")
        input()


def scrape_facebook(urls: list[str], check_login: bool = True) -> dict[str, list[dict]]:
    results: dict[str, list[dict]] = {}
    _, context = _ensure_context("facebook")
    page = _get_or_new_page(context)
    if check_login:
        _fb_ensure_login(page)

    for url in urls:
        print(f"\n[FB] 正在访问：{url}")
        try:
            safe_goto(page, url)
        except Exception as e:
            print(f"[Facebook] 网络错误，跳过当前账号，已保留 {len(results)} 个结果：{e}")
            results[url] = []
            continue
        time.sleep(4)
        _wait_for_content(page, FB_VIEW_SELECTOR, timeout=5000)

        if _fb_page_has_login_wall(page):
            print(f"[FB] 检测到登录墙，跳过抓取，曝光量记为 -1：{url}")
            results[url] = [{"index": 1, "views": -1, "url": url}]
            continue

        seen: set[str] = set()
        items: list[dict] = []
        last_count = 0
        no_new = 0
        dom_count = len(page.query_selector_all(FB_VIEW_SELECTOR))

        while True:
            data = page.evaluate(f"""
                () => {{
                    const out = [];
                    for (const span of document.querySelectorAll("{FB_VIEW_SELECTOR}")) {{
                        const text = span.innerText.trim();
                        if (!text) continue;
                        let el = span, reelHref = null;
                        for (let i = 0; i < 15; i++) {{
                            el = el.parentElement;
                            if (!el) break;
                            const a = el.querySelector('a[href*="/reel/"]');
                            if (a) {{ reelHref = a.getAttribute("href"); break; }}
                            if (el.tagName === "A" && (el.href || "").includes("/reel/")) {{
                                reelHref = el.getAttribute("href"); break;
                            }}
                        }}
                        out.push({{ text, reelHref }});
                    }}
                    return out;
                }}
            """)

            for item in data:
                vc = parse_view_count(item.get("text", ""))
                if vc is None:
                    continue
                href = item.get("reelHref") or ""
                # 有 reelHref 用链接去重；无 reelHref 用"文本+播放量"去重
                key = href if href else f"nolink:{item.get('text')}"
                if key in seen:
                    continue
                seen.add(key)
                reel_url = ("https://www.facebook.com" + href) if href else "(未知链接)"
                items.append({"index": len(items) + 1, "views": vc, "url": reel_url})

            cur = len(items)
            print(f"[FB]   已收集 {cur} 条 Reel...")
            if cur == last_count:
                no_new += 1
                if no_new >= 3:
                    break
            else:
                no_new = 0
            last_count = cur
            # 滚动后轮询 DOM 元素数，新元素一出现立即继续，无需等满 networkidle 超时
            dom_count = scroll_and_wait_new_elements(page, FB_VIEW_SELECTOR, dom_count)

        results[url] = items
    _close_platform_context("facebook")
    return results


# ──────────────────────────────────────────────
# TikTok
# ──────────────────────────────────────────────

def scrape_tiktok(urls: list[str], check_login: bool = True) -> dict[str, list[dict]]:  # noqa: check_login unused
    results: dict[str, list[dict]] = {}
    _, _, context = _ensure_tiktok()

    for url in urls:
        print(f"\n[TT] 正在访问：{url}")
        page = context.new_page()
        try:
            safe_goto(page, url)
            page.wait_for_selector('[data-e2e="video-views"]', timeout=30000)
            time.sleep(2)

            items: list[dict] = []
            last_count = 0
            no_new = 0

            while True:
                elements = page.query_selector_all('[data-e2e="video-views"]')
                for el in elements[last_count:]:
                    text = el.inner_text().strip()
                    vc = parse_view_count(text)
                    if vc is not None:
                        items.append({"index": len(items) + 1, "views": vc, "url": url})

                cur = len(elements)
                print(f"[TT]   已加载 {cur} 个视频...")
                if cur == last_count:
                    no_new += 1
                    if no_new >= 3:
                        break
                else:
                    no_new = 0
                last_count = cur
                scroll_until_stable(page)
                _wait_for_content(page, '[data-e2e="video-views"]')

            results[url] = items
        except Exception as e:
            print(f"[TT] 网络错误，跳过当前账号，已保留 {len(results)} 个结果：{e}")
            results[url] = []
        finally:
            page.close()

    if _TT_STORE:
        p, browser, ctx, _ = _TT_STORE
        _TT_STORE.clear()
        try:
            ctx.close()
        except Exception:
            pass
        try:
            browser.close()
        except Exception:
            pass
        try:
            p.stop()
        except Exception:
            pass
    return results


# ──────────────────────────────────────────────
# 工具：从 URL 提取账号名
# ──────────────────────────────────────────────

_SKIP_SEGMENTS = {"reels", "videos", "posts", "profile.php", ""}

def extract_account(url: str) -> str:
    # Facebook：从 id= 参数提取
    if "id=" in url:
        m = re.search(r'id=(\d+)', url)
        return m.group(1) if m else url
    # 取 path 各段，过滤掉已知非用户名的部分，取第一个有效段
    from urllib.parse import urlparse
    path_parts = [p.split("?")[0] for p in urlparse(url).path.strip("/").split("/")]
    for part in path_parts:
        if part and part not in _SKIP_SEGMENTS:
            return part
    return url


# ──────────────────────────────────────────────
# 汇总输出（终端）
# ──────────────────────────────────────────────

def print_results(platform: str, all_results: dict[str, list[dict]]):
    print(f"\n{'='*10} {platform} {'='*10}")
    for url, items in all_results.items():
        account = extract_account(url)
        total = sum(i["views"] for i in items)
        print(f"\n  账号: {account}  共 {len(items)} 条内容")
        for item in items:
            print(f"    {item['index']:>4}: {item['views']:>12,} 次  {item.get('url', '')}")
        print(f"  合计: {total:,}")


# ──────────────────────────────────────────────
# Google Sheets 读取
# ──────────────────────────────────────────────

_PLATFORM_KW = ["x.com", "twitter.com", "instagram.com",
                "facebook.com", "tiktok.com", "youtube.com"]


def _read_xlsx(xlsx_path: str, label: str = "Sheet") -> dict[str, dict]:
    """
    读取 Excel 文件，按行解析，识别"序号"/"sheet名称"/"链接"列。
    直接解析 xlsx zip 内的 XML，完全绕开 openpyxl 样式解析。
    返回 {url: {"seq": str, "name": str, "tab": str}} 映射。
    """
    import zipfile
    import xml.etree.ElementTree as ET

    _SS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    _REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    _PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
    _url_re = re.compile(r'https?://\S+')

    def _col_idx(ref: str) -> int:
        """将列字母 'A'/'AB' 转为 0-based 整数索引"""
        m = re.match(r'([A-Za-z]+)', ref)
        if not m:
            return 0
        idx = 0
        for ch in m.group(1).upper():
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        return idx - 1

    def _cell_text(c_el, shared: list[str]) -> str:
        v_el = c_el.find(f"{{{_SS}}}v")
        if v_el is None or not v_el.text:
            return ""
        if c_el.get("t") == "s":
            try:
                return shared[int(v_el.text)]
            except (IndexError, ValueError):
                return ""
        return v_el.text.strip()

    # 表头关键词映射
    _SEQ_KW  = {"序号", "no", "no.", "#"}
    _NAME_KW = {"姓名", "sheet名称", "名称", "账号", "name"}
    _URL_KW  = {"链接", "主页地址", "url", "网址", "地址"}
    _TMPL_KW = {"模版", "模板", "template"}

    url_meta: dict[str, dict] = {}
    _tmpl_col_map: dict[str, int | None] = {}  # sheet -> tmpl col idx

    try:
        with zipfile.ZipFile(xlsx_path) as zf:
            names = set(zf.namelist())

            # 共享字符串表
            shared: list[str] = []
            if "xl/sharedStrings.xml" in names:
                root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
                for si in root.findall(f".//{{{_SS}}}si"):
                    parts = [t.text for t in si.findall(f".//{{{_SS}}}t") if t.text]
                    shared.append("".join(parts))

            # sheet tab 名称
            rid_to_name: dict[str, str] = {}
            if "xl/workbook.xml" in names:
                root = ET.fromstring(zf.read("xl/workbook.xml"))
                for sh in root.findall(f".//{{{_SS}}}sheet"):
                    rid = sh.get(f"{{{_REL}}}id", "")
                    rid_to_name[rid] = sh.get("name", rid)

            rid_to_path: dict[str, str] = {}
            if "xl/_rels/workbook.xml.rels" in names:
                root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
                for rel in root.findall(f".//{{{_PKG}}}Relationship"):
                    rid_to_path[rel.get("Id", "")] = rel.get("Target", "")

            for rid, sname in rid_to_name.items():
                target = rid_to_path.get(rid, "")
                if not target:
                    continue
                sheet_path = target if target in names else f"xl/{target}"
                if sheet_path not in names:
                    continue

                root = ET.fromstring(zf.read(sheet_path))

                # 解析每行：{row_num: {col_idx: text}}
                rows: dict[int, dict[int, str]] = {}
                for row_el in root.findall(f".//{{{_SS}}}row"):
                    r = int(row_el.get("r", 0))
                    row_data: dict[int, str] = {}
                    for c_el in row_el.findall(f"{{{_SS}}}c"):
                        cidx = _col_idx(c_el.get("r", ""))
                        text = _cell_text(c_el, shared)
                        if text:
                            row_data[cidx] = text
                    if row_data:
                        rows[r] = row_data

                # 在前 10 行中找表头
                seq_col = name_col = url_col = tmpl_col = None
                header_row = 0
                for r in sorted(rows)[:10]:
                    for cidx, val in rows[r].items():
                        v = val.strip().lower()
                        if v in _SEQ_KW:
                            seq_col = cidx
                        elif v in _NAME_KW:
                            name_col = cidx
                        elif v in _URL_KW:
                            url_col = cidx
                        elif v in _TMPL_KW:
                            tmpl_col = cidx
                    if seq_col is not None or url_col is not None:
                        header_row = r
                        break

                count = 0
                for r in sorted(rows):
                    if r <= header_row:
                        continue
                    row_data = rows[r]

                    # 找 URL：优先指定列，否则全行扫描
                    url = ""
                    search_cells = (
                        [row_data[url_col]] if url_col is not None and url_col in row_data
                        else list(row_data.values())
                    )
                    for text in search_cells:
                        for u in _url_re.findall(text):
                            u = u.rstrip(".,;:!?）】》」'\"")
                            if any(k in u for k in _PLATFORM_KW):
                                url = u
                                break
                        if url:
                            break

                    if not url or url in url_meta:
                        continue

                    seq  = row_data.get(seq_col, "") if seq_col is not None else ""
                    name = row_data.get(name_col, "") if name_col is not None else ""
                    tmpl = row_data.get(tmpl_col, "") if tmpl_col is not None else ""
                    url_meta[url] = {"seq": seq, "name": name, "tab": sname, "tmpl": tmpl}
                    count += 1

                print(f"[{label}] Sheet '{sname}' → {count} 条新 URL")

    except Exception as e:
        print(f"[{label}] 读取 xlsx 失败: {e}")
    finally:
        try:
            os.unlink(xlsx_path)
        except Exception:
            pass

    return url_meta


def fetch_google_sheet_urls() -> dict[str, dict]:
    """
    下载 Google Sheets 并解析平台 URL。支持两种格式：
    - 公开发布链接（…/pub?output=xlsx）：直接 HTTP 下载，无需浏览器
    - 私有文档链接（…/edit…）：打开浏览器登录后下载
    """
    import tempfile
    import urllib.request

    if not GOOGLE_SHEET_URL:
        return {}

    tmp_path = os.path.join(tempfile.gettempdir(),
                            f"google_sheet_{int(time.time())}.xlsx")

    # ── 判断是否为公开发布链接 ──────────────────────────────────────
    is_public = (
        "/pub?" in GOOGLE_SHEET_URL
        or "/pub#" in GOOGLE_SHEET_URL
        or GOOGLE_SHEET_URL.rstrip("/").endswith("/pub")
        or "/d/e/" in GOOGLE_SHEET_URL
    )

    if is_public:
        # 确保 output=xlsx 参数存在
        if "output=xlsx" in GOOGLE_SHEET_URL:
            download_url = GOOGLE_SHEET_URL
        elif "?" in GOOGLE_SHEET_URL:
            download_url = GOOGLE_SHEET_URL + "&output=xlsx"
        else:
            download_url = GOOGLE_SHEET_URL + "?output=xlsx"

        print(f"[Google] 检测到公开发布链接，直接下载...")
        try:
            urllib.request.urlretrieve(download_url, tmp_path)
            print(f"[Google] xlsx 已下载: {tmp_path}")
        except Exception as e:
            print(f"[Google] 下载失败: {e}")
            return {}

    else:
        # 私有文档：用浏览器登录后下载
        m = re.search(r'/spreadsheets/d/([^/]+)', GOOGLE_SHEET_URL)
        if not m:
            print("[Google] 无法从 URL 提取 Sheet ID")
            return {}
        sheet_id = m.group(1)
        export_url = (f"https://docs.google.com/spreadsheets/d/"
                      f"{sheet_id}/export?format=xlsx")
        profile_dir = os.path.join(_BASE, ".profile_google")

        with sync_playwright() as p:
            context = p.chromium.launch_persistent_context(
                user_data_dir=profile_dir,
                executable_path=CHROME_PATH,
                headless=False,
                args=["--disable-blink-features=AutomationControlled"],
                viewport={"width": 1280, "height": 900},  # type: ignore[arg-type]
                locale="zh-CN",
            )
            page = context.new_page()

            print("[Google] 打开 Google Sheets，如需登录请在浏览器中完成...")
            safe_goto(page, GOOGLE_SHEET_URL)

            try:
                page.wait_for_selector(
                    "#docs-editor, .waffle-name-box, .grid-container",
                    timeout=600000)
                print("[Google] 文档已加载完成。")
            except Exception:
                print("[Google] 等待超时，尝试继续...")
            time.sleep(2)

            print("[Google] 开始下载 xlsx...")
            try:
                with page.expect_download(timeout=30000) as dl_info:
                    try:
                        page.goto(export_url, timeout=10000)
                    except Exception:
                        pass
                dl_info.value.save_as(tmp_path)
                print(f"[Google] xlsx 已下载: {tmp_path}")
            except Exception as e:
                print(f"[Google] 下载失败: {e}")
                context.close()
                return {}

            context.close()

    url_meta = _read_xlsx(tmp_path, label="Google")
    print(f"[Google] 共读取 {len(url_meta)} 条 URL")
    return url_meta


# ──────────────────────────────────────────────
# 导出 Excel（增量追加，按日期归档）
# ──────────────────────────────────────────────

_EXCEL_PATH: str = ""


def get_today_excel_path() -> str:
    global _EXCEL_PATH
    if not _EXCEL_PATH:
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        _EXCEL_PATH = os.path.join(desktop, f"播放量汇总_{ts}.xlsx")
    return _EXCEL_PATH


def export_excel(all_platform_results: dict[str, dict]):
    """按 Google 文档原始顺序一次性写入 Excel"""
    # 展平为 url → items
    url_to_items: dict[str, list[dict]] = {}
    for results in all_platform_results.values():
        url_to_items.update(results)

    if not url_to_items:
        print("[Excel] 无数据可写入")
        return

    # 按原始 URLS 顺序排列，未在 URLS 中的补到末尾
    seen: set[str] = set()
    ordered: list[str] = []
    for u in URLS:
        if u in url_to_items and u not in seen:
            ordered.append(u)
            seen.add(u)
    for u in url_to_items:
        if u not in seen:
            ordered.append(u)

    path = get_today_excel_path()
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(fill_type="solid", fgColor="4472C4")
    center      = Alignment(horizontal="center", vertical="center")

    # 构建 profile_url → platform 映射
    url_to_platform: dict[str, str] = {}
    for platform_key, results in all_platform_results.items():
        for u in results:
            url_to_platform[u] = platform_key

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "明细"
    detail_headers = ["姓名", "唯一id", "平台", "模板", "播放量"]
    ws.append(detail_headers)
    for col_idx in range(1, len(detail_headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center

    total_views = 0
    for url in ordered:
        items    = url_to_items[url]
        meta     = URL_META.get(url, {})
        name     = meta.get("name", "") or extract_account(url)
        tmpl     = meta.get("tmpl", "")
        platform = url_to_platform.get(url, detect_platform(url) or "")
        for item in items:
            unique_id = item.get("href") or item.get("url", "")
            ws.append([name, unique_id, platform, tmpl, item["views"]])
            total_views += item["views"]

    for col in ws.columns:
        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

    wb.save(path)
    print(f"\n[Excel] 共 {len(ordered)} 个账号，总曝光量 {total_views:,}")
    print(f"[Excel] 数据文件：{path}")


# ──────────────────────────────────────────────
# 主入口：并行抓取四个平台
# ──────────────────────────────────────────────

def main(check_login: bool = False, test_mode: bool = False):
    global URLS, URL_META, _EXCEL_PATH
    _EXCEL_PATH = ""  # 每次运行重置，强制生成新文件名

    # 从 Google Sheets 读取 URL 数据源
    if GOOGLE_SHEET_URL:
        sheet_map = fetch_google_sheet_urls()
        if sheet_map:
            URL_META = sheet_map
            URLS = list(sheet_map.keys())
            print(f"\n[Google] 共读取 {len(URLS)} 条 URL，开始抓取...\n")
        else:
            print("[Google] 未从表格读取到有效 URL，使用配置中的静态 URLS")

    scrape_fn = {
        "X":         scrape_x,
        "Instagram": scrape_instagram,
        "Facebook":  scrape_facebook,
        "TikTok":    scrape_tiktok,
    }

    grouped = group_urls_by_platform(URLS)
    if test_mode:
        grouped = {p: urls[:3] for p, urls in grouped.items()}
        print("[测试模式] 每个平台最多处理 3 条 URL\n")

    tasks = {
        platform: (scrape_fn[platform], urls)
        for platform, urls in grouped.items()
        if platform in scrape_fn
    }

    if not tasks:
        print("[警告] 没有可识别的平台 URL，跳过抓取。")
        export_excel({})
        return

    all_platform_results: dict[str, dict] = {}

    with ThreadPoolExecutor(max_workers=len(tasks)) as executor:
        futures = {
            executor.submit(fn, urls, check_login): platform
            for platform, (fn, urls) in tasks.items()
        }
        for future in as_completed(futures):
            platform = futures[future]
            try:
                all_platform_results[platform] = future.result()
            except Exception as e:
                print(f"\n[{platform}] 抓取失败: {e}")
                all_platform_results[platform] = {}

    print("\n\n" + "=" * 50)
    print("  汇总结果")
    print("=" * 50)
    for platform, results in all_platform_results.items():
        print_results(platform, results)

    export_excel(all_platform_results)


def _close_all_contexts():
    """关闭所有平台的浏览器上下文"""
    for platform, entry in list(_CTX_STORE.items()):
        p, ctx, _ = entry
        try:
            ctx.close()
        except Exception:
            pass
        try:
            p.stop()
        except Exception:
            pass
    _CTX_STORE.clear()

    if _TT_STORE:
        p, browser, ctx, _ = _TT_STORE
        try:
            ctx.close()
        except Exception:
            pass
        try:
            browser.close()
        except Exception:
            pass
        try:
            p.stop()
        except Exception:
            pass
        _TT_STORE.clear()


def _wait_until_next_8am():
    """阻塞到明天 08:00:00"""
    now = datetime.now()
    next_run = now.replace(hour=8, minute=0, second=0, microsecond=0)
    if now >= next_run:
        # 今天 8 点已过，等到明天
        from datetime import timedelta
        next_run += timedelta(days=1)
    wait_sec = (next_run - now).total_seconds()
    print(f"\n下次运行时间：{next_run.strftime('%Y-%m-%d %H:%M:%S')}，等待 {wait_sec/3600:.1f} 小时...")
    time.sleep(wait_sec)


if __name__ == "__main__":
    log_path = os.path.join(_BASE, "bao-guang.log")
    while True:
        try:
            main()
        except KeyboardInterrupt:
            raise
        except Exception:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
            print(f"\n运行出错，详情已写入：{log_path}")
        try:
            _wait_until_next_8am()
        except KeyboardInterrupt:
            break
