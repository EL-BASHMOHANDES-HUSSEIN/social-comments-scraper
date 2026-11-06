# -*- coding: utf-8 -*-
"""
Universal Social Media Comments Scraper (Facebook + Instagram)

Install:
    pip install selenium webdriver-manager pandas openpyxl tqdm

Run:
    python universal_scraper.py
    (optionally pass the URL directly:  python universal_scraper.py "https://...")

First run for each platform: log in manually in the Chrome window, then press Enter
in the terminal. Sessions are stored in ./fb_profile and ./ig_profile.

Notice: automating Facebook/Instagram violates their Terms of Service and may lead to
temporary restrictions on your account. Use at your own risk, slowly, and only on data
you are entitled to collect.
"""

import random
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional
from urllib.parse import urlparse

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from tqdm import tqdm
from webdriver_manager.chrome import ChromeDriverManager

BASE_DIR = Path(__file__).resolve().parent
DESKTOP = Path.home() / "Desktop"
if not DESKTOP.exists():
    DESKTOP = Path.home()

MAX_IDLE_ROUNDS = 10      # rounds with no new comments before stopping
SAVE_EVERY = 300          # backup save every N comments


# ============================================================================
#  JavaScript snippets (run inside the page)
# ============================================================================

# ----------------------------- Facebook ------------------------------------
FB_EXTRACT_JS = r"""
const out = [];
for (const a of document.querySelectorAll('div[role="article"]')) {
  const label = (a.getAttribute('aria-label') || '').trim();
  if (!/^(Comment|Reply|تعليق|رد)\b/i.test(label)) continue;

  let name = '';
  const nameEl = a.querySelector('a[role="link"] span, a[role="link"]');
  if (nameEl) name = (nameEl.innerText || '').trim();
  if (!name) {
    const m = label.match(/^(?:Comment by|Reply by|تعليق من|رد من)\s+(.+?)(?:\s+(?:\d+|منذ|\w+ ago).*)?$/i);
    if (m) name = m[1].trim();
  }

  const parts = [];
  for (const b of a.querySelectorAll('div[dir="auto"], span[dir="auto"]')) {
    if (b.closest('a[role="link"]')) continue;
    const t = (b.innerText || '').trim();
    if (t && t !== name && !parts.includes(t)) parts.push(t);
  }
  const text = parts.join('\n').trim();

  let date = '', link = '';
  const links = a.querySelectorAll('a[href*="comment_id"], a[href*="reply_comment_id"]');
  if (links.length) {
    const l = links[links.length - 1];
    date = (l.innerText || l.getAttribute('aria-label') || '').trim();
    link = l.href;
  }
  if (!date) {
    const ab = a.querySelector('abbr, [aria-label][role="link"][tabindex]');
    if (ab) date = (ab.getAttribute('aria-label') || ab.innerText || '').trim();
  }
  out.push({name, text, date, link, is_reply: /^(Reply|رد)\b/i.test(label)});
}
return out;
"""

FB_LOAD_MORE_JS = r"""
const keywords = arguments[0];
let clicked = 0;
for (const b of document.querySelectorAll('div[role="button"], span[role="button"]')) {
  const t = (b.innerText || '').trim();
  if (!t || t.length > 60) continue;
  if (keywords.some(k => t.includes(k))) {
    try { b.scrollIntoView({block: 'center'}); b.click(); clicked++; } catch (e) {}
    if (clicked >= 3) break;
  }
}
return clicked;
"""

FB_REPLIES_JS = r"""
let clicked = 0;
for (const b of document.querySelectorAll('div[role="button"], span[role="button"]')) {
  const t = (b.innerText || '').trim();
  if (!t || t.length > 40) continue;
  if (/^(View|عرض)\s+(all\s+)?\d*\s*(more\s+)?(repl|رد)/i.test(t) || /^\d+\s+(repl|رد)/i.test(t)) {
    try { b.scrollIntoView({block: 'center'}); b.click(); clicked++; } catch (e) {}
    if (clicked >= 3) break;
  }
}
return clicked;
"""

FB_SCROLL_JS = r"""
const arts = document.querySelectorAll('div[role="article"]');
if (arts.length) arts[arts.length - 1].scrollIntoView({block: 'end'});
window.scrollBy(0, arguments[0]);
"""

FB_MORE_KEYWORDS = [
    "عرض المزيد من التعليقات", "عرض التعليقات السابقة", "عرض المزيد من الردود",
    "عرض المزيد", "View more comments", "View previous comments", "View more replies",
]
FB_MOST_RELEVANT = ["الأكثر صلة", "Most relevant", "الأكثر ملاءمة"]
FB_ALL_COMMENTS = ["جميع التعليقات", "All comments"]

# ----------------------------- Instagram -----------------------------------
IG_EXTRACT_JS = r"""
const SKIP = new Set(['reply','like','see translation','hidden','edited','pinned','verified','author',
  'ردّ','رد','إعجاب','أعجبني','عرض الترجمة','تم التعديل','مثبّت','موثّق','الكاتب']);
const userRe = /^\/([A-Za-z0-9._]+)\/$/;
const commentLinks = el => new Set([...el.querySelectorAll('a[href*="/c/"]')]
  .map(x => x.getAttribute('href')).filter(h => /\/c\/\d+/.test(h)));
const out = [], seen = new Set();

for (const t of document.querySelectorAll('time[datetime]')) {
  const a = t.closest('a');
  if (!a) continue;
  if (!/\/c\/\d+/.test(a.getAttribute('href') || '')) continue;
  const link = a.href;
  if (seen.has(link)) continue;

  // climb to the largest container that holds exactly one comment
  let box = a.parentElement, best = null;
  for (let i = 0; i < 12 && box && box !== document.body; i++) {
    if (commentLinks(box).size > 1) break;
    if ([...box.querySelectorAll('a[href]')].some(x => userRe.test(x.getAttribute('href')))) best = box;
    box = box.parentElement;
  }
  if (!best) continue;
  box = best;

  const userLink = [...box.querySelectorAll('a[href]')].find(x => userRe.test(x.getAttribute('href')));
  const name = (userLink.innerText || '').trim() || userLink.getAttribute('href').replace(/\//g, '');

  const parts = [];
  for (const el of box.querySelectorAll('span[dir="auto"], div[dir="auto"]')) {
    if (el.querySelector('span[dir="auto"], div[dir="auto"]')) continue;
    if (userLink.contains(el) || el.contains(userLink)) continue;
    if (el.closest('time, button, [role="button"]')) continue;
    if (el.closest('a[href*="/c/"]')) continue;
    const s = (el.innerText || '').trim();
    if (s && s !== name && !SKIP.has(s.toLowerCase()) && !parts.includes(s)) parts.push(s);
  }
  let text = parts.join(' ').trim();

  if (!text) {
    const timeText = (t.innerText || '').trim();
    text = box.innerText.split('\n').map(s => s.trim()).filter(Boolean)
      .filter(s => s !== name && s !== timeText && !SKIP.has(s.toLowerCase())
                   && !/^\d+\s*(likes?|إعجاب\w*)$/i.test(s)
                   && !/^(view|عرض)\s.*(repl|رد)/i.test(s)
                   && !/^(hide|إخفاء)\s.*(repl|رد)/i.test(s)
                   && !/^[-—–]+$/.test(s))
      .join(' ');
  }

  const ul = box.closest('ul');
  const isReply = !!(ul && ul.parentElement && ul.parentElement.closest('ul'));
  seen.add(link);
  out.push({name, text, date: t.getAttribute('datetime'), link, is_reply: isReply});
}
return out;
"""

IG_LOAD_MORE_JS = r"""
let clicked = 0;
for (const l of ['Load more comments', 'تحميل المزيد من التعليقات', 'عرض المزيد من التعليقات']) {
  for (const svg of document.querySelectorAll(`svg[aria-label="${l}"]`)) {
    const b = svg.closest('button, div[role="button"]');
    if (b) { try { b.scrollIntoView({block:'center'}); b.click(); clicked++; } catch(e){} }
    if (clicked >= 2) return clicked;
  }
}
return clicked;
"""

IG_REPLIES_JS = r"""
let clicked = 0;
for (const b of document.querySelectorAll('div[role="button"], button, span[role="button"]')) {
  const t = (b.innerText || '').trim();
  if (!t || t.length > 50) continue;
  if (/^(view|عرض)\b.*(repl|رد)/i.test(t) && !/hide|إخفاء/i.test(t)) {
    try { b.scrollIntoView({block:'center'}); b.click(); clicked++; } catch(e){}
    if (clicked >= 3) break;
  }
}
return clicked;
"""

IG_SCROLL_JS = r"""
const amount = arguments[0];
const times = [...document.querySelectorAll('time[datetime]')];
const last = times.length ? times[times.length - 1] : null;
let el = last, scroller = null;
while (el && el !== document.body) {
  el = el.parentElement;
  if (el && el.scrollHeight > el.clientHeight + 20) {
    const ov = getComputedStyle(el).overflowY;
    if (ov === 'auto' || ov === 'scroll') { scroller = el; break; }
  }
}
if (scroller) scroller.scrollBy({top: amount, behavior: 'smooth'});
else window.scrollBy({top: amount, behavior: 'smooth'});
if (last) last.scrollIntoView({block: 'end'});
"""

IG_DISMISS_JS = r"""
for (const b of document.querySelectorAll('button, div[role="button"]')) {
  const t = (b.innerText || '').trim();
  if (['Not now','Not Now','ليس الآن','Close','إغلاق'].includes(t)) { b.click(); return true; }
}
return false;
"""


# ============================================================================
#  Helpers
# ============================================================================

def human_sleep(a: float = 1.5, b: float = 4.0) -> None:
    time.sleep(random.uniform(a, b))


def info(msg: str) -> None:
    tqdm.write(f"[INFO] {msg}")


def warn(msg: str) -> None:
    tqdm.write(f"[WARN] {msg}")


def ask_choice(prompt: str, options: list, default: int = 1) -> int:
    """Simple numbered menu. Returns the 1-based index chosen."""
    print(f"\n{prompt}")
    for i, opt in enumerate(options, 1):
        print(f"  {i}) {opt}")
    while True:
        raw = input(f"Select [default {default}]: ").strip()
        if not raw:
            return default
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return int(raw)
        print(f"Please enter a number between 1 and {len(options)}.")


def ask_int(prompt: str, default: int) -> int:
    while True:
        raw = input(f"{prompt} [default {default}]: ").strip()
        if not raw:
            return default
        if raw.isdigit() and int(raw) > 0:
            return int(raw)
        print("Please enter a positive number.")


_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL",
                   *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}


def sanitize_filename(raw: str, default: str, fmt: str) -> str:
    """
    Turns user input into a safe file stem (no extension, no illegal characters).
    Falls back to `default` when the input is empty or unusable.
    """
    name = raw.strip().strip('"').strip("'")
    # drop a typed extension such as ".xlsx" / ".csv" (the chosen format decides the extension)
    for ext in (".xlsx", ".csv"):
        if name.lower().endswith(ext):
            name = name[: -len(ext)]
    name = _INVALID_FILENAME_CHARS.sub("_", name).strip(" .")
    if not name or name.upper() in _RESERVED_NAMES:
        return default
    return name[:120]


def unique_stem(stem: str, fmt: str) -> str:
    """
    Never overwrite an existing file: if the target (or its backup) already exists on the
    Desktop, append _1, _2, ... to the name.
    """
    def taken(s: str) -> bool:
        return (DESKTOP / f"{s}.{fmt}").exists() or (DESKTOP / f"{s}_backup.{fmt}").exists()

    if not taken(stem):
        return stem
    i = 1
    while taken(f"{stem}_{i}"):
        i += 1
    return f"{stem}_{i}"


def ask_filename(default_stem: str, fmt: str) -> str:
    """Asks for the output file name and returns the final (collision-free) stem."""
    raw = input("\nEnter the desired filename for the Excel/CSV file (or press Enter to use default): ")
    chosen = sanitize_filename(raw, default_stem, fmt)
    typed = raw.strip().strip('"').strip("'")
    for ext in (".xlsx", ".csv"):
        if typed.lower().endswith(ext):
            typed = typed[: -len(ext)]
    if typed and chosen != typed:
        print(f"[INFO] Filename adjusted to a valid name: {chosen}")
    final = unique_stem(chosen, fmt)
    if final != chosen:
        print(f"[INFO] '{chosen}.{fmt}' already exists on your Desktop - "
              f"using '{final}.{fmt}' instead so nothing is overwritten.")
    return final


# ============================================================================
#  Platform definitions + Smart Router
# ============================================================================

@dataclass
class Platform:
    key: str                       # "facebook" | "instagram"
    label: str
    home_url: str
    profile_dir: Path
    output_name: str
    extract_js: str
    scroll_js: str
    scroll_range: tuple
    need_login: Callable           # (driver) -> bool
    prepare: Optional[Callable]    # called once after the post is opened
    load_more: Callable            # (driver, include_replies) -> int clicked


def _fb_need_login(driver) -> bool:
    return bool(driver.find_elements(By.NAME, "pass")) or "login" in driver.current_url


def _ig_need_login(driver) -> bool:
    return bool(driver.find_elements(By.NAME, "username")) or bool(driver.find_elements(By.NAME, "password"))


def _fb_prepare(driver) -> None:
    """Switch comment ordering from 'Most relevant' to 'All comments'."""
    try:
        for b in driver.find_elements(By.CSS_SELECTOR, 'div[role="button"], span[role="button"]'):
            t = (b.text or "").strip()
            if any(k in t for k in FB_MOST_RELEVANT):
                b.click()
                human_sleep(1, 2)
                for it in driver.find_elements(By.CSS_SELECTOR, 'div[role="menuitem"], div[role="menuitemradio"]'):
                    if any(k in (it.text or "") for k in FB_ALL_COMMENTS):
                        it.click()
                        human_sleep(2, 3)
                        info("Comment order switched to: All comments")
                        return
                break
    except WebDriverException:
        pass


def _ig_prepare(driver) -> None:
    driver.execute_script(IG_DISMISS_JS)


def _fb_load_more(driver, include_replies: bool) -> int:
    n = driver.execute_script(FB_LOAD_MORE_JS, FB_MORE_KEYWORDS)
    if include_replies:
        n += driver.execute_script(FB_REPLIES_JS)
    return n


def _ig_load_more(driver, include_replies: bool) -> int:
    n = driver.execute_script(IG_LOAD_MORE_JS)
    if include_replies:
        n += driver.execute_script(IG_REPLIES_JS)
    return n


FACEBOOK = Platform(
    key="facebook", label="Facebook", home_url="https://www.facebook.com/",
    profile_dir=BASE_DIR / "fb_profile", output_name="facebook_comments",
    extract_js=FB_EXTRACT_JS, scroll_js=FB_SCROLL_JS, scroll_range=(400, 1200),
    need_login=_fb_need_login, prepare=_fb_prepare, load_more=_fb_load_more,
)

INSTAGRAM = Platform(
    key="instagram", label="Instagram", home_url="https://www.instagram.com/",
    profile_dir=BASE_DIR / "ig_profile", output_name="instagram_comments",
    extract_js=IG_EXTRACT_JS, scroll_js=IG_SCROLL_JS, scroll_range=(300, 900),
    need_login=_ig_need_login, prepare=_ig_prepare, load_more=_ig_load_more,
)

IG_URL_RE = re.compile(r"/(?:[A-Za-z0-9._]+/)?(p|reel|reels|tv)/([A-Za-z0-9_-]+)")


def route_url(raw: str):
    """
    Smart router. Returns (Platform, clean_url, token_that_must_appear_in_final_url)
    or None if the URL is invalid/unsupported (an English error is printed).
    """
    raw = raw.strip().strip('"').strip("'")
    if not raw:
        print("[ERROR] No URL entered. Please paste a post URL.")
        return None
    if not re.match(r"^https?://", raw, re.I):
        raw = "https://" + raw

    try:
        parsed = urlparse(raw)
    except ValueError:
        print("[ERROR] Could not parse this URL. Please enter a valid link.")
        return None

    host = (parsed.hostname or "").lower()

    if host == "instagram.com" or host.endswith(".instagram.com"):
        m = IG_URL_RE.search(parsed.path + "/")
        if not m:
            print("[ERROR] This Instagram link is not a post/reel. "
                  "Expected something like https://www.instagram.com/p/XXXX/")
            return None
        kind = "reel" if m.group(1) == "reels" else m.group(1)
        return INSTAGRAM, f"https://www.instagram.com/{kind}/{m.group(2)}/", m.group(2)

    if host in ("fb.watch",) or host == "facebook.com" or host.endswith(".facebook.com") \
            or host == "fb.com" or host.endswith(".fb.com"):
        # Facebook links (esp. /share/ links) depend on their path; keep the URL as is.
        return FACEBOOK, raw, ""

    print(f"[ERROR] Unsupported platform: '{host or raw}'. "
          "Only Facebook (facebook.com, fb.watch) and Instagram (instagram.com) links are supported. "
          "Please enter a valid link.")
    return None


# ============================================================================
#  Scraper engine
# ============================================================================

def build_driver(profile_dir: Path):
    opts = Options()
    opts.add_argument(f"--user-data-dir={profile_dir}")
    opts.add_argument("--start-maximized")
    opts.add_argument("--lang=en-US")
    opts.add_argument("--disable-notifications")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)


def ensure_login(driver, platform: Platform) -> None:
    driver.get(platform.home_url)
    human_sleep(4, 6)
    if platform.need_login(driver):
        print(f"\n>> Please log in to {platform.label} manually in the Chrome window.")
        input(">> After you are logged in, press Enter here to continue... ")
    else:
        info("Saved session found - already logged in.")


def open_post(driver, url: str, token: str) -> bool:
    for attempt in range(1, 4):
        try:
            driver.get(url)
            human_sleep(4, 7)
            if not token or token in driver.current_url:
                info(f"Post opened: {driver.current_url[:100]}")
                return True
            warn(f"Unexpected page ({driver.current_url[:70]}). Retrying ({attempt}/3)...")
        except WebDriverException as e:
            warn(f"Connection error ({attempt}/3): {str(e.msg)[:80]}")
            time.sleep(10)
    return False


def save_records(records: dict, stem: str, fmt: str, final: bool) -> Optional[Path]:
    if not records:
        return None
    df = pd.DataFrame(list(records.values()))
    df = df.rename(columns={"name": "Username", "text": "Comment", "date": "Date", "link": "Link"})
    df = df[["Username", "Comment", "Date", "Link"]]
    path = DESKTOP / f"{stem}{'' if final else '_backup'}.{fmt}"
    if fmt == "xlsx":
        df.to_excel(path, index=False, engine="openpyxl")
    else:
        df.to_csv(path, index=False, encoding="utf-8-sig")
    return path


def scrape(platform: Platform, url: str, token: str, fmt: str,
           include_replies: bool, expected: int, stem: str) -> None:
    driver = build_driver(platform.profile_dir)
    records: dict = {}
    try:
        ensure_login(driver, platform)
        if not open_post(driver, url, token):
            print("[ERROR] Could not open the post. Check the link, your connection, or access permissions.")
            return
        if platform.prepare:
            platform.prepare(driver)

        idle, last_count, last_save = 0, 0, 0
        pbar = tqdm(total=expected, unit="comment", desc=f"Scraping {platform.label}",
                    dynamic_ncols=True)

        while idle < MAX_IDLE_ROUNDS:
            try:
                for r in driver.execute_script(platform.extract_js):
                    if not r["name"] and not r["text"]:
                        continue
                    if r["is_reply"] and not include_replies:
                        continue
                    key = r["link"] or (r["name"], r["text"], r["date"])
                    old = records.get(key)
                    # never replace a record that has text with an empty one
                    if old is None or (not old["text"] and r["text"]):
                        records[key] = r

                clicked = platform.load_more(driver, include_replies)

                driver.execute_script(platform.scroll_js, random.randint(*platform.scroll_range))
                human_sleep(1.6, 4.2)
                if random.random() < 0.1:       # occasional longer, human-like pause
                    human_sleep(6, 12)

                n = len(records)
                if n > pbar.total:
                    pbar.total = n + 300
                pbar.n = n
                pbar.set_postfix(idle=f"{idle}/{MAX_IDLE_ROUNDS}", refresh=False)
                pbar.refresh()

                idle = idle + 1 if (n == last_count and clicked == 0) else 0
                last_count = n

                if n - last_save >= SAVE_EVERY:
                    save_records(records, stem, fmt, final=False)
                    last_save = n

            except WebDriverException as e:
                warn(f"Temporary error: {str(e.msg)[:80]} - retrying in 15s")
                time.sleep(15)
                idle += 1
                try:
                    driver.refresh()
                    human_sleep(5, 8)
                except WebDriverException:
                    pass

        pbar.close()
        info("No more new comments were found - finishing.")

    except KeyboardInterrupt:
        print("\n[INFO] Stopped by user. Saving what has been collected so far...")
    finally:
        path = save_records(records, stem, fmt, final=True)
        if path:
            empty = sum(1 for r in records.values() if not r["text"])
            print(f"\n[DONE] Saved {len(records)} comments to:\n       {path}")
            if empty:
                print(f"[NOTE] {empty} comments have empty text (emoji/sticker only, or the page layout changed).")
        else:
            print("\n[DONE] No comments were found. The page layout may have changed, "
                  "or the post has no visible comments.")
        try:
            driver.quit()
        except Exception:
            pass


# ============================================================================
#  Interactive CLI
# ============================================================================

def main() -> None:
    print("=" * 60)
    print("   Universal Comments Scraper  -  Facebook & Instagram")
    print("=" * 60)

    # 1) URL + smart routing
    routed = None
    pending = sys.argv[1] if len(sys.argv) > 1 else ""
    while routed is None:
        raw = pending or input("\nEnter the post URL: ")
        pending = ""
        routed = route_url(raw)
    platform, url, token = routed
    print(f"\n[OK] Detected platform: {platform.label}")
    print(f"[OK] Target: {url}")

    # 2) Options
    fmt = "xlsx" if ask_choice("Choose the output format:", ["Excel (.xlsx)", "CSV (.csv)"]) == 1 else "csv"
    include_replies = ask_choice("Scrape replies to comments as well?",
                                 ["Yes - include replies", "No - top-level comments only"]) == 1
    expected = ask_int("Approximate number of comments (only used to size the progress bar)", 1000)
    stem = ask_filename(platform.output_name, fmt)

    print("\n--- Summary ---")
    print(f"Platform : {platform.label}")
    print(f"Format   : {fmt.upper()}")
    print(f"Replies  : {'yes' if include_replies else 'no'}")
    print(f"Output   : {DESKTOP / (stem + '.' + fmt)}")
    if input("\nPress Enter to start (or type 'q' to quit): ").strip().lower() == "q":
        print("Cancelled.")
        return

    scrape(platform, url, token, fmt, include_replies, expected, stem)


if __name__ == "__main__":
    main()
