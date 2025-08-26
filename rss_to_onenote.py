#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, sys, time, html, re
from datetime import datetime
from urllib.parse import urljoin, quote

import requests
import feedparser
from bs4 import BeautifulSoup

import msal
from msal import PublicClientApplication
from msal_extensions import PersistedTokenCache, FilePersistence

# ========== 配置 ==========
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Notes.ReadWrite.CreatedByApp"]

BASE_SECTION = (os.getenv("ONENOTE_SECTION_NAME") or "").strip()   # 基础分区名，例如 "RSS syn2"
SECTION_MONTHLY = (os.getenv("SECTION_MONTHLY") or "false").lower() == "true"

TOKEN_CACHE_FILE = "token_cache.bin"
PROCESSED_ITEMS_FILE = "processed_items.txt"

MAX_ITEMS_PER_RUN = 50
REQUEST_TIMEOUT = 25
REQUEST_DELAY = 2

# --------- RSS 列表（可自行增删）---------
ORIGINAL_FEEDS = [
    "http://www.jintiankansha.me/rss/GIZDGNRVGJ6GIZRWGRRGMYZUGIYTOOBUMEYDSZTDMJQTEOJXHAZTGOBVGJRDMZJWMFSDSZJRHE3A====",
    "http://www.jintiankansha.me/rss/GEYDQMJQHF6DQOJRMZTDIZTGHAYDMZJSMYYWMOBUGM3GEZTGMQ2DIMBRHA3GGNLEGFSDENDEMEYQ====",
    "http://www.jintiankansha.me/rss/GMZTONZSGR6DMMZXGA4DQMZWGE3TAMTDMEYGGOJTGRSGMNLBGM4GGY3CGAZGENLGGQ3TMN3GMMZQ====",
    "http://www.jintiankansha.me/rss/GMZTSNZUGJ6GMMRRHFSTCMZWGUYDGZTCHFTGCZTGG4YTIYZVMMZDMMZQMFTGGNJVMQ2TEOBZMVSQ====",
    "http://www.jintiankansha.me/rss/GMZTOOJSGB6DOZLCGQZWCZRQGA2TEY3CGMZWCYJWGQ3TKYJTMRRDKM3FMJRDOZJUGYYDSMBUMQZQ====",
    "http://www.jintiankansha.me/rss/GEYDAOJZGB6DGNZRGBQTEMJQGMYGGYZVMU2DIOLDGQYDAN3EGRRTMZRVHA4DCZLFGAYWEMLDMIYA====",
    "http://www.jintiankansha.me/rss/GMZTQMZZHB6DMOJUGJSWIYJZMVTGEZDCMUZWGZJQMRQTMZRQGNQTEZDCHEZTGNRWMJQTIZRRGI2A====",
    "http://www.jintiankansha.me/rss/GM2DANJSGF6GCNLBGA3DIZRVMZQTSMRYGI4DOOJUMY4WEY3CMZQTSYJZGRRGKNBUGI2GEZDCMVRQ====",
    "http://www.jintiankansha.me/rss/GMZTONZSGB6GMNZWG43TENZZMEYGMNRXMMYTKMBSGQ4DSMJZGU4DEODCGA4DQZJYGM4DAYRYMMYQ====",
    "http://www.jintiankansha.me/rss/GM2DCOBTGJ6DKY3DMFTGEMRYMFQWENRUGAZTEYRTMYZGENDFGA2DQYJTGI3DOY3CGNQTMOJYMM4A====",
    "http://www.jintiankansha.me/rss/GMZTMNBRHB6DOMDBGFRTCZDCGRQTOMTDMY3TMMBSME2DQMRWGUZWMNLBMUYDCZRRG44DINZXHE3Q====",
    "http://www.jintiankansha.me/rss/GMZTONZSGF6DMZDFGU3DIMLCGIYWINRVGM3WCMZWGA3TSZBQGRRTCODCMNRWCMDGMI3DAZRUMFRQ====",
    "http://www.jintiankansha.me/rss/GMZTONZSGN6DMOBSHFSDSMLEGQYTEZBRHBRTQNDCGMZWIMBTGYYDOMBWG44DQY3FHAZGMOJUMI4A====",
    "http://www.jintiankansha.me/rss/GM2TSMRUGZ6DAMZSGY4DQNJRGZSDKMJUMQ4GEY3GG4ZGINDFMJRGGYRTMMYGGOBSHBTDGNDBMM4A====",
    "http://www.jintiankansha.me/rss/GEYDANZXHF6DCMDBMZSDGMDDMRTDMZRYGIYDANTEGA3TMYRTGAYGINLEGYZDIMLBGQYWEMBTHAZQ====",
    "https://hbdd2025.github.io/my-hexun-rss/hexun_insurance_rss.xml",
    "https://hbdd2025.github.io/nfra-rss-feed/nfra_rss.xml",
    "http://www.jintiankansha.me/rss/GEYTGOBYPQZDKNDGMQZDQMRWGA2TQZRRGU2DGMTEG4ZWMMRXME3TSODBGFSTENRQGBT",
    "http://www.jintiankansha.me/rss/GMZTONZRHF6DQZJSMY4GIMJQG5TGMZBWHA4WIMJWGUYTKNBYG5QTKYJVMQYGINJYGAYDEODEGFTA====",
]

FEED_SOURCES = {}  # 可选：不需要来源名的话留空

# ========== 简单工具 ==========
def user_agent():
    return {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36'}

def clean_html(content: str) -> str:
    if not content:
        return ""
    try:
        soup = BeautifulSoup(content, "lxml")
        for s in soup(["script", "style", "nav", "footer", "header", "aside", "form", "iframe", "button"]):
            s.decompose()
        for a in soup.find_all("a"):
            a.unwrap()
        return str(soup)
    except Exception:
        return content

def get_full_content(url: str):
    """尽力抓正文，失败就返回 None"""
    try:
        if not url or not url.startswith("http"):
            return None
        r = requests.get(url, headers=user_agent(), timeout=REQUEST_TIMEOUT, allow_redirects=True)
        r.raise_for_status()
        r.encoding = r.apparent_encoding or "utf-8"
        soup = BeautifulSoup(r.text, "lxml")
        selectors = [
            "article", ".article-content", ".entry-content", "div.rich_media_content",
            "div#js_content", "div.art_contextBox", "div.content", "div#zoom"
        ]
        for sel in selectors:
            node = soup.select_one(sel)
            if node and len(node.get_text(strip=True)) > 100:
                for img in node.find_all("img"):
                    src = img.get("src") or img.get("data-src")
                    if src and not src.startswith(("http", "data:")):
                        img["src"] = urljoin(r.url, src)
                return str(node)
    except Exception:
        pass
    return None

# ========== OneNote 认证与写入 ==========
class OneNoteSync:
    def __init__(self):
        self.cache_file = os.path.abspath(os.path.join(os.path.dirname(__file__), TOKEN_CACHE_FILE))
        try:
            persistence = FilePersistence(self.cache_file)
            self.token_cache = PersistedTokenCache(persistence)
            print(f"[缓存] {self.cache_file}")
        except Exception as e:
            print(f"[缓存警告] FilePersistence 失败: {e}，改用内存缓存。")
            self.token_cache = msal.SerializableTokenCache()

        self.app = PublicClientApplication(
            client_id=CLIENT_ID,
            authority=AUTHORITY,
            token_cache=self.token_cache
        )

    def get_token(self):
        if not CLIENT_ID:
            sys.exit("错误：AZURE_CLIENT_ID 未设置。")
        accounts = self.app.get_accounts()
        token_result = self.app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None

        if not token_result:
            print("[认证] 需要登录…")
            try:
                if (os.getenv("CI") or "").lower() == "true":
                    flow = self.app.initiate_device_flow(scopes=SCOPES)
                    if "user_code" not in flow:
                        print(f"[认证错误] 启动设备码失败: {flow}")
                        return None
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    print("!!! 请在 15 分钟内完成设备码登录 !!!")
                    print(flow["message"])
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    sys.stdout.flush()
                    token_result = self.app.acquire_token_by_device_flow(flow)
                else:
                    token_result = self.app.acquire_token_interactive(scopes=SCOPES)
            except Exception as e:
                print(f"[认证异常] {e}")
                return None

        if token_result and "access_token" in token_result:
            print("[认证] 成功获取访问令牌")
            return token_result["access_token"]
        else:
            print(f"[认证失败] {token_result}")
            return None

    def _call(self, method, url, headers=None, data=None):
        token = self.get_token()
        if not token:
            return None
        h = {"Authorization": f"Bearer {token}"}
        if headers:
            h.update(headers)
        try:
            r = requests.request(method, url, headers=h, data=data, timeout=REQUEST_TIMEOUT)
            r.raise_for_status()
            return r
        except requests.exceptions.HTTPError as e:
            print(f"[HTTP错误] {e.response.status_code} {e.response.reason} @ {url}")
            try:
                print("  详情:", e.response.json())
            except Exception:
                print("  响应文本:", e.response.text[:500])
            return None
        except Exception as e:
            print(f"[请求异常] {e}")
            return None

    def _resolve_section(self):
        """返回最后用于 Graph 的 sectionName（支持按月后缀），并 url-encode"""
        if not BASE_SECTION:
            return None
        section = BASE_SECTION
        if SECTION_MONTHLY:
            section = f"{BASE_SECTION} · {datetime.utcnow().strftime('%Y-%m')}"
        return quote(section, safe="")

    def create_page(self, title, html_body):
        if BASE_SECTION:
            section_q = self._resolve_section()
            url = f"https://graph.microsoft.com/v1.0/me/onenote/pages?sectionName={section_q}"
            print(f"[OneNote] 目标分区：{html.escape(quote(section_q))}（已 URL 编码）")
        else:
            url = "https://graph.microsoft.com/v1.0/me/onenote/pages"
            print("[OneNote] 未指定分区，将保存到默认位置")

        headers = {"Content-Type": "application/xhtml+xml"}
        safe_title = html.escape(title)
        created = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%S.000Z')
        doc = f'<!DOCTYPE html><html><head><title>{safe_title}</title>' \
              f'<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />' \
              f'<meta name="created" content="{created}" /></head><body>{html_body}</body></html>'
        r = self._call("POST", url, headers=headers, data=doc.encode("utf-8"))
        return r is not None and r.status_code in (201, 200)

# ========== RSS 抓取与去重 ==========
def load_processed():
    path = os.path.join(os.path.dirname(__file__), PROCESSED_ITEMS_FILE)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return set(x.strip() for x in f if x.strip())
    return set()

def save_processed(ids):
    path = os.path.join(os.path.dirname(__file__), PROCESSED_ITEMS_FILE)
    with open(path, "a", encoding="utf-8") as f:
        for i in ids:
            f.write(i + "\n")

def fetch_all_entries():
    all_entries = []
    for feed in ORIGINAL_FEEDS:
        print(f"  抓取：{feed}")
        d = feedparser.parse(feed, agent=user_agent()["User-Agent"])
        if getattr(d, "bozo", False):
            print("    × 解析失败，跳过")
            continue
        for e in d.entries:
            pid = e.get("id") or e.get("link")
            title = e.get("title", "无标题")
            link = e.get("link")
            published = None
            if hasattr(e, "published_parsed") and e.published_parsed:
                published = datetime.fromtimestamp(time.mktime(e.published_parsed))
            else:
                published = datetime.utcnow()
            summary = e.get("summary", "")
            all_entries.append({
                "id": pid, "title": title, "link": link,
                "time": published, "summary": summary, "feed": feed
            })
    all_entries.sort(key=lambda x: x["time"], reverse=True)
    return all_entries

# ========== 主流程 ==========
def main():
    print(f"=== 开始同步：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
    processed = load_processed()
    entries = fetch_all_entries()
    new_items = [e for e in entries if e["id"] not in processed]
    if not new_items:
        print("没有新的条目。")
        return 0

    to_process = new_items[:MAX_ITEMS_PER_RUN]
    print(f"发现 {len(new_items)} 条新条目，本次处理 {len(to_process)} 条。")

    on = OneNoteSync()
    if not on.get_token():
        print("错误：无法获取 OneNote 访问令牌。")
        return 1

    ok, fail = 0, 0
    done_ids = set()

    for it in to_process:
        print(f"\n--- 处理：{it['title'][:60]}")
        date_prefix = it["time"].strftime("%y%m%d")
        full_title = f"{date_prefix}-{it['title']}"

        full = get_full_content(it["link"])
        body = clean_html(full or it["summary"] or "")
        source_html = f'<p style="font-size:9pt;color:gray;">来源：{html.escape(it["feed"])} | ' \
                      f'<a href="{html.escape(it["link"])}">原文链接</a></p><hr/>'
        page_html = f"<h1>{html.escape(full_title)}</h1>{source_html}<div>{body}</div>"

        if on.create_page(full_title, page_html):
            ok += 1
            done_ids.add(it["id"])
            print("  >> 成功")
        else:
            fail += 1
            print("  >> 失败")
        time.sleep(REQUEST_DELAY)

    if done_ids:
        save_processed(done_ids)
    print(f"\n[统计] 成功 {ok} 条 / 失败 {fail} 条")
    # 即使有失败，也返回 0 让后续加密 & push 能执行
    return 0

if __name__ == "__main__":
    rc = main()
    sys.exit(rc)
