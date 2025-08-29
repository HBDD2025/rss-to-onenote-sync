#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os, sys, time, html, re, warnings
from datetime import datetime
from urllib.parse import urljoin, quote

import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning
warnings.simplefilter('ignore', InsecureRequestWarning)

import feedparser
from bs4 import BeautifulSoup

import msal
from msal import PublicClientApplication
from msal_extensions import PersistedTokenCache, FilePersistence
from dotenv import load_dotenv

# ================= 配置区 =================
BASE_DIR = os.path.dirname(__file__)
env_path = os.path.join(BASE_DIR, '.env')
if os.path.exists(env_path): load_dotenv(env_path)

CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
ONENOTE_SECTION_NAME = os.getenv("ONENOTE_SECTION_NAME", "").strip()
SECTION_MONTHLY = os.getenv("SECTION_MONTHLY", "false").lower() == "true"

# 关键：加入 offline_access，让刷新令牌可复用
SCOPES = ["Notes.ReadWrite.CreatedByApp"]
AUTHORITY = "https://login.microsoftonline.com/consumers"

# 状态文件（明文，只在 runner 上用；仓库里是 .enc）
TOKEN_CACHE_FILE = "token_cache.bin"
PROCESSED_ITEMS_FILE = "processed_items.txt"

# 每次处理条数：先读环境变量，默认 50；你现在想测 2 就在 workflow 里传 2
MAX_ITEMS_PER_RUN = int(os.getenv("MAX_ITEMS_PER_RUN", "50"))
REQUEST_TIMEOUT = 25
REQUEST_DELAY = 3

# 你的源 & 名称映射（按你之前的版本放全即可；此处保留示例）
ORIGINAL_FEEDS = [
    "https://hbdd2025.github.io/my-hexun-rss/hexun_insurance_rss.xml",
    "https://hbdd2025.github.io/nfra-rss-feed/nfra_rss.xml",
]
FEED_SOURCES = {
    "https://hbdd2025.github.io/my-hexun-rss/hexun_insurance_rss.xml": "和讯保险",
    "https://hbdd2025.github.io/nfra-rss-feed/nfra_rss.xml": "总局官网",
}
# =========================================

def get_user_agent():
    return {
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/120.0 Safari/537.36'
        )
    }

def clean_extracted_html(html_content):
    if not html_content or not isinstance(html_content, str): return ""
    try:
        soup = BeautifulSoup(html_content, 'lxml')
        for sel in ['script','style','nav','footer','header','aside','form','iframe','button',
                    '.sidebar','#sidebar','.related-posts','.comments','#comments']:
            for e in soup.select(sel):
                e.decompose()
        for a in soup.find_all('a'):
            a.unwrap()
        return str(soup)
    except Exception as e:
        print(f"  >> HTML 清理出错: {e}")
        return html_content

def get_full_content_from_link(url, _feed_url):
    if not url or not url.startswith('http'): return None, None
    try:
        r = requests.get(url, headers=get_user_agent(), timeout=REQUEST_TIMEOUT, allow_redirects=True)
        r.raise_for_status()
        r.encoding = r.apparent_encoding or 'utf-8'
        final_url = r.url
        soup = BeautifulSoup(r.text, 'lxml')
        selectors = [
            'div.rich_media_content','div#js_content','div.art_contextBox','div.art_context',
            'div.content','div#zoom','article','.article-content','.entry-content'
        ]
        article = None
        for sel in selectors:
            t = soup.select_one(sel)
            if t and len(t.get_text(strip=True)) > 100:
                article = t; break
        if not article: return None, None
        for img in article.find_all('img'):
            src = img.get('src') or img.get('data-src')
            if src and not src.startswith(('http','data:')):
                img['src'] = urljoin(final_url, src)
        return str(article), None
    except Exception:
        return None, None

def fetch_rss_feeds():
    all_entries = []
    for feed_url in ORIGINAL_FEEDS:
        source_name = FEED_SOURCES.get(feed_url) or feed_url  # 兜底：没有映射就显示 URL
        print(f"  正在处理: {source_name}")
        fd = feedparser.parse(feed_url, agent=get_user_agent()['User-Agent'])
        if fd.bozo: 
            print(f"  >> 解析失败，跳过: {feed_url}"); 
            continue
        for entry in fd.entries:
            all_entries.append({
                'id': entry.get('id', entry.link),
                'title': entry.title,
                'link': entry.link,
                'published_time_rss': datetime.fromtimestamp(time.mktime(entry.published_parsed)) if hasattr(entry,'published_parsed') else datetime.now(),
                'content_summary': entry.get('summary', ''),
                'source_name': source_name,
                'feed_url': feed_url
            })
    all_entries.sort(key=lambda x: x['published_time_rss'], reverse=True)
    return all_entries

def load_processed_items(filename=PROCESSED_ITEMS_FILE):
    p = os.path.join(BASE_DIR, filename)
    if os.path.exists(p):
        with open(p, 'r', encoding='utf-8') as f:
            return set(line.strip() for line in f if line.strip())
    return set()

def save_processed_items(new_ids, filename=PROCESSED_ITEMS_FILE):
    p = os.path.join(BASE_DIR, filename)
    with open(p, 'a', encoding='utf-8') as f:
        f.writelines(f"{i}\n" for i in new_ids)

class OneNoteSync:
    def __init__(self):
        self.cache_path = os.path.abspath(os.path.join(BASE_DIR, TOKEN_CACHE_FILE))
        print(f"[缓存] 使用文件持久化: {self.cache_path}")
        try:
            persistence = FilePersistence(self.cache_path)
        except Exception as e:
            print(f"[缓存错误] 初始化 FilePersistence 失败: {e}，改用内存缓存")
            persistence = None
        self.token_cache = PersistedTokenCache(persistence) if persistence else msal.SerializableTokenCache()
        self.app = PublicClientApplication(
            client_id=CLIENT_ID, authority=AUTHORITY, token_cache=self.token_cache
        )

    def _persist_cache(self, tag=""):
        # 关键：强制把缓存落盘
        try:
            if hasattr(self.token_cache, "has_state_changed") and self.token_cache.has_state_changed:
                self.token_cache.save()
                print(f"[缓存] 已保存 {tag} 到 {self.cache_path}")
        except Exception as e:
            print(f"[缓存错误] 保存失败: {e}")

    def get_token(self):
        if not CLIENT_ID:
            print("错误：AZURE_CLIENT_ID 未设置"); return None

        accounts = self.app.get_accounts()
        if accounts:
            print(f"[认证] 找到账户缓存: {accounts[0].get('username','未知')}")
            result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._persist_cache("silent")
                print("[认证] 缓存命中，已获取访问令牌。")
                return result["access_token"]

        print("[认证] 缓存未命中，需要人工登录一次...")
        try:
            if os.getenv("CI","").lower() == "true":
                flow = self.app.initiate_device_flow(scopes=SCOPES)
                if "message" not in flow:
                    print(f"[认证错误] 启动设备代码流程失败: {flow.get('error_description','未知错误')}")
                    return None
                print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                print("!!! 需要人工操作：请在浏览器中完成设备登录 !!!")
                print(flow["message"])
                print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                sys.stdout.flush()
                result = self.app.acquire_token_by_device_flow(flow)
            else:
                result = self.app.acquire_token_interactive(scopes=SCOPES)
        except Exception as e:
            print(f"[认证错误] 认证流程异常: {e}")
            return None

        if result and "access_token" in result:
            print("[认证] 成功获取访问令牌。")
            self._persist_cache("interactive/device_flow")
            return result["access_token"]

        print(f"[认证失败] 未能获取访问令牌: {result.get('error_description','未知错误') if result else '空返回'}")
        return None

    def _api(self, method, url, headers=None, json_data=None, data=None):
        token = self.get_token()
        if not token:
            print("[API] 因无法获取令牌而跳过调用。"); return None
        h = {"Authorization": f"Bearer {token}"}
        if headers: h.update(headers)
        try:
            r = requests.request(method, url, headers=h, json=json_data, data=data, timeout=REQUEST_TIMEOUT)
            r.raise_for_status(); return r
        except requests.exceptions.HTTPError as e:
            print(f"[API HTTP错误] {e.response.status_code} {e.response.reason} url={url}")
            try: print("  详情:", e.response.json())
            except: print("  文本:", e.response.text)
            return None
        except Exception as e:
            print(f"[API 未知错误] {e}"); return None

    def create_page(self, title, content_html):
        if ONENOTE_SECTION_NAME:
            section = ONENOTE_SECTION_NAME
            if SECTION_MONTHLY:
                # 形如：RSS syn2 2025-08
                y_m = datetime.utcnow().strftime("%Y-%m")
                section = f"{section} {y_m}"
            url = f"https://graph.microsoft.com/v1.0/me/onenote/pages?sectionName={quote(section, safe='')}"
            print(f"[OneNote] 将尝试保存到分区: {section}")
        else:
            url = "https://graph.microsoft.com/v1.0/me/onenote/pages"
            print("[OneNote] 未指定分区，将保存到默认位置。")

        headers = {"Content-Type": "application/xhtml+xml"}
        safe_title = html.escape(title)
        created = time.strftime('%Y-%m-%dT%H:%M:%S.000Z', time.gmtime())
        page = f"""<!DOCTYPE html><html lang="zh-CN"><head><title>{safe_title}</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="created" content="{created}" /></head><body>{content_html}</body></html>"""
        r = self._api("POST", url, headers=headers, data=page.encode('utf-8'))
        return (r is not None) and (r.status_code == 201)

def build_body(entry, full_html):
    title = f"{entry['published_time_rss'].strftime('%y%m%d')}-{entry['title']}"
    src = html.escape(entry['source_name'])
    link = html.escape(entry['link'])
    body = f"""
        <h1>{html.escape(title)}</h1>
        <p style="font-size:9pt;color:gray;">来源: {src} | <a href="{link}">原文链接</a></p><hr/>
        <div>{full_html}</div>
    """
    return title, body

if __name__ == "__main__":
    print(f"=== RSS to OneNote Sync 开始于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
    if not CLIENT_ID:
        sys.exit("错误：AZURE_CLIENT_ID 未能加载。")

    processed = load_processed_items()
    entries = fetch_rss_feeds()
    new_items = [e for e in entries if e['id'] not in processed]
    if not new_items:
        print("没有需要同步的新条目。"); sys.exit(0)

    batch = new_items[:MAX_ITEMS_PER_RUN]
    print(f"发现 {len(new_items)} 条新条目，本次处理 {len(batch)} 条。")

    sync = OneNoteSync()
    if not sync.get_token(): sys.exit("错误：无法获取 OneNote 访问令牌。")

    ok, fail = 0, 0
    done_ids = set()
    for e in batch:
        print(f"\n--- 正在处理: {e['title'][:60]} ...")
        full_html, _ = get_full_content_from_link(e['link'], e['feed_url'])
        final_html = clean_extracted_html(full_html or e['content_summary'])
        title, body = build_body(e, final_html)
        if sync.create_page(title, body):
            ok += 1; done_ids.add(e['id']); print("  >> 成功同步到 OneNote。")
        else:
            fail += 1; print("  >> 同步失败。")
        time.sleep(REQUEST_DELAY)

    if done_ids: save_processed_items(done_ids)
    print(f"\n[同步统计] 成功 {ok} 条，失败 {fail} 条。")
