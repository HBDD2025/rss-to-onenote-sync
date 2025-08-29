#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RSS to OneNote 同步脚本（稳健来源名 + 分区名含空格OK）

- OneNote 分区名通过 sectionName 参数并用 urllib.parse.quote 编码
- 来源名优先：FEED_SOURCES 映射 -> RSS feed 自带标题 -> 兜底为域名
- 兼容 GitHub Actions：
    - CI=true 时使用 Device Code（日志中给出设备码，需要人工在15分钟内完成一次授权）
    - 本地运行自动弹出交互式登录
- token_cache.bin / processed_items.txt 由工作流加密后再提交，脚本本身只读写明文文件
"""

import os
import sys
import time
import html
import warnings
from datetime import datetime
from urllib.parse import urljoin, quote, urlparse

import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning

import feedparser
from bs4 import BeautifulSoup

from dotenv import load_dotenv
import msal
from msal import PublicClientApplication
from msal_extensions import PersistedTokenCache, FilePersistence

warnings.simplefilter('ignore', InsecureRequestWarning)

# ========== 环境&常量 ==========
env_path = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv(env_path)

CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
ONENOTE_SECTION_NAME = os.getenv("ONENOTE_SECTION_NAME")

AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Notes.ReadWrite.CreatedByApp"]

TOKEN_CACHE_FILE = "token_cache.bin"         # 工作流会把它加密为 *.enc 再提交
PROCESSED_ITEMS_FILE = "processed_items.txt" # 同上
MAX_ITEMS_PER_RUN = 50
REQUEST_TIMEOUT = 25
REQUEST_DELAY = 3

# ====== RSS 源（与您仓库保持一致，可自行增减）======
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

# ====== 来源映射（键需要与 feed URL 完全一致；也会有更稳健兜底）======
FEED_SOURCES = {
    "http://www.jintiankansha.me/rss/GIZDGNRVGJ6GIZRWGRRGMYZUGIYTOOBUMEYDSZTDMJQTEOJXHAZTGOBVGJRDMZJWMFSDSZJRHE3A====": "总局公众号",
    "http://www.jintiankansha.me/rss/GEYDQMJQHF6DQOJRMZTDIZTGHAYDMZJSMYYWMOBUGM3GEZTGMQ2DIMBRHA3GGNLEGFSDENDEMEYQ====": "慧保天下",
    "http://www.jintiankansha.me/rss/GMZTONZSGR6DMMZXGA4DQMZWGE3TAMTDMEYGGOJTGRSGMNLBGM4GGY3CGAZGENLGGQ3TMN3GMMZQ====": "十三精",
    "http://www.jintiankansha.me/rss/GMZTSNZUGJ6GMMRRHFSTCMZWGUYDGZTCHFTGCZTGG4YTIYZVMMZDMMZQMFTGGNJVMQ2TEOBZMVSQ====": "蓝鲸",
    "http://www.jintiankansha.me/rss/GMZTOOJSGB6DOZLCGQZWCZRQGA2TEY3CGMZWCYJWGQ3TKYJTMRRDKM3FMJRDOZJUGYYDSMBUMQZQ====": "中保新知",
    "http://www.jintiankansha.me/rss/GEYDAOJZGB6DGNZRGBQTEMJQGMYGGYZVMU2DIOLDGQYDAN3EGRRTMZRVHA4DCZLFGAYWEMLDMIYA====": "保险一哥",
    "http://www.jintiankansha.me/rss/GMZTQMZZHB6DMOJUGJSWIYJZMVTGEZDCMUZWGZJQMRQTMZRQGNQTEZDCHEZTGNRWMJQTIZRRGI2A====": "中国银行保险报",
    "http://www.jintiankansha.me/rss/GM2DANJSGF6GCNLBGA3DIZRVMZQTSMRYGI4DOOJUMY4WEY3CMZQTSYJZGRRGKNBUGI2GEZDCMVRQ====": "保观",
    "http://www.jintiankansha.me/rss/GMZTONZSGB6GMNZWG43TENZZMEYGMNRXMMYTKMBSGQ4DSMJZGU4DEODCGA4DQZJYGM4DAYRYMMYQ====": "保契",
    "http://www.jintiankansha.me/rss/GM2DCOBTGJ6DKY3DMFTGEMRYMFQWENRUGAZTEYRTMYZGENDFGA2DQYJTGI3DOY3CGNQTMOJYMM4A====": "精算视觉",
    "http://www.jintiankansha.me/rss/GMZTMNBRHB6DOMDBGFRTCZDCGRQTOMTDMY3TMMBSME2DQMRWGUZWMNLBMUYDCZRRG44DINZXHE3Q====": "中保学",
    "http://www.jintiankansha.me/rss/GMZTONZSGF6DMZDFGU3DIMLCGIYWINRVGM3WCMZWGA3TSZBQGRRTCODCMNRWCMDGMI3DAZRUMFRQ====": "说点保",
    "http://www.jintiankansha.me/rss/GMZTONZSGN6DMOBSHFSDSMLEGQYTEZBRHBRTQNDCGMZWIMBTGYYDOMBWG44DQY3FHAZGMOJUMI4A====": "今日保",
    "http://www.jintiankansha.me/rss/GM2TSMRUGZ6DAMZSGY4DQNJRGZSDKMJUMQ4GEY3GG4ZGINDFMJRGGYRTMMYGGOBSHBTDGNDBMM4A====": "国寿研究声",
    "http://www.jintiankansha.me/rss/GEYDANZXHF6DCMDBMZSDGMDDMRTDMZRYGIYDANTEGA3TMYRTGAYGINLEGYZDIMLBGQYWEMBTHAZQ====": "欣琦看金融",
    "https://hbdd2025.github.io/my-hexun-rss/hexun_insurance_rss.xml": "和讯保险",
    "https://hbdd2025.github.io/nfra-rss-feed/nfra_rss.xml": "总局官网",
    "http://www.jintiankansha.me/rss/GEYTGOBYPQZDKNDGMQZDQMRWGA2TQZRRGU2DGMTEG4ZWMMRXME3TSODBGFSTENRQGBT": "中金点睛",
    "http://www.jintiankansha.me/rss/GMZTONZRHF6DQZJSMY4GIMJQG5TGMZBWHA4WIMJWGUYTKNBYG5QTKYJVMQYGINJYGAYDEODEGFTA====": "保煎烩",
}

# ====== 工具函数 ======
def get_user_agent():
    return {
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/120.0.0.0 Safari/537.36'
        )
    }

def clean_extracted_html(html_content: str) -> str:
    if not html_content or not isinstance(html_content, str):
        return ""
    try:
        soup = BeautifulSoup(html_content, 'lxml')
        # 移除无关元素
        tags_to_remove = [
            'script', 'style', 'nav', 'footer', 'header',
            'aside', 'form', 'iframe', 'button',
            '.sidebar', '#sidebar', '.related-posts',
            '.comments', '#comments'
        ]
        for sel in tags_to_remove:
            for el in soup.select(sel):
                el.decompose()
        # 取消 <a> 包裹
        for a in soup.find_all('a'):
            a.unwrap()
        return str(soup)
    except Exception as e:
        print(f"  >> HTML 清理出错: {e}")
        return html_content

def get_full_content_from_link(url: str, feed_url: str):
    """尝试抓取原文正文；失败则返回 (None, None)"""
    if not url or not url.startswith('http'):
        return None, None
    try:
        resp = requests.get(url, headers=get_user_agent(), timeout=REQUEST_TIMEOUT, allow_redirects=True)
        resp.raise_for_status()
        final_url = resp.url
        resp.encoding = resp.apparent_encoding or 'utf-8'
        soup = BeautifulSoup(resp.text, 'lxml')

        article_body = None
        selectors = [
            'div.rich_media_content', 'div#js_content',
            'div.art_contextBox', 'div.art_context',
            'div.content', 'div#zoom',
            'article', '.article-content', '.entry-content',
        ]
        for sel in selectors:
            t = soup.select_one(sel)
            if t and len(t.get_text(strip=True)) > 100:
                article_body = t
                break
        if not article_body:
            return None, None

        # 修复相对图片
        for img in article_body.find_all('img'):
            src = img.get('src') or img.get('data-src')
            if src and not src.startswith(('http', 'data:')):
                img['src'] = urljoin(final_url, src)

        return str(article_body), None
    except Exception:
        return None, None

def _normalize_url_for_match(u: str) -> str:
    """用于宽松匹配：小写 scheme/host，去掉末尾斜杠"""
    try:
        p = urlparse(u)
        host = (p.netloc or "").lower()
        scheme = (p.scheme or "").lower()
        path = (p.path or "").rstrip("/")
        return f"{scheme}://{host}{path}" if scheme and host else u
    except Exception:
        return u

def _source_name_for(feed_url: str, feed_data) -> str:
    """来源名优先：映射 -> RSS 标题 -> 域名"""
    # 1) 精确映射
    if feed_url in FEED_SOURCES:
        return FEED_SOURCES[feed_url]

    # 2) 宽松匹配一次（处理 http/https、末尾斜杠）
    norm = _normalize_url_for_match(feed_url)
    for k, v in FEED_SOURCES.items():
        if _normalize_url_for_match(k) == norm:
            return v

    # 3) 用 RSS feed 自带标题
    try:
        title = (getattr(feed_data, "feed", {}) or {}).get("title")
        if title and isinstance(title, str) and title.strip():
            return title.strip()
    except Exception:
        pass

    # 4) 兜底：域名
    host = urlparse(feed_url).netloc.replace("www.", "")
    return host or "未知来源"

def fetch_rss_feeds():
    all_entries = []
    for feed_url in ORIGINAL_FEEDS:
        print(f"  正在处理源: {feed_url}")
        feed_data = feedparser.parse(feed_url, agent=get_user_agent()['User-Agent'])
        if feed_data.bozo:
            print("    >> 该源解析异常，已跳过。")
            continue

        source_name = _source_name_for(feed_url, feed_data)

        for entry in feed_data.entries:
            pub = None
            if hasattr(entry, 'published_parsed') and entry.published_parsed:
                try:
                    pub = datetime.fromtimestamp(time.mktime(entry.published_parsed))
                except Exception:
                    pub = datetime.now()
            else:
                pub = datetime.now()

            all_entries.append({
                'id': entry.get('id', entry.link),
                'title': entry.title,
                'link': entry.link,
                'published_time_rss': pub,
                'content_summary': entry.get('summary', ''),
                'source_name': source_name,
                'feed_url': feed_url
            })

    # 新到旧
    all_entries.sort(key=lambda x: x['published_time_rss'], reverse=True)
    return all_entries

def load_processed_items(filename=PROCESSED_ITEMS_FILE):
    fp = os.path.join(os.path.dirname(__file__), filename)
    if os.path.exists(fp):
        with open(fp, 'r', encoding='utf-8') as f:
            return set(line.strip() for line in f if line.strip())
    return set()

def save_processed_items(new_ids, filename=PROCESSED_ITEMS_FILE):
    fp = os.path.join(os.path.dirname(__file__), filename)
    with open(fp, 'a', encoding='utf-8') as f:
        f.writelines(f"{i}\n" for i in new_ids)

# ====== OneNote 交互 ======
class OneNoteSync:
    def __init__(self):
        self.cache_file_path = os.path.abspath(os.path.join(os.path.dirname(__file__), TOKEN_CACHE_FILE))
        try:
            persistence = FilePersistence(self.cache_file_path)
            print(f"[缓存] 使用文件持久化: {self.cache_file_path}")
        except Exception as e:
            print(f"[缓存错误] 初始化 FilePersistence 失败: {e}。改用内存缓存。")
            persistence = None

        self.token_cache = PersistedTokenCache(persistence) if persistence else msal.SerializableTokenCache()
        self.app = PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY, token_cache=self.token_cache)

    def get_token(self):
        token_result = None
        accounts = self.app.get_accounts()
        if accounts:
            print(f"[认证] 找到缓存账户: {accounts[0].get('username', '未知')}")
            token_result = self.app.acquire_token_silent(SCOPES, account=accounts[0])

        if not token_result:
            print("[认证] 缓存未命中，需要人工登录一次...")
            try:
                if os.getenv('CI') == 'true':
                    flow = self.app.initiate_device_flow(scopes=SCOPES)
                    if "message" not in flow:
                        print(f"[认证错误] 启动设备码流程失败: {flow.get('error_description', '未知错误')}")
                        return None
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    print("!!! 需要人工操作：请在浏览器中完成设备登录 !!!")
                    print(flow["message"])
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    sys.stdout.flush()
                    token_result = self.app.acquire_token_by_device_flow(flow)
                else:
                    token_result = self.app.acquire_token_interactive(scopes=SCOPES)
            except Exception as e:
                print(f"[认证错误] 登录流程异常: {e}")
                return None

        if token_result and "access_token" in token_result:
            print("[认证] 成功获取访问令牌。")
            return token_result["access_token"]
        else:
            print(f"[认证失败] 未能获取访问令牌: {token_result.get('error_description', '未知错误') if token_result else '空响应'}")
            return None

    def _make_api_call(self, method, url, headers=None, json_data=None, data=None):
        token = self.get_token()
        if not token:
            print("[API] 因无法获取令牌而跳过 API 调用。")
            return None

        req_headers = {"Authorization": f"Bearer {token}"}
        if headers:
            req_headers.update(headers)

        try:
            resp = requests.request(method, url, headers=req_headers, json=json_data, data=data, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            return resp
        except requests.exceptions.HTTPError as e:
            print(f"[API HTTP错误] {e.response.status_code} {e.response.reason} {url}")
            try:
                print(f"  详情: {e.response.json()}")
            except Exception:
                print(f"  文本: {e.response.text}")
            return None
        except Exception as e:
            print(f"[API 未知错误] {e}")
            return None

    def create_onenote_page_in_app_notebook(self, title: str, content_html: str) -> bool:
        if ONENOTE_SECTION_NAME:
            section = ONENOTE_SECTION_NAME.strip()
            url = f"https://graph.microsoft.com/v1.0/me/onenote/pages?sectionName={quote(section, safe='')}"
            print(f"[OneNote] 指定分区: {section}")
        else:
            url = "https://graph.microsoft.com/v1.0/me/onenote/pages"
            print("[OneNote] 未指定分区，保存到默认位置。")

        headers = {"Content-Type": "application/xhtml+xml"}
        safe_title = html.escape(title)
        creation_time_str = time.strftime('%Y-%m-%dT%H:%M:%S.000Z', time.gmtime())

        onenote_page_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <title>{safe_title}</title>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <meta name="created" content="{creation_time_str}" />
</head>
<body>
{content_html}
</body>
</html>"""

        resp = self._make_api_call("POST", url, headers=headers, data=onenote_page_content.encode('utf-8'))
        return resp is not None and resp.status_code == 201

# ====== 主流程 ======
if __name__ == "__main__":
    print(f"=== RSS to OneNote Sync 开始于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

    if not CLIENT_ID:
        sys.exit("错误：AZURE_CLIENT_ID 未配置。")

    processed_ids = load_processed_items()
    all_entries = fetch_rss_feeds()
    new_entries = [e for e in all_entries if e['id'] not in processed_ids]

    if not new_entries:
        print("没有需要同步的新条目。")
        sys.exit(0)

    entries_to_process = new_entries[:MAX_ITEMS_PER_RUN]
    print(f"发现 {len(new_entries)} 条新条目，本次处理 {len(entries_to_process)} 条。")

    onenote = OneNoteSync()
    if not onenote.get_token():
        sys.exit("错误：无法获取 OneNote 访问令牌。")

    success, fail = 0, 0
    newly_processed_ids_in_run = set()

    for entry in entries_to_process:
        print(f"\n--- 正在处理: {entry['title'][:50]} ...")

        date_prefix = entry['published_time_rss'].strftime('%y%m%d')
        formatted_title = f"{date_prefix}-{entry['title']}"

        full_content, _ = get_full_content_from_link(entry['link'], entry['feed_url'])
        final_body_html = clean_extracted_html(full_content or entry['content_summary'])

        # 来源名 + 原文链接（来源名已经稳健，不会再出现整段URL）
        source_name = html.escape(entry['source_name'])
        onenote_content = f"""
<h1>{html.escape(formatted_title)}</h1>
<p style="font-size:9pt; color:gray;">
  来源: {source_name} | <a href="{html.escape(entry['link'])}">原文链接</a>
</p>
<hr/>
<div>{final_body_html}</div>
"""

        if onenote.create_onenote_page_in_app_notebook(formatted_title, onenote_content):
            success += 1
            newly_processed_ids_in_run.add(entry['id'])
            print("  >> 同步成功。")
        else:
            fail += 1
            print("  >> 同步失败。")

        time.sleep(REQUEST_DELAY)

    if newly_processed_ids_in_run:
        save_processed_items(newly_processed_ids_in_run)

    print(f"\n[同步统计] 成功 {success} 条，失败 {fail} 条。")
