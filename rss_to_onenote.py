#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import time
import html
import re
import warnings
from datetime import datetime
from urllib.parse import urljoin, quote

import requests
from dotenv import load_dotenv

import msal
from msal import PublicClientApplication
from msal_extensions import PersistedTokenCache, FilePersistence

import feedparser
from bs4 import BeautifulSoup

# 关闭 requests 的某些噪声告警（可选）
from requests.packages.urllib3.exceptions import InsecureRequestWarning
warnings.simplefilter('ignore', InsecureRequestWarning)

# ========== 环境与常量 ==========
BASE_DIR = os.path.dirname(__file__)
env_path = os.path.join(BASE_DIR, '.env')
load_dotenv(env_path)

CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
ONENOTE_SECTION_NAME = os.getenv("ONENOTE_SECTION_NAME")  # 例如：RSS syn2（允许空格）
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Notes.ReadWrite.CreatedByApp"]

TOKEN_CACHE_FILE = "token_cache.bin"          # 工作流会将它加密后提交（*.enc）
PROCESSED_ITEMS_FILE = "processed_items.txt"  # 同上（*.enc）

MAX_ITEMS_PER_RUN = 50
REQUEST_TIMEOUT = 25
REQUEST_DELAY = 3  # 每篇之间的轻微延迟，避免触发限流

# ========== RSS 列表与映射 ==========
ORIGINAL_FEEDS = [
    "http://www.jintiankansha.me/rss/GIZDGNRVGJ6GIZRWGRRGMYZUGIYTOOBUMEYDSZTDMJQTEOJXHAZTGOBVGJRDMZJWMFSDSZJRHE3A====",  # 总局公众号
    "http://www.jintiankansha.me/rss/GEYDQMJQHF6DQOJRMZTDIZTGHAYDMZJSMYYWMOBUGM3GEZTGMQ2DIMBRHA3GGNLEGFSDENDEMEYQ====",  # 慧保天下
    "http://www.jintiankansha.me/rss/GMZTONZSGR6DMMZXGA4DQMZWGE3TAMTDMEYGGOJTGRSGMNLBGM4GGY3CGAZGENLGGQ3TMN3GMMZQ====",  # 十三精
    "http://www.jintiankansha.me/rss/GMZTSNZUGJ6GMMRRHFSTCMZWGUYDGZTCHFTGCZTGG4YTIYZVMMZDMMZQMFTGGNJVMQ2TEOBZMVSQ====",  # 蓝鲸
    "http://www.jintiankansha.me/rss/GMZTOOJSGB6DOZLCGQZWCZRQGA2TEY3CGMZWCYJWGQ3TKYJTMRRDKM3FMJRDOZJUGYYDSMBUMQZQ====",  # 中保新知
    "http://www.jintiankansha.me/rss/GEYDAOJZGB6DGNZRGBQTEMJQGMYGGYZVMU2DIOLDGQYDAN3EGRRTMZRVHA4DCZLFGAYWEMLDMIYA====",  # 保险一哥
    "http://www.jintiankansha.me/rss/GMZTQMZZHB6DMOJUGJSWIYJZMVTGEZDCMUZWGZJQMRQTMZRQGNQTEZDCHEZTGNRWMJQTIZRRGI2A====",  # 中国银行保险报
    "http://www.jintiankansha.me/rss/GM2DANJSGF6GCNLBGA3DIZRVMZQTSMRYGI4DOOJUMY4WEY3CMZQTSYJZGRRGKNBUGI2GEZDCMVRQ====",  # 保观
    "http://www.jintiankansha.me/rss/GMZTONZSGB6GMNZWG43TENZZMEYGMNRXMMYTKMBSGQ4DSMJZGU4DEODCGA4DQZJYGM4DAYRYMMYQ====",  # 保契
    "http://www.jintiankansha.me/rss/GM2DCOBTGJ6DKY3DMFTGEMRYMFQWENRUGAZTEYRTMYZGENDFGA2DQYJTGI3DOY3CGNQTMOJYMM4A====",  # 精算视觉
    "http://www.jintiankansha.me/rss/GMZTMNBRHB6DOMDBGFRTCZDCGRQTOMTDMY3TMMBSME2DQMRWGUZWMNLBMUYDCZRRG44DINZXHE3Q====",  # 中保学
    "http://www.jintiankansha.me/rss/GMZTONZSGF6DMZDFGU3DIMLCGIYWINRVGM3WCMZWGA3TSZBQGRRTCODCMNRWCMDGMI3DAZRUMFRQ====",  # 说点保
    "http://www.jintiankansha.me/rss/GMZTONZSGN6DMOBSHFSDSMLEGQYTEZBRHBRTQNDCGMZWIMBTGYYDOMBWG44DQY3FHAZGMOJUMI4A====",  # 今日保
    "http://www.jintiankansha.me/rss/GM2TSMRUGZ6DAMZSGY4DQNJRGZSDKMJUMQ4GEY3GG4ZGINDFMJRGGYRTMMYGGOBSHBTDGNDBMM4A====",  # 国寿研究声
    "http://www.jintiankansha.me/rss/GEYDANZXHF6DCMDBMZSDGMDDMRTDMZRYGIYDANTEGA3TMYRTGAYGINLEGYZDIMLBGQYWEMBTHAZQ====",  # 欣琦看金融
    "https://hbdd2025.github.io/my-hexun-rss/hexun_insurance_rss.xml",  # 和讯保险
    "https://hbdd2025.github.io/nfra-rss-feed/nfra_rss.xml",            # 总局官网
    "http://www.jintiankansha.me/rss/GEYTGOBYPQZDKNDGMQZDQMRWGA2TQZRRGU2DGMTEG4ZWMMRXME3TSODBGFSTENRQGBT",            # 中金点睛
    "http://www.jintiankansha.me/rss/GMZTONZRHF6DQZJSMY4GIMJQG5TGMZBWHA4WIMJWGUYTKNBYG5QTKYJVMQYGINJYGAYDEODEGFTA====",  # 保煎烩
]

FEED_SOURCES = {
    ORIGINAL_FEEDS[0]:  "总局公众号",
    ORIGINAL_FEEDS[1]:  "慧保天下",
    ORIGINAL_FEEDS[2]:  "十三精",
    ORIGINAL_FEEDS[3]:  "蓝鲸",
    ORIGINAL_FEEDS[4]:  "中保新知",
    ORIGINAL_FEEDS[5]:  "保险一哥",
    ORIGINAL_FEEDS[6]:  "中国银行保险报",
    ORIGINAL_FEEDS[7]:  "保观",
    ORIGINAL_FEEDS[8]:  "保契",
    ORIGINAL_FEEDS[9]:  "精算视觉",
    ORIGINAL_FEEDS[10]: "中保学",
    ORIGINAL_FEEDS[11]: "说点保",
    ORIGINAL_FEEDS[12]: "今日保",
    ORIGINAL_FEEDS[13]: "国寿研究声",
    ORIGINAL_FEEDS[14]: "欣琦看金融",
    ORIGINAL_FEEDS[15]: "和讯保险",
    ORIGINAL_FEEDS[16]: "总局官网",
    ORIGINAL_FEEDS[17]: "中金点睛",
    ORIGINAL_FEEDS[18]: "保煎烩",
}

# ========== 工具函数 ==========
def get_user_agent():
    return {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'}

def clean_extracted_html(html_content: str) -> str:
    """清理正文：去脚本/样式/导航等，去 a 包裹，去空段落"""
    if not html_content or not isinstance(html_content, str):
        return ""
    try:
        try:
            soup = BeautifulSoup(html_content, 'lxml')
        except Exception:
            soup = BeautifulSoup(html_content, 'html.parser')

        # 去掉不需要的标签/区域
        selectors = [
            'script', 'style', 'nav', 'footer', 'header', 'aside', 'form', 'iframe', 'button',
            '.sidebar', '#sidebar', '.related-posts', '.comments', '#comments',
            '.navigation', '.pagination', '.share-buttons', 'ins'
        ]
        for sel in selectors:
            for el in (soup.select(sel) if sel.startswith(('.', '#')) else soup.find_all(sel)):
                el.decompose()

        # 去除危险/冗余属性
        remove_attrs = {'style', 'class', 'id', 'onclick', 'onload', 'onerror', 'align', 'valign', 'bgcolor'}
        for tag in soup.find_all(True):
            if tag.attrs:
                tag.attrs = {k: v for k, v in tag.attrs.items() if k not in remove_attrs}

        # 去掉链接但保留文本
        for a in soup.find_all('a'):
            a.unwrap()

        # 移除空段落（无文字且无图）
        for p in soup.find_all('p'):
            if not p.get_text(strip=True) and not p.find('img'):
                p.decompose()

        return str(soup)
    except Exception as e:
        print(f"  >> [HTML 清理] 出错: {e}")
        return html_content

def get_full_content_from_link(url: str, source_name: str, feed_url: str, verify_ssl: bool = True):
    """访问原文页，尽力抽取正文与图片，并进行基本清洗。"""
    if not url or not url.startswith('http'):
        return None, None
    try:
        resp = requests.get(url, headers=get_user_agent(), timeout=REQUEST_TIMEOUT, allow_redirects=True, verify=verify_ssl)
        resp.raise_for_status()
        final_url = resp.url
        resp.encoding = resp.apparent_encoding or 'utf-8'

        try:
            soup = BeautifulSoup(resp.text, 'lxml')
        except Exception:
            soup = BeautifulSoup(resp.text, 'html.parser')

        # 可能的正文选择器（按优先级）
        selectors = []
        is_jintiankansha = feed_url and 'jintiankansha.me' in feed_url
        if source_name and source_name.startswith('总局'):
            selectors.extend(['div.content', 'div#zoom', '.pages_content', 'div.view.TRS_UEDITOR'])
        if source_name and (source_name == '和讯保险' or source_name.startswith('和讯-')):
            selectors.extend(['div.art_contextBox', 'div.art_context'])
        if is_jintiankansha:
            selectors.extend(['div.rich_media_content', 'div#js_content', 'div.wx-content'])

        # 通用回退
        selectors.extend([
            'article', '.article-content', '.entry-content', '.post-content', '.post-body',
            '#article-body', '#entry-content', 'div[itemprop="articleBody"]',
            'main', 'div.main-content', 'div.entry'
        ])

        article = None
        for sel in selectors:
            try:
                node = soup.select_one(sel)
                if node and (len(node.get_text(strip=True)) > 50 or node.find('img')):
                    article = node
                    break
            except Exception:
                continue

        if not article:
            return None, None

        # 修复图片相对路径 / data-src
        for img in article.find_all('img'):
            try:
                src = img.get('src') or img.get('data-src')
                if not src:
                    continue
                if not src.startswith(('http://', 'https://', 'data:')):
                    img['src'] = urljoin(final_url, src)
                elif 'data-src' in img.attrs and img['data-src'] and img['data-src'] != img.get('src', ''):
                    ds = img['data-src']
                    img['src'] = ds if ds.startswith(('http://', 'https://', 'data:')) else urljoin(final_url, ds)
            except Exception:
                continue

        cleaned = clean_extracted_html(str(article))
        return cleaned, None  # 这里不强求从网页再取时间，直接返回 None
    except requests.exceptions.Timeout:
        return "[错误：抓取原文链接超时]", None
    except requests.exceptions.RequestException as e:
        return f"[错误：抓取原文链接失败 - {e}]", None
    except Exception as e:
        return f"[错误：解析原文页面出错 - {e}]", None

def load_processed_items(filename: str = PROCESSED_ITEMS_FILE):
    processed = set()
    path = os.path.join(BASE_DIR, filename)
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                processed.update(line.strip() for line in f if line.strip())
            print(f"[状态] 已加载 {len(processed)} 条已处理 ID")
        except Exception as e:
            print(f"[状态] 读取 {filename} 失败: {e}")
    else:
        print(f"[状态] 未发现 {filename}（首次运行可能）")
    return processed

def save_processed_items(new_ids, filename: str = PROCESSED_ITEMS_FILE):
    if not new_ids:
        return
    path = os.path.join(BASE_DIR, filename)
    try:
        with open(path, 'a', encoding='utf-8') as f:
            f.writelines(f"{_id}\n" for _id in new_ids if _id)
        print(f"[状态] 已追加保存 {len(new_ids)} 条处理完成的 ID")
    except Exception as e:
        print(f"[状态] 写入 {filename} 失败: {e}")

def fetch_rss_feeds():
    print("\n[开始抓取 RSS 源...]")
    all_entries = []
    for feed_url in ORIGINAL_FEEDS:
        source_name = FEED_SOURCES.get(feed_url, "未知来源")
        print(f"  处理中: {source_name} ({feed_url})")
        try:
            feed = feedparser.parse(feed_url, agent=get_user_agent()['User-Agent'], request_headers=get_user_agent())
        except Exception as e:
            print(f"  >> 解析失败: {e}")
            continue

        if getattr(feed, 'bozo', False):
            print(f"  >> 警告: 解析异常: {feed.bozo_exception}")

        if not feed.entries:
            print("  >> 无条目")
            continue

        for entry in feed.entries:
            _id = entry.get('id') or entry.get('link') or f"{entry.get('title','无标题')}_{entry.get('updated', entry.get('published',''))}"
            title = (entry.get('title') or '无标题').strip()
            link = (entry.get('link') or '').strip()

            pub_parsed = entry.get('published_parsed') or entry.get('updated_parsed')
            pub_dt = None
            if pub_parsed:
                try:
                    pub_dt = datetime.fromtimestamp(time.mktime(pub_parsed))
                except Exception:
                    pub_dt = None

            # 内容摘要
            summary_html = ""
            if 'content' in entry and entry.content:
                html_part = next((c.value for c in entry.content if getattr(c, 'type', '') and 'html' in c.type.lower()), None)
                if html_part:
                    summary_html = html_part
                else:
                    text_part = next((c.value for c in entry.content if getattr(c, 'type', '') and 'text' in c.type.lower()), "")
                    if text_part:
                        summary_html = f"<p>{html.escape(text_part)}</p>"
            if not summary_html:
                summary_html = entry.get('summary', entry.get('description', ""))

            display_source_name = source_name

            # 和讯：有时 title 前半部是频道，可用于显示源名称
            try:
                if source_name == "和讯保险" and " | " in title:
                    channel = title.split(" | ", 1)[0].strip()
                    if channel:
                        display_source_name = f"和讯-{channel}"
            except Exception:
                pass

            all_entries.append({
                'id': _id,
                'title': title,
                'link': link,
                'published_time_rss': pub_dt,
                'content_summary': summary_html,
                'source_name': source_name,
                'display_source_name': display_source_name,
                'feed_url': feed_url
            })

    print(f"\n[RSS 完成] 共获取 {len(all_entries)} 条。按时间降序排序...")
    all_entries.sort(key=lambda x: x['published_time_rss'] or datetime.min, reverse=True)
    return all_entries

# ========== OneNote 适配类 ==========
class OneNoteSync:
    def __init__(self):
        self.cache_file_path = os.path.abspath(os.path.join(BASE_DIR, TOKEN_CACHE_FILE))
        try:
            persistence = FilePersistence(self.cache_file_path)
            print(f"[缓存] 使用文件持久化: {self.cache_file_path}")
        except Exception as e:
            print(f"[缓存错误] 初始化 FilePersistence 失败: {e}，改用内存缓存。")
            persistence = None

        self.token_cache = PersistedTokenCache(persistence) if persistence else msal.SerializableTokenCache()
        self.app = PublicClientApplication(
            client_id=CLIENT_ID,
            authority=AUTHORITY,
            token_cache=self.token_cache
        )

    def get_token(self):
        token_result = None
        accounts = self.app.get_accounts()
        if accounts:
            print(f"[认证] 找到缓存账户: {accounts[0].get('username', '未知用户')}")
            token_result = self.app.acquire_token_silent(SCOPES, account=accounts[0])

        if not token_result:
            print("[认证] 缓存未命中，准备认证...")
            try:
                if os.getenv('CI') == 'true':
                    # GitHub Actions：设备码流程
                    flow = self.app.initiate_device_flow(scopes=SCOPES)
                    if "message" not in flow:
                        print(f"[认证错误] 启动设备代码流程失败: {flow.get('error_description', '未知错误')}")
                        return None
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    print("!!! 需要人工操作：请在浏览器中完成设备登录 !!!")
                    print(flow["message"])
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    sys.stdout.flush()
                    token_result = self.app.acquire_token_by_device_flow(flow)
                else:
                    # 本地：弹出交互式登录
                    token_result = self.app.acquire_token_interactive(scopes=SCOPES)
            except Exception as e:
                print(f"[认证错误] 认证流程异常: {e}")
                return None

        if token_result and "access_token" in token_result:
            print("[认证] 成功获取访问令牌。")
            return token_result["access_token"]

        print(f"[认证失败] 未能获取访问令牌: {token_result.get('error_description', '未知错误') if token_result else '流程返回空'}")
        return None

    def _make_api_call(self, method, url, headers=None, json_data=None, data=None):
        token = self.get_token()
        if not token:
            print("[API] 因认证失败跳过调用。")
            return None

        req_headers = {"Authorization": f"Bearer {token}"}
        if headers:
            req_headers.update(headers)

        try:
            resp = requests.request(method, url, headers=req_headers, json=json_data, data=data, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            return resp
        except requests.exceptions.HTTPError as e:
            code = e.response.status_code if e.response is not None else "N/A"
            reason = e.response.reason if e.response is not None else "HTTPError"
            print(f"[API 错误] {code} {reason} -> {url}")
            try:
                print(f"  详情: {e.response.json()}")
            except Exception:
                if e.response is not None:
                    print(f"  响应文本: {e.response.text}")
            return None
        except Exception as e:
            print(f"[API 未知错误] {e}")
            return None

    def create_onenote_page_in_app_notebook(self, title: str, content_html: str):
        # 关键：分区名可能包含空格，必须 URL 编码
        if ONENOTE_SECTION_NAME:
            section = ONENOTE_SECTION_NAME.strip()
            url = f"https://graph.microsoft.com/v1.0/me/onenote/pages?sectionName={quote(section, safe='')}"
            print(f"[OneNote] 将尝试保存到指定分区: {section}")
        else:
            url = "https://graph.microsoft.com/v1.0/me/onenote/pages"
            print("[OneNote] 未指定分区，将保存到默认位置。")

        safe_title = html.escape(title)
        created_iso = time.strftime('%Y-%m-%dT%H:%M:%S.000Z', time.gmtime())

        # OneNote 要求 xhtml+xml
        payload = f"""<!DOCTYPE html>
<html lang="zh-CN">
  <head>
    <title>{safe_title}</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="created" content="{created_iso}" />
  </head>
  <body>
    {content_html}
  </body>
</html>""".strip()

        headers = {"Content-Type": "application/xhtml+xml"}
        resp = self._make_api_call("POST", url, headers=headers, data=payload.encode('utf-8'))
        return (resp is not None) and (resp.status_code == 201)

# ========== 主流程 ==========
if __name__ == "__main__":
    print(f"=== RSS to OneNote Sync 开始于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

    if not CLIENT_ID:
        sys.exit("错误：AZURE_CLIENT_ID 未能加载（请在仓库 Secrets 或 .env 中配置）。")

    # 载入已处理 ID
    processed_ids = load_processed_items()

    # 抓取 RSS
    entries = fetch_rss_feeds()
    new_entries = [e for e in entries if e['id'] not in processed_ids]
    print(f"\n[筛选] 新条目 {len(new_entries)} / 总计 {len(entries)}")

    if not new_entries:
        print("没有需要同步的新条目。")
        sys.exit(0)

    # 限流处理
    to_process = new_entries[:MAX_ITEMS_PER_RUN]
    print(f"[计划] 本次处理 {len(to_process)} 条（上限 {MAX_ITEMS_PER_RUN}）")

    # 初始化 OneNote
    onenote = OneNoteSync()
    if not onenote.get_token():
        sys.exit("错误：无法获取 OneNote 访问令牌。")

    success, failed = 0, 0
    done_ids = set()

    for idx, e in enumerate(to_process, 1):
        t0 = time.time()
        print(f"\n--- 处理 {idx}/{len(to_process)}: {e['title'][:60]} ---")
        print(f"  来源: {e.get('display_source_name', e['source_name'])} | 链接: {e['link']}")

        # 选择标题
        pub_dt = e.get('published_time_rss')
        date_prefix = (pub_dt or datetime.now()).strftime('%y%m%d')
        # 避免标题太长
        base_title = e['title'].strip()
        max_title_len = 180  # 预留日期前缀等
        if len(base_title) > max_title_len:
            base_title = base_title[:max_title_len] + "…"
        title = f"{date_prefix}-{base_title}"

        # 获取内容：优先原文抓取，失败回退 RSS 摘要
        content_html, _ = get_full_content_from_link(e['link'], e['source_name'], e['feed_url'], verify_ssl=True)
        if isinstance(content_html, str) and content_html.startswith("[错误："):
            print(f"  >> 抓取原文出错：{content_html}，回退至 RSS 摘要。")
            final_body = clean_extracted_html(e['content_summary'])
        elif content_html:
            final_body = content_html
        else:
            print("  >> 未能抓取原文，使用 RSS 摘要。")
            final_body = clean_extracted_html(e['content_summary'])

        if not final_body:
            final_body = "<p><i>[无法自动提取正文，请访问原文链接查看]</i></p>"

        pub_str = (pub_dt.strftime('%Y-%m-%d %H:%M:%S') if pub_dt else "未知")
        meta_src = html.escape(e.get('display_source_name', e['source_name']))
        meta_link = html.escape(e['link'] or '')

        page_html = f"""
<h1>{html.escape(title)}</h1>
<p style="font-size:9pt;color:gray;">发布时间: {pub_str} | 来源: {meta_src} | <a href="{meta_link}">原文链接</a></p>
<hr/>
<div>{final_body}</div>
        """.strip()

        ok = onenote.create_onenote_page_in_app_notebook(title, page_html)
        if ok:
            success += 1
            done_ids.add(e['id'])
            print(f"  >> 成功（耗时 {time.time() - t0:.2f}s）")
        else:
            failed += 1
            print("  !! 失败")

        time.sleep(REQUEST_DELAY)

    print(f"\n[同步统计] 成功 {success} 条，失败 {failed} 条。")
    if done_ids:
        save_processed_items(done_ids)

    print(f"\n=== 结束: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
