#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RSS -> OneNote 同步脚本（加固版）
改进点：
- 设备码登录缓存：token_cache.bin（工作流会加密保存为 token_cache.enc）
- 支持 OneNote 分区按月分卷（避免 507 满额）：ONENOTE_SECTION_NAME + 当月后缀
- API 调用带重试（429 / 5xx 指数回退）、速率控制、详细日志
- 抓取正文时做 HTML 清理、图片绝对化
- processed_items.txt 记录去重（工作流加密保存为 processed_items.enc）

环境变量（由工作流注入）:
- AZURE_CLIENT_ID（必填，GitHub Secrets）
- ONENOTE_SECTION_NAME（建议填写，例如 “RSS syn2”，GitHub Secrets）
- CI=true（工作流自动注入）
可选控制：
- SECTION_MONTHLY=true/false（默认 true：按月分卷）
- MAX_ITEMS_PER_RUN=50（默认 50）
- REQUEST_TIMEOUT=25（秒）
- REQUEST_DELAY=3（秒）
"""

import os
import sys
import time
import html
import re
import warnings
from urllib.parse import urljoin, quote
from datetime import datetime, timezone

import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from dotenv import load_dotenv
import msal
from msal import PublicClientApplication
from msal_extensions import PersistedTokenCache, FilePersistence
import feedparser
from bs4 import BeautifulSoup

warnings.simplefilter('ignore', InsecureRequestWarning)

# ---------------- 基本配置 ----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
env_path = os.path.join(BASE_DIR, '.env')
load_dotenv(env_path)  # 本地调试时可用，Actions 里依赖 Secrets

CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Notes.ReadWrite.CreatedByApp"]

TOKEN_CACHE_FILE = os.path.join(BASE_DIR, "token_cache.bin")
PROCESSED_ITEMS_FILE = os.path.join(BASE_DIR, "processed_items.txt")

# 运行参数（支持环境变量覆盖）
MAX_ITEMS_PER_RUN = int(os.getenv("MAX_ITEMS_PER_RUN", "50"))
REQUEST_TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "25"))
REQUEST_DELAY = float(os.getenv("REQUEST_DELAY", "3"))

ONENOTE_SECTION_BASE = (os.getenv("ONENOTE_SECTION_NAME") or "").strip()  # 建议在 Secrets 里设为 "RSS syn2"
SECTION_MONTHLY = (os.getenv("SECTION_MONTHLY", "true").lower() != "false")  # 默认 true：按月分卷

# ---------------- RSS 源 ----------------
ORIGINAL_FEEDS = [
    "http://www.jintiankansha.me/rss/GIZDGNRVGJ6GIZRWGRRGMYZUGIYTOOBUMEYDSZTDMJQTEOJXHAZTGOBVGJRDMZJWMFSDSZJRHE3A====", # 总局公众号
    "http://www.jintiankansha.me/rss/GEYDQMJQHF6DQOJRMZTDIZTGHAYDMZJSMYYWMOBUGM3GEZTGMQ2DIMBRHA3GGNLEGFSDENDEMEYQ====", # 慧保天下
    "http://www.jintiankansha.me/rss/GMZTONZSGR6DMMZXGA4DQMZWGE3TAMTDMEYGGOJTGRSGMNLBGM4GGY3CGAZGENLGGQ3TMN3GMMZQ====", # 十三精
    "http://www.jintiankansha.me/rss/GMZTSNZUGJ6GMMRRHFSTCMZWGUYDGZTCHFTGCZTGG4YTIYZVMMZDMMZQMFTGGNJVMQ2TEOBZMVSQ====", # 蓝鲸
    "http://www.jintiankansha.me/rss/GMZTOOJSGB6DOZLCGQZWCZRQGA2TEY3CGMZWCYJWGQ3TKYJTMRRDKM3FMJRDOZJUGYYDSMBUMQZQ====", # 中保新知
    "http://www.jintiankansha.me/rss/GEYDAOJZGB6DGNZRGBQTEMJQGMYGGYZVMU2DIOLDGQYDAN3EGRRTMZRVHA4DCZLFGAYWEMLDMIYA====", # 保险一哥
    "http://www.jintiankansha.me/rss/GMZTQMZZHB6DMOJUGJSWIYJZMVTGEZDCMUZWGZJQMRQTMZRQGNQTEZDCHEZTGNRWMJQTIZRRGI2A====", # 中国银行保险报
    "http://www.jintiankansha.me/rss/GM2DANJSGF6GCNLBGA3DIZRVMZQTSMRYGI4DOOJUMY4WEY3CMZQTSYJZGRRGKNBUGI2GEZDCMVRQ====", # 保观
    "http://www.jintiankansha.me/rss/GMZTONZSGB6GMNZWG43TENZZMEYGMNRXMMYTKMBSGQ4DSMJZGU4DEODCGA4DQZJYGM4DAYRYMMYQ====", # 保契
    "http://www.jintiankansha.me/rss/GM2DCOBTGJ6DKY3DMFTGEMRYMFQWENRUGAZTEYRTMYZGENDFGA2DQYJTGI3DOY3CGNQTMOJYMM4A====", # 精算视觉
    "http://www.jintiankansha.me/rss/GMZTMNBRHB6DOMDBGFRTCZDCGRQTOMTDMY3TMMBSME2DQMRWGUZWMNLBMUYDCZRRG44DINZXHE3Q====", # 中保学
    "http://www.jintiankansha.me/rss/GMZTONZSGF6DMZDFGU3DIMLCGIYWINRVGM3WCMZWGA3TSZBQGRRTCODCMNRWCMDGMI3DAZRUMFRQ====", # 说点保
    "http://www.jintiankansha.me/rss/GMZTONZSGN6DMOBSHFSDSMLEGQYTEZBRHBRTQNDCGMZWIMBTGYYDOMBWG44DQY3FHAZGMOJUMI4A====", # 今日保
    "http://www.jintiankansha.me/rss/GM2TSMRUGZ6DAMZSGY4DQNJRGZSDKMJUMQ4GEY3GG4ZGINDFMJRGGYRTMMYGGOBSHBTDGNDBMM4A====", # 国寿研究声
    "http://www.jintiankansha.me/rss/GEYDANZXHF6DCMDBMZSDGMDDMRTDMZRYGIYDANTEGA3TMYRTGAYGINLEGYZDIMLBGQYWEMBTHAZQ====", # 欣琦看金融
    "https://hbdd2025.github.io/my-hexun-rss/hexun_insurance_rss.xml", # 和讯保险
    "https://hbdd2025.github.io/nfra-rss-feed/nfra_rss.xml", # 总局官网
    "http://www.jintiankansha.me/rss/GEYTGOBYPQZDKNDGMQZDQMRWGA2TQZRRGU2DGMTEG4ZWMMRXME3TSODBGFSTENRQGBT", # 中金点睛
    "http://www.jintiankansha.me/rss/GMZTONZRHF6DQZJSMY4GIMJQG5TGMZBWHA4WIMJWGUYTKNBYG5QTKYJVMQYGINJYGAYDEODEGFTA====", # 保煎烩
]

FEED_SOURCES = {
    ORIGINAL_FEEDS[0]: "总局公众号",
    ORIGINAL_FEEDS[1]: "慧保天下",
    ORIGINAL_FEEDS[2]: "十三精",
    ORIGINAL_FEEDS[3]: "蓝鲸",
    ORIGINAL_FEEDS[4]: "中保新知",
    ORIGINAL_FEEDS[5]: "保险一哥",
    ORIGINAL_FEEDS[6]: "中国银行保险报",
    ORIGINAL_FEEDS[7]: "保观",
    ORIGINAL_FEEDS[8]: "保契",
    ORIGINAL_FEEDS[9]: "精算视觉",
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

# ---------------- 工具函数 ----------------
def log(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)

def get_user_agent():
    return {
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/115.0 Safari/537.36'
        )
    }

def clean_extracted_html(html_content: str) -> str:
    if not html_content or not isinstance(html_content, str):
        return ""
    try:
        try:
            soup = BeautifulSoup(html_content, 'lxml')
        except Exception:
            soup = BeautifulSoup(html_content, 'html.parser')

        # 移除噪音
        remove_selectors = [
            'script', 'style', 'nav', 'footer', 'header', 'aside',
            'form', 'iframe', 'button', '.sidebar', '#sidebar',
            '.related-posts', '.comments', '#comments', '.navigation',
            '.pagination', '.share-buttons', 'figure > figcaption',
            '.advertisement', '.ad', 'ins'
        ]
        for sel in remove_selectors:
            try:
                for el in (soup.select(sel) if sel.startswith(('.', '#')) else soup.find_all(sel)):
                    el.decompose()
            except Exception:
                pass

        # 去掉链接外壳，只保留文本（避免 OneNote 外链杂质）
        for a in soup.find_all('a'):
            a.unwrap()

        # 去除空段落
        for p in soup.find_all('p'):
            if not p.get_text(strip=True) and not p.find(['img']):
                p.decompose()

        return str(soup)
    except Exception as e:
        log(f"[HTML清理] 异常：{e}")
        return html_content

def get_full_content_from_link(url: str, source_name: str, feed_url: str, verify_ssl=True):
    """
    抓正文 + 尝试提取发布时间（部分站点）
    返回: (clean_html, datetime|None)
    """
    if not url or not url.startswith("http"):
        return "[错误：无效的文章链接]", None

    try:
        log(f"  >> 抓取原文: {url[:100]} ...")
        resp = requests.get(url, headers=get_user_agent(), timeout=REQUEST_TIMEOUT,
                            allow_redirects=True, verify=verify_ssl)
        resp.raise_for_status()
        final_url = resp.url
        resp.encoding = resp.apparent_encoding or 'utf-8'

        try:
            soup = BeautifulSoup(resp.text, 'lxml')
        except Exception:
            soup = BeautifulSoup(resp.text, 'html.parser')

        # ---- 特例提取时间（和讯 / 总局 / 微信镜像）----
        extracted_date = None
        try:
            if source_name == '和讯保险' or (source_name and source_name.startswith('和讯-')):
                s = soup.select_one('span.pr20') or (soup.select_one('div.tip.fl.gray').find('span') if soup.select_one('div.tip.fl.gray') else None)
                if s:
                    import re
                    m = re.search(r'(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})', s.get_text(strip=True))
                    if m:
                        extracted_date = datetime.strptime(m.group(1), '%Y-%m-%d %H:%M:%S')

            elif source_name and source_name.startswith('总局'):
                info = soup.select_one('div.pages-detail-info')
                if info:
                    for span in info.find_all('span', recursive=False):
                        t = span.get_text(strip=True)
                        if "发布日期：" in t:
                            d = t.split("发布日期：")[-1].strip()
                            import re
                            m = re.search(r'^(\d{4}-\d{2}-\d{2})', d) or re.search(r'^(\d{4}年\d{1,2}月\d{1,2}日)', d)
                            if m:
                                s = m.group(1)
                                if '-' in s:
                                    extracted_date = datetime.strptime(s, '%Y-%m-%d')
                                else:
                                    extracted_date = datetime.strptime(s, '%Y年%m月%d日')

            elif feed_url and 'jintiankansha.me' in feed_url:
                em = soup.select_one('em#publish_time')
                if em:
                    t = em.get_text(strip=True)
                    try:
                        extracted_date = datetime.strptime(t, '%Y-%m-%d %H:%M') if ' ' in t else datetime.strptime(t, '%Y-%m-%d')
                    except Exception:
                        pass
        except Exception:
            pass

        # ---- 提取正文区域 ----
        selectors = []
        if source_name and source_name.startswith('总局'):
            selectors += ['div.content', 'div#zoom', '.pages_content', 'div.view.TRS_UEDITOR.trs_paper_default.trs_word']
        elif source_name == '和讯保险' or (source_name and source_name.startswith('和讯-')):
            selectors += ['div.art_contextBox', 'div.art_context']
        elif feed_url and 'jintiankansha.me' in feed_url:
            selectors += ['div.rich_media_content', 'div#js_content', 'div.wx-content']

        selectors += [
            'article', '.article-content', '.entry-content', '.post-content', '.post-body',
            '#article-body', '#entry-content', 'div[itemprop="articleBody"]',
            'div.main-content', 'div.entry', 'div[class*=content]', 'div[id*=content]', 'main'
        ]

        article_body, used_sel = None, None
        for sel in selectors:
            try:
                target = soup.select_one(sel)
                if target and (len(target.get_text(strip=True)) > 60 or target.find('img')):
                    article_body = target
                    used_sel = sel
                    break
            except Exception:
                continue

        if not article_body:
            return None, extracted_date

        # 图片绝对化
        fixed = 0
        for img in article_body.find_all('img'):
            try:
                src = img.get('src') or img.get('data-src')
                if src and not src.startswith(('http://', 'https://', 'data:')):
                    img['src'] = urljoin(final_url, src)
                    fixed += 1
            except Exception:
                pass
        if fixed:
            log(f"  >> 图片链接修复 {fixed} 个（选择器: {used_sel}）")

        cleaned = clean_extracted_html(str(article_body))
        return cleaned, extracted_date

    except requests.exceptions.Timeout:
        return "[错误：抓取超时]", None
    except Exception as e:
        return f"[错误：抓取或解析失败 - {e}]", None

def load_processed_items(filename=PROCESSED_ITEMS_FILE):
    items = set()
    if os.path.exists(filename):
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                for line in f:
                    s = line.strip()
                    if s:
                        items.add(s)
        except Exception as e:
            log(f"[状态] 读取 {filename} 失败：{e}")
    else:
        log(f"[状态] 未找到 {filename}，将创建新文件。")
    return items

def save_processed_items(new_ids, filename=PROCESSED_ITEMS_FILE):
    if not new_ids:
        return
    try:
        with open(filename, 'a', encoding='utf-8') as f:
            for x in new_ids:
                if x:
                    f.write(x + "\n")
        log(f"[状态] 已追加保存 {len(new_ids)} 条处理记录到 {filename}")
    except Exception as e:
        log(f"[状态] 保存 {filename} 失败：{e}")

def monthly_section_name(base: str, dt: datetime) -> str:
    if not base:
        return ""
    return f"{base} · {dt.strftime('%Y-%m')}"

# ---------------- OneNote 操作 ----------------
class OneNoteSync:
    def __init__(self):
        # Token Cache
        try:
            persistence = FilePersistence(TOKEN_CACHE_FILE)
            self.token_cache = PersistedTokenCache(persistence)
            log(f"[缓存] 使用文件持久化: {TOKEN_CACHE_FILE}")
        except Exception as e:
            log(f"[缓存] FilePersistence 初始化失败：{e}，改用内存缓存。")
            self.token_cache = msal.SerializableTokenCache()

        self.app = PublicClientApplication(
            client_id=CLIENT_ID,
            authority=AUTHORITY,
            token_cache=self.token_cache
        )

    def get_token(self):
        # 优先静默
        accounts = self.app.get_accounts()
        token_result = None
        if accounts:
            log(f"[认证] 找到缓存账户：{accounts[0].get('username', '未知')}")
            token_result = self.app.acquire_token_silent(SCOPES, account=accounts[0])

        if not token_result:
            # CI 环境走设备码，本地走交互
            if os.getenv('CI') == 'true':
                log("[认证] 缓存未命中 -> 设备码流程")
                flow = self.app.initiate_device_flow(scopes=SCOPES)
                if "user_code" not in flow:
                    log(f"[认证] 设备码启动失败：{flow}")
                    return None
                print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                print("!!! 需要人工操作：请在浏览器中完成设备登录 !!!")
                print(flow["message"])
                print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                sys.stdout.flush()
                token_result = self.app.acquire_token_by_device_flow(flow)
            else:
                log("[认证] 本地交互式登录")
                token_result = self.app.acquire_token_interactive(scopes=SCOPES)

        if token_result and "access_token" in token_result:
            log("[认证] 成功获取访问令牌")
            return token_result["access_token"]

        log(f"[认证] 失败：{token_result.get('error_description', '未知错误') if token_result else '空结果'}")
        return None

    def _request_with_retry(self, method, url, headers=None, json=None, data=None, max_retry=4):
        """对 429/5xx 做指数回退重试"""
        backoff = 2.0
        for attempt in range(1, max_retry + 1):
            try:
                resp = requests.request(method, url, headers=headers, json=json, data=data,
                                        timeout=REQUEST_TIMEOUT)
                if resp.status_code in (429, 500, 502, 503, 504):
                    ra = resp.headers.get("Retry-After")
                    wait = float(ra) if ra else backoff
                    log(f"[API] {resp.status_code}，{wait:.1f}s 后重试（第 {attempt}/{max_retry} 次）")
                    time.sleep(wait)
                    backoff *= 1.6
                    continue
                resp.raise_for_status()
                return resp
            except requests.exceptions.HTTPError as e:
                # 507 等直接返回
                return resp
            except Exception as e:
                if attempt >= max_retry:
                    log(f"[API] 重试失败：{e}")
                    return None
                log(f"[API] 异常 {e} ，{backoff:.1f}s 后重试（第 {attempt}/{max_retry} 次）")
                time.sleep(backoff)
                backoff *= 1.6
        return None

    def create_page(self, title: str, content_html: str, target_section_name: str):
        """
        优先写入指定分区名；不指定则写默认位置。
        遇 507（分区满）会尝试追加后缀重试（最多 2 次）。
        """
        safe_title = html.escape(title)
        created_ts = time.strftime('%Y-%m-%dT%H:%M:%S.000Z', time.gmtime())
        body = f'<!DOCTYPE html><html lang="zh-CN"><head><title>{safe_title}</title>' \
               f'<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />' \
               f'<meta name="created" content="{created_ts}" /></head><body>{content_html}</body></html>'

        tries = []
        if target_section_name:
            tries.append(target_section_name)
        else:
            tries.append("")  # 默认

        # 如果指定分区，遇 507，自动尝试“同月 · 溢出N”以分摊
        if target_section_name:
            tries.append(target_section_name + " · 溢出1")
            tries.append(target_section_name + " · 溢出2")

        for idx, sec in enumerate(tries, 1):
            if sec:
                url = f"https://graph.microsoft.com/v1.0/me/onenote/pages?sectionName={quote(sec, safe='')}"
                log(f"[OneNote] 目标分区：{sec}（尝试 {idx}/{len(tries)}）")
            else:
                url = "https://graph.microsoft.com/v1.0/me/onenote/pages"
                log(f"[OneNote] 使用默认位置（尝试 {idx}/{len(tries)}）")

            token = self.get_token()
            if not token:
                return False

            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/xhtml+xml"
            }
            resp = self._request_with_retry("POST", url, headers=headers, data=body)
            if resp is None:
                continue

            if resp.status_code == 201:
                return True

            # 507 分区满，进入下一轮分摊
            if resp.status_code == 507:
                log("[OneNote] 目标分区容量已满（507），尝试溢出分区...")
                continue

            # 其他错误，打印并放弃该次
            try:
                log(f"[OneNote] 创建失败：{resp.status_code} {resp.text[:300]}")
            except Exception:
                log(f"[OneNote] 创建失败：{resp.status_code}")
        return False


# ---------------- RSS 抓取 ----------------
def fetch_rss_feeds():
    log("[RSS] 开始抓取所有源...")
    all_entries = []
    for feed_url in ORIGINAL_FEEDS:
        source = FEED_SOURCES.get(feed_url, feed_url)
        log(f"  源：{source}")

        try:
            fp = feedparser.parse(feed_url, agent=get_user_agent()['User-Agent'],
                                  request_headers=get_user_agent())
        except Exception as e:
            log(f"  解析失败：{e}")
            continue

        if getattr(fp, "bozo", 0):
            log(f"  警告：解析器提示异常：{getattr(fp, 'bozo_exception', '')}")

        if not fp.entries:
            log("  警告：无条目")
            continue

        for entry in fp.entries:
            entry_id = entry.get('id', entry.get('link')) or f"{entry.get('title','无标题')}_{entry.get('published', entry.get('updated', time.time()))}"
            title = (entry.get('title') or '无标题').strip()
            link = (entry.get('link') or '').strip()

            pub_dt = None
            t = entry.get('published_parsed') or entry.get('updated_parsed')
            if t:
                try:
                    pub_dt = datetime.fromtimestamp(time.mktime(t))
                except Exception:
                    pub_dt = None

            summary = ""
            if 'content' in entry and entry.content:
                html_content = next((c.value for c in entry.content if getattr(c, 'type', '') and 'html' in c.type.lower()), None)
                summary = html_content or next((f"<p>{html.escape(c.value)}</p>" for c in entry.content if getattr(c, 'type', '') and 'text' in c.type.lower()), "")
            if not summary and entry.get('summary'):
                summary = entry.summary
            if not summary and entry.get('description'):
                summary = entry.description

            all_entries.append({
                'id': entry_id,
                'title': title,
                'link': link,
                'published_time_rss': pub_dt,
                'content_summary': summary,
                'source_name': FEED_SOURCES.get(feed_url, source),
                'feed_url': feed_url
            })

    # 按发布时间降序
    all_entries.sort(key=lambda x: x['published_time_rss'] or datetime.min, reverse=True)
    log(f"[RSS] 抓取完成，共 {len(all_entries)} 条。")
    return all_entries

# ---------------- 主流程 ----------------
def main():
    log(f"=== RSS to OneNote Sync 开始于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

    if not CLIENT_ID:
        sys.exit("错误：AZURE_CLIENT_ID 未设置。请到 GitHub Secrets 配置。")

    processed = load_processed_items()
    to_save_ids = set()

    entries = fetch_rss_feeds()
    new_entries = [e for e in entries if e['id'] not in processed]
    log(f"[筛选] 新条目 {len(new_entries)} 条。")

    if not new_entries:
        log("没有需要同步的新条目。")
        return

    # 本次最多处理
    batch = new_entries[:MAX_ITEMS_PER_RUN]
    log(f"[计划] 本次处理 {len(batch)} 条（上限 {MAX_ITEMS_PER_RUN}）。")

    # 分区名：固定 or 按月
    if ONENOTE_SECTION_BASE:
        if SECTION_MONTHLY:
            section_name = monthly_section_name(ONENOTE_SECTION_BASE, datetime.now())
        else:
            section_name = ONENOTE_SECTION_BASE
    else:
        section_name = ""  # 未设置则写默认位置
    if section_name:
        log(f"[分区] 目标分区：{section_name}（SECTION_MONTHLY={'on' if SECTION_MONTHLY else 'off'}）")
    else:
        log("[分区] 未设置 ONENOTE_SECTION_NAME，写默认位置。")

    onenote = OneNoteSync()

    success, fail = 0, 0
    for idx, e in enumerate(batch, 1):
        t0 = time.time()
        log(f"\n--- [{idx}/{len(batch)}] {e['title'][:70]} ---")
        log(f"  链接：{e['link']}")

        # 抓正文
        body_html, ext_time = get_full_content_from_link(e['link'], e['source_name'], e['feed_url'], verify_ssl=True)
        final_time = ext_time or e.get('published_time_rss') or datetime.now()
        timestr = final_time.strftime('%Y-%m-%d %H:%M:%S')

        # 标题加日期前缀：YYMMDD-
        date_prefix = final_time.strftime('%y%m%d')
        sep = "-" if e['title'] and e['title'][0].isdigit() else ""
        formatted_title = f"{date_prefix}{sep}{e['title']}"[:200]

        # 正文兜底
        if not body_html or (isinstance(body_html, str) and body_html.startswith("[错误：")):
            log("  >> 原文正文抓取失败，回退使用摘要。")
            body_html = clean_extracted_html(e['content_summary']) or "<p><i>[无法自动提取正文，请访问原文链接]</i></p>"

        onenote_content = f"""
            <h1>{html.escape(formatted_title)}</h1>
            <p style="font-size:9pt; color:gray;">
                发布时间: {timestr} | 来源: {html.escape(e['source_name'])}
                | <a href="{html.escape(e['link'])}">原文链接</a>
            </p>
            <hr/>
            <div>{body_html}</div>
        """

        ok = onenote.create_page(formatted_title, onenote_content, section_name)
        if ok:
            success += 1
            to_save_ids.add(e['id'])
            log(f"  >> 成功（耗时 {time.time()-t0:.1f}s）")
        else:
            fail += 1
            log(f"  !! 失败（耗时 {time.time()-t0:.1f}s）")

        time.sleep(REQUEST_DELAY)

    log(f"\n[统计] 成功 {success} 条，失败 {fail} 条。")
    if to_save_ids:
        save_processed_items(to_save_ids)

    log(f"=== 结束于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

if __name__ == "__main__":
    main()
