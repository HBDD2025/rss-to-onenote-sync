#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
RSS → OneNote 同步脚本（按月滚动分区 + 507自动换分区重试版）

✅ 功能要点
- 从预置 RSS 源抓取最新文章，尽量抓取原文正文，清洗后写入 OneNote。
- OneNote 分区按月滚动：<基础分区名> · YYYY-MM，例如 “RSS syn2 · 2025-08”。
- 保险重试：若分区写满（HTTP 507），自动改用 “(2)”、“(3)” 结尾的新分区再试。
- 支持在 CI 环境（GitHub Actions）里用设备码登录；本地则交互式登录。
- 令牌缓存：token_cache.bin；去重状态：processed_items.txt
  （你工作流里会把它们加密为 .enc 再提交）

⚙️ 必要环境变量（GitHub Actions 里传入）
- AZURE_CLIENT_ID（必填）
- ONENOTE_SECTION_NAME（建议填你建的新分区基础名，比如 “RSS syn2”，脚本会自动加上 “ · YYYY-MM”）
- CI=true（Actions 中已设置）

依赖：requests, python-dotenv, msal, msal-extensions, feedparser, beautifulsoup4, lxml
"""

import os
import sys
import time
import html
import re
import warnings
from datetime import datetime
from urllib.parse import urljoin, quote

import requests
import feedparser
from bs4 import BeautifulSoup
from dotenv import load_dotenv

import msal
from msal import PublicClientApplication
from msal_extensions import PersistedTokenCache, FilePersistence

# -----------------------------
# 环境 & 常量
# -----------------------------
# 忽略 InsecureRequestWarning 警告（某些源 https 证书链问题）
from requests.packages.urllib3.exceptions import InsecureRequestWarning
warnings.simplefilter('ignore', InsecureRequestWarning)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, '.env'))

CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
ONENOTE_SECTION_NAME = os.getenv("ONENOTE_SECTION_NAME")  # 基础分区名（脚本会拼上 " · YYYY-MM"）
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Notes.ReadWrite.CreatedByApp"]

TOKEN_CACHE_FILE = os.path.join(BASE_DIR, "token_cache.bin")
PROCESSED_ITEMS_FILE = os.path.join(BASE_DIR, "processed_items.txt")

MAX_ITEMS_PER_RUN = 50
REQUEST_TIMEOUT = 25
REQUEST_DELAY = 3  # 每条之间延时，别过于激进

# -----------------------------
# RSS 源
# -----------------------------
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
    "https://hbdd2025.github.io/my-hexun-rss/hexun_insurance_rss.xml",                               # 和讯保险
    "https://hbdd2025.github.io/nfra-rss-feed/nfra_rss.xml",                                        # 总局官网
    "http://www.jintiankansha.me/rss/GEYTGOBYPQZDKNDGMQZDQMRWGA2TQZRRGU2DGMTEG4ZWMMRXME3TSODBGFSTENRQGBT", # 中金点睛
    "http://www.jintiankansha.me/rss/GMZTONZRHF6DQZJSMY4GIMJQG5TGMZBWHA4WIMJWGUYTKNBYG5QTKYJVMQYGINJYGAYDEODEGFTA====", # 保煎烩
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

# -----------------------------
# 工具函数
# -----------------------------
def get_user_agent():
    return {
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/123.0 Safari/537.36'
        )
    }

def clean_extracted_html(html_content: str) -> str:
    """尽量清理文章HTML内容，去广告/脚本/空段落/外链包装等"""
    if not html_content or not isinstance(html_content, str):
        return ""
    try:
        try:
            soup = BeautifulSoup(html_content, 'lxml')
        except Exception:
            soup = BeautifulSoup(html_content, 'html.parser')

        # 去除无关元素
        selectors = [
            'script', 'style', 'nav', 'footer', 'header', 'aside', 'form',
            'iframe', 'button', '.sidebar', '#sidebar', '.related-posts',
            '.comments', '#comments', '.navigation', '.pagination',
            '.share-buttons', 'figure > figcaption', '.advertisement', '.ad',
            'ins'
        ]
        for sel in selectors:
            for el in soup.select(sel):
                el.decompose()

        # 去除多余属性，保留结构
        attrs_rm = {'style', 'class', 'id', 'onclick', 'onerror', 'onload', 'align', 'valign', 'bgcolor'}
        for tag in soup.find_all(True):
            # 只删不安全/样式类属性，保留 src/href 等关键属性（图片需要）
            for k in list(tag.attrs.keys()):
                if k in attrs_rm:
                    del tag.attrs[k]

        # 去掉外链包装
        for a in soup.find_all('a'):
            a.unwrap()

        # 去掉空段落
        for p in soup.find_all('p'):
            if not p.get_text(strip=True) and not p.find('img'):
                p.decompose()

        return str(soup)
    except Exception as e:
        print(f"  >> [HTML 清理] 出错: {e}")
        return html_content


def get_full_content_from_link(url: str, feed_url: str):
    """尽量抓取原文正文；失败则返回 (None, None)"""
    if not url or not url.startswith('http'):
        return None, None
    try:
        resp = requests.get(url, headers=get_user_agent(), timeout=REQUEST_TIMEOUT, allow_redirects=True)
        resp.raise_for_status()
        final_url = resp.url
        resp.encoding = resp.apparent_encoding or 'utf-8'
        try:
            soup = BeautifulSoup(resp.text, 'lxml')
        except Exception:
            soup = BeautifulSoup(resp.text, 'html.parser')

        # 针对常见站点的正文选择器（从最具体到通用）
        selectors = [
            # 微信 / 今日看啥
            'div.rich_media_content', 'div#js_content', 'div.wx-content',
            # 和讯
            'div.art_contextBox', 'div.art_context',
            # 总局、政务站常用
            'div.content', 'div#zoom', '.pages_content', 'div.view.TRS_UEDITOR',
            # 通用博客/资讯
            'article', '.article-content', '.entry-content', '.post-content', '.post-body',
            'main', 'div[itemprop="articleBody"]'
        ]
        body = None
        for sel in selectors:
            t = soup.select_one(sel)
            if t and (len(t.get_text(strip=True)) > 60 or t.find('img')):
                body = t
                break

        if not body:
            return None, None

        # 修复图片相对路径
        for img in body.find_all('img'):
            src = img.get('src') or img.get('data-src')
            if src:
                if not src.startswith(('http://', 'https://', 'data:')):
                    img['src'] = urljoin(final_url, src)
                elif 'data-src' in img.attrs and img.get('src') != img['data-src']:
                    ds = img['data-src']
                    img['src'] = urljoin(final_url, ds) if not ds.startswith(('http', 'data:')) else ds

        return str(body), None
    except Exception as e:
        print(f"  >> [内容抓取失败] {e}")
        return None, None


def load_processed_items(filename=PROCESSED_ITEMS_FILE):
    processed = set()
    if os.path.exists(filename):
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                for line in f:
                    v = line.strip()
                    if v:
                        processed.add(v)
            print(f"[状态] 已加载 {len(processed)} 个已处理条目ID")
        except Exception as e:
            print(f"[状态] 读取已处理条目失败：{e}")
    else:
        print("[状态] 未找到已处理条目文件，首次运行。")
    return processed


def save_processed_items(new_ids, filename=PROCESSED_ITEMS_FILE):
    if not new_ids:
        return
    try:
        with open(filename, 'a', encoding='utf-8') as f:
            for item_id in new_ids:
                if item_id:
                    f.write(item_id + "\n")
        print(f"[状态] 已追加写入 {len(new_ids)} 个新处理条目ID")
    except Exception as e:
        print(f"[状态] 保存已处理条目失败：{e}")


# -----------------------------
# OneNote 同步
# -----------------------------
class OneNoteSync:
    def __init__(self):
        # 持久化令牌缓存
        try:
            persistence = FilePersistence(TOKEN_CACHE_FILE)
            print(f"[缓存] 使用文件持久化: {TOKEN_CACHE_FILE}")
        except Exception as e:
            print(f"[缓存] FilePersistence 初始化失败: {e}，改用内存缓存（运行后不会落盘）。")
            persistence = None

        self.token_cache = PersistedTokenCache(persistence) if persistence else msal.SerializableTokenCache()

        self.app = PublicClientApplication(
            client_id=CLIENT_ID,
            authority=AUTHORITY,
            token_cache=self.token_cache
        )

    def get_token(self):
        token_result = None
        # 先静默刷新
        accounts = self.app.get_accounts()
        if accounts:
            print(f"[认证] 找到账户缓存：{accounts[0].get('username', '未知')}")
            token_result = self.app.acquire_token_silent(SCOPES, account=accounts[0])

        # 无缓存 or 失效，走交互/设备码
        if not token_result:
            print("[认证] 缓存未命中/无效，需要认证…")
            try:
                if os.getenv("CI") == "true":
                    flow = self.app.initiate_device_flow(scopes=SCOPES)
                    if "message" not in flow:
                        print(f"[认证] 启动设备码流程失败: {flow.get('error_description', '未知错误')}")
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
                print(f"[认证] 流程异常: {e}")
                return None

        if token_result and "access_token" in token_result:
            print("[认证] 成功获取访问令牌。")
            return token_result["access_token"]

        print(f"[认证] 未能获取访问令牌: {token_result.get('error_description', '未知错误') if token_result else '无返回'}")
        return None

    def _make_api_call(self, method, url, headers=None, json_data=None, data=None):
        token = self.get_token()
        if not token:
            print("[API] 因无令牌而取消调用。")
            return None

        req_headers = {"Authorization": f"Bearer {token}"}
        if headers:
            req_headers.update(headers)
        try:
            resp = requests.request(
                method, url, headers=req_headers, json=json_data, data=data,
                timeout=REQUEST_TIMEOUT
            )
            # 不立即 raise，让调用方能看状态码
            return resp
        except Exception as e:
            print(f"[API] 调用异常: {e}")
            return None

    def _build_section_name_candidates(self):
        """
        返回候选分区名（最多3个）：
        base · YYYY-MM
        base · YYYY-MM (2)
        base · YYYY-MM (3)
        """
        base = (ONENOTE_SECTION_NAME or "RSS syn2").strip()
        ym = datetime.utcnow().strftime("%Y-%m")  # 月滚动
        first = f"{base} · {ym}"
        return [first, f"{first} (2)", f"{first} (3)"]

    def create_onenote_page_with_fallback(self, title, content_html):
        """
        在按月分区写入页面；若遇 507（分区满）自动换下一个候选分区重试。
        成功返回 True，否则 False。
        """
        headers = {"Content-Type": "application/xhtml+xml"}
        safe_title = html.escape(title)
        creation_time_str = time.strftime('%Y-%m-%dT%H:%M:%S.000Z', time.gmtime())
        page_body = f"""<!DOCTYPE html>
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

        candidates = self._build_section_name_candidates()
        last_status = None

        for idx, section in enumerate(candidates, start=1):
            url = f"https://graph.microsoft.com/v1.0/me/onenote/pages?sectionName={quote(section, safe='')}"
            print(f"[OneNote] 目标分区（尝试 {idx}/{len(candidates)}）: {section}")

            resp = self._make_api_call("POST", url, headers=headers, data=page_body.encode('utf-8'))
            if resp is None:
                print("  >> [OneNote] 调用失败（无响应）。")
                last_status = None
                break

            last_status = resp.status_code
            if resp.status_code == 201:
                print("  >> [OneNote] 创建页面成功。")
                return True

            # 打印错误详情
            try:
                err_json = resp.json()
                print(f"  >> [OneNote] 调用失败 {resp.status_code}: {err_json}")
            except Exception:
                print(f"  >> [OneNote] 调用失败 {resp.status_code}: {resp.text[:500]}")

            # 507 Insufficient Storage -> 分区写满，换下一个候选
            if resp.status_code == 507:
                print("  >> [OneNote] 分区写满，尝试下一个候选分区…")
                continue

            # 其他错误直接放弃重试
            print("  >> [OneNote] 非容量错误，放弃本条。")
            break

        print(f"[OneNote] 创建页面最终失败（最后状态: {last_status}）。")
        return False


# -----------------------------
# RSS 抓取/筛选
# -----------------------------
def fetch_rss_feeds():
    print("\n[开始抓取 RSS 源] …")
    all_entries = []
    for feed_url in ORIGINAL_FEEDS:
        source_name = FEED_SOURCES.get(feed_url, "未知来源")
        print(f"  抓取：{source_name}")
        try:
            feed = feedparser.parse(feed_url, agent=get_user_agent()['User-Agent'])
            # feed.bozo 为 True 代表解析有问题，但有时仍能取到 entries，这里仅提示不直接丢弃
            if getattr(feed, 'bozo', False):
                print(f"  >> 警告：解析异常：{getattr(feed, 'bozo_exception', '未知')}")
            for e in feed.entries:
                entry_id = e.get('id') or e.get('link') or (e.get('title', '') + str(e.get('published', '')))
                title = e.get('title', '无标题').strip()
                link = (e.get('link') or '').strip()

                pub_parsed = e.get('published_parsed') or e.get('updated_parsed')
                if pub_parsed:
                    try:
                        # time module used:
                        import time as _t
                        published_dt = datetime.fromtimestamp(_t.mktime(pub_parsed))
                    except Exception:
                        published_dt = None
                else:
                    published_dt = None

                summary_html = ""
                if 'content' in e and e.content:
                    html_content = next((c.value for c in e.content if getattr(c, 'type', '') and 'html' in c.type.lower()), None)
                    summary_html = html_content or next((f"<p>{html.escape(c.value)}</p>" for c in e.content if getattr(c, 'type', '') and 'text' in c.type.lower()), "")
                if not summary_html:
                    summary_html = e.get('summary', e.get('description', ''))

                all_entries.append({
                    'id': entry_id,
                    'title': title,
                    'link': link,
                    'published_time_rss': published_dt,
                    'content_summary': summary_html,
                    'source_name': source_name,
                    'feed_url': feed_url
                })
        except Exception as ex:
            print(f"  >> 抓取失败：{ex}")

    print(f"[RSS 完成] 共抓到 {len(all_entries)} 条。")
    # 最新在前
    all_entries.sort(key=lambda x: x['published_time_rss'] or datetime.min, reverse=True)
    return all_entries


# -----------------------------
# 主流程
# -----------------------------
def main():
    print(f"=== RSS to OneNote Sync 开始于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

    if not CLIENT_ID:
        print("错误：AZURE_CLIENT_ID 未设置。")
        sys.exit(1)

    processed_ids = load_processed_items()
    all_entries = fetch_rss_feeds()

    new_entries = [e for e in all_entries if e['id'] and e['id'] not in processed_ids]
    print(f"[筛选] 新条目 {len(new_entries)} 条。")

    if not new_entries:
        print("没有需要同步的新条目。")
        return

    entries_to_process = new_entries[:MAX_ITEMS_PER_RUN]
    print(f"[计划] 本次处理 {len(entries_to_process)} 条（上限 {MAX_ITEMS_PER_RUN}）。")

    onenote = OneNoteSync()
    token = onenote.get_token()
    if not token:
        print("错误：无法获取 OneNote 访问令牌，终止。")
        sys.exit(1)

    success_count = 0
    fail_count = 0
    newly_processed_ids = set()

    for idx, entry in enumerate(entries_to_process, start=1):
        print(f"\n--- [{idx}/{len(entries_to_process)}] {entry['title'][:60]} ---")
        print(f"  链接: {entry['link']}")
        start_t = time.time()

        # 正文抓取
        body_html, _ = get_full_content_from_link(entry['link'], entry['feed_url'])
        if not body_html:
            body_html = entry['content_summary'] or "<p><i>[无法自动提取正文，请访问原文链接查看]</i></p>"

        body_html = clean_extracted_html(body_html)

        # 标题加上日期前缀（若RSS有时间）
        dt = entry.get('published_time_rss') or datetime.utcnow()
        prefix = dt.strftime('%y%m%d')
        # 避免标题首字符就是数字时导致阅读混淆，加个分隔符
        sep = "-" if entry['title'] and entry['title'][0].isdigit() else ""
        final_title = f"{prefix}{sep}{entry['title']}"
        final_title = final_title[:200]  # OneNote title 最长保护

        # 页面内容骨架
        info_line = (
            f'<p style="font-size:9pt;color:gray;">'
            f'来源: {html.escape(entry["source_name"])} | '
            f'<a href="{html.escape(entry["link"])}">原文链接</a>'
            f'</p><hr/>'
        )
        content = f"<h1>{html.escape(final_title)}</h1>\n{info_line}\n<div>{body_html}</div>"

        ok = onenote.create_onenote_page_with_fallback(final_title, content)
        if ok:
            success_count += 1
            newly_processed_ids.add(entry['id'])
            print(f"  >> 成功（耗时 {time.time() - start_t:.2f}s）")
        else:
            fail_count += 1
            print("  >> 失败")

        time.sleep(REQUEST_DELAY)

    if newly_processed_ids:
        save_processed_items(newly_processed_ids)

    print(f"\n[同步统计] 成功 {success_count} 条，失败 {fail_count} 条。")
    print(f"=== 结束于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")


if __name__ == "__main__":
    main()
