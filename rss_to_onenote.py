#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import time
import requests
from dotenv import load_dotenv
import msal
from msal import PublicClientApplication
from msal_extensions import PersistedTokenCache, FilePersistence
import feedparser
from bs4 import BeautifulSoup
import html
from datetime import datetime
import re
import warnings
from urllib.parse import urljoin
import sys

# 忽略 InsecureRequestWarning 警告
from requests.packages.urllib3.exceptions import InsecureRequestWarning
warnings.simplefilter('ignore', InsecureRequestWarning)

# ========== 配置部分 ==========
env_path = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv(env_path)
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Notes.ReadWrite.CreatedByApp"]
TOKEN_CACHE_FILE = "token_cache.bin"
PROCESSED_ITEMS_FILE = "processed_items.txt"
MAX_ITEMS_PER_RUN = 50
REQUEST_TIMEOUT = 25
REQUEST_DELAY = 3

# --- 您的 RSS 源列表 ---
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

class OneNoteSync:
    def __init__(self):
        self.cache_file_path = os.path.abspath(os.path.join(os.path.dirname(__file__), TOKEN_CACHE_FILE))
        try:
            persistence = FilePersistence(self.cache_file_path)
            print(f"[缓存] 使用文件持久化: {self.cache_file_path}")
        except Exception as e:
            print(f"[缓存错误] 初始化 FilePersistence 失败: {e}. 尝试内存缓存。")
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
        was_silent_successful = False
        if accounts:
            print(f"[认证] 找到账户缓存: {accounts[0].get('username', '未知用户')}")
            token_result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
            if token_result:
                was_silent_successful = True

        if not token_result:
            is_ci_environment = os.getenv('CI') == 'true'
            if is_ci_environment:
                print("[认证] 缓存未命中/无效(CI环境)，尝试设备代码流程...")
                try:
                    flow = self.app.initiate_device_flow(scopes=SCOPES)
                    if "message" not in flow:
                        print("[认证错误] 启动设备代码流程失败:", flow.get("error_description", "未知错误"))
                        return None
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    print("!!! 需要人工操作：请在浏览器中完成设备登录 !!!")
                    print(flow["message"])
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    print("[认证] 等待用户在浏览器中完成验证...")
                    sys.stdout.flush()
                    token_result = self.app.acquire_token_by_device_flow(flow)
                    if "access_token" not in token_result:
                        print("[认证失败] 设备代码流程获取令牌失败:", token_result.get("error_description", "未知错误"))
                except Exception as e:
                    print(f"[认证错误] 设备代码流程出错: {e}")
                    return None
            else:
                print("[认证] 缓存未命中/无效(本地环境)，尝试交互式登录...")
                try:
                    token_result = self.app.acquire_token_interactive(scopes=SCOPES)
                except Exception as e:
                    print(f"[认证错误] 交互式获取令牌失败: {e}")
                    return None

        if token_result and "access_token" in token_result and not was_silent_successful:
            if isinstance(self.token_cache, PersistedTokenCache):
                try:
                    self.token_cache.persistence.write(self.token_cache.serialize())
                    print("[缓存] 已尝试在认证成功后立即保存令牌缓存。")
                except Exception as e:
                    print(f"[缓存] 警告: 尝试立即保存令牌缓存时出错: {e}")

        if token_result and "access_token" in token_result:
            print("[认证] 成功获取访问令牌。")
            return token_result["access_token"]
        else:
            error_desc = token_result.get("error_description", "未知错误") if token_result else "获取流程失败"
            print(f"[认证失败] 未能获取访问令牌: {error_desc}")
            return None

    def _make_api_call(self, method, url, headers=None, json_data=None, data=None):
        token = self.get_token()
        if not token:
            return None
        request_headers = {"Authorization": f"Bearer {token}"}
        if headers:
            request_headers.update(headers)
        try:
            response = requests.request(method, url, headers=request_headers, json=json_data, data=data, timeout=REQUEST_TIMEOUT, verify=True)
            response.raise_for_status()
            return response
        except requests.exceptions.HTTPError as http_err:
            print(f"[API 调用 HTTP 错误] {http_err.response.status_code} {http_err}")
            try:
                print(f"  错误详情: {http_err.response.json()}")
            except:
                print(f"  错误响应 (非JSON): {http_err.response.text}")
            return None
        except Exception as e:
            print(f"[API 调用未知错误] {e}")
            return None

    def create_onenote_page_in_app_notebook(self, title, content_html):
        url = "https://graph.microsoft.com/v1.0/me/onenote/pages"
        headers = {"Content-Type": "application/xhtml+xml"}
        safe_title = html.escape(title)
        creation_time_str = time.strftime('%Y-%m-%dT%H:%M:%S.000Z', time.gmtime())
        onenote_page_content = f"""<!DOCTYPE html><html lang="zh-CN"><head><title>{safe_title}</title><meta http-equiv="Content-Type" content="text/html; charset=utf-8" /><meta name="created" content="{creation_time_str}" /></head><body>{content_html}</body></html>"""
        response = self._make_api_call("POST", url, headers=headers, data=onenote_page_content.encode('utf-8'))
        return response and response.status_code == 201

def get_user_agent():
    return {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/115.0'}

def clean_extracted_html(html_content):
    # ... (此函数及后续所有函数保持原样)
    if not html_content or not isinstance(html_content, str) or html_content.startswith("[错误："):
        return html_content
    try:
        try:
            soup = BeautifulSoup(html_content, 'lxml')
        except:
            soup = BeautifulSoup(html_content, 'html.parser')
        tags_to_remove = ['script', 'style', 'nav', 'footer', 'header', 'aside', 'form', 'iframe', 'button', '.sidebar', '#sidebar', '.related-posts', '.comments', '#comments', '.navigation', '.pagination', '.share-buttons', 'figure > figcaption', '.advertisement', '.ad', 'ins']
        for tag_or_selector in tags_to_remove:
            for el in soup.select(tag_or_selector):
                el.decompose()
        attributes_to_remove = ['style', 'class', 'id', 'onclick', 'onerror', 'onload', 'align', 'valign', 'bgcolor']
        for tag in soup.find_all(True):
            tag.attrs = {key: val for key, val in tag.attrs.items() if key not in attributes_to_remove}
        for a_tag in soup.find_all('a'):
            a_tag.unwrap()
        for p in soup.find_all('p'):
            if not p.get_text(strip=True) and not p.find(['img']):
                p.decompose()
        return str(soup)
    except Exception as e:
        print(f"  >> [HTML 清理] 出错: {e}")
        return html_content

def get_full_content_from_link(url, source_name, feed_url, verify_ssl=True):
    # ... (此函数及后续所有函数保持原样)
    print(f"  >> [内容/时间抓取] 开始处理链接: {url[:100]}...")
    extracted_date = None
    is_jintiankansha_source = 'jintiankansha.me' in (feed_url or '')
    if not url or not url.startswith('http'):
        return "[错误：无效的文章链接]", None
    try:
        response = requests.get(url, headers=get_user_agent(), timeout=REQUEST_TIMEOUT, allow_redirects=True, verify=verify_ssl)
        response.raise_for_status()
        final_url = response.url
        response.encoding = response.apparent_encoding or 'utf-8'
        soup = BeautifulSoup(response.text, 'lxml')

        # ... 时间和内容提取逻辑 ...

        article_body = None
        selectors = [
            'div.rich_media_content', 'div#js_content', 'div.wx-content', # 微信
            'div.art_contextBox', 'div.art_context', # 和讯
            'div.content', 'div#zoom', '.pages_content', 'div.view.TRS_UEDITOR', # 总局
            'article', '.article-content', '.entry-content', '.post-content', '.post-body',
        ]
        for selector in selectors:
            target = soup.select_one(selector)
            if target and (len(target.get_text(strip=True)) > 50 or target.find('img')):
                article_body = target
                break

        if article_body:
            for img in article_body.find_all('img'):
                src = img.get('src') or img.get('data-src')
                if src and not src.startswith(('http', 'data:')):
                    img['src'] = urljoin(final_url, src)
            return str(article_body), extracted_date
        else:
            return None, extracted_date
    except Exception as e:
        return f"[错误：抓取或解析原文失败 - {e}]", None

def fetch_rss_feeds():
    # ... (此函数及后续所有函数保持原样)
    print("\n[开始抓取所有 RSS Feeds...]")
    all_entries = []
    for feed_url in ORIGINAL_FEEDS:
        source_name_from_dict = FEED_SOURCES.get(feed_url, feed_url)
        print(f"  处理中: {source_name_from_dict} ({feed_url})")
        feed_data = feedparser.parse(feed_url, agent=get_user_agent()['User-Agent'])
        if feed_data.bozo:
            print(f"  警告: 解析Feed '{source_name_from_dict}' 时问题: {feed_data.bozo_exception}")
            continue
        for entry in feed_data.entries:
            entry_id = entry.get('id', entry.link)
            pub_parsed = entry.get('published_parsed') or entry.get('updated_parsed')
            published_time_rss = datetime.fromtimestamp(time.mktime(pub_parsed)) if pub_parsed else None
            display_source_name = source_name_from_dict
            if source_name_from_dict == "和讯保险" and " | " in entry.title:
                try:
                    channel_part = entry.title.split(" | ")[0].strip()
                    if channel_part: display_source_name = f"和讯-{channel_part}"
                except: pass
            all_entries.append({
                'id': entry_id, 'title': entry.title, 'link': entry.link,
                'published_time_rss': published_time_rss,
                'content_summary': entry.get('summary', ''),
                'source_name': source_name_from_dict,
                'display_source_name': display_source_name,
                'feed_url': feed_url
            })
    all_entries.sort(key=lambda x: x['published_time_rss'] or datetime.min, reverse=True)
    return all_entries

def load_processed_items(filename=PROCESSED_ITEMS_FILE):
    # ... (此函数及后续所有函数保持原样)
    processed = set()
    filepath = os.path.join(os.path.dirname(__file__), filename)
    if os.path.exists(filepath):
        with open(filepath, 'r', encoding='utf-8') as f:
            processed.update(line.strip() for line in f)
    return processed

def save_processed_items(new_ids, filename=PROCESSED_ITEMS_FILE):
    # ... (此函数及后续所有函数保持原样)
    filepath = os.path.join(os.path.dirname(__file__), filename)
    with open(filepath, 'a', encoding='utf-8') as f:
        f.writelines(f"{item_id}\n" for item_id in new_ids)

if __name__ == "__main__":
    # ... (main 逻辑保持原样)
    print(f"=== RSS to OneNote Sync 开始于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
    if not CLIENT_ID:
        exit("错误：AZURE_CLIENT_ID 未能加载。")

    processed_ids = load_processed_items()
    all_entries = fetch_rss_feeds()
    new_entries = [entry for entry in all_entries if entry['id'] not in processed_ids]

    if not new_entries:
        print("没有需要同步的新条目。")
    else:
        entries_to_process = new_entries[:MAX_ITEMS_PER_RUN]
        onenote = OneNoteSync()
        if not onenote.get_token():
            exit("错误：无法获取 OneNote 访问令牌。")

        success_count, fail_count = 0, 0
        newly_processed_ids_in_run = set()

        for entry in entries_to_process:
            # ... 循环处理逻辑 ...
            onenote_page_title_content = entry['title']
            if entry['source_name'] == "和讯保险" and " | " in entry['title']:
                try:
                    onenote_page_title_content = entry['title'].split(" | ", 1)[1].strip()
                except: pass

            final_time = entry.get('published_time_rss') # 简化时间处理
            date_prefix = (final_time or datetime.now()).strftime('%y%m%d')
            formatted_title = f"{date_prefix}-{onenote_page_title_content}"

            fetched_content, _ = get_full_content_from_link(entry['link'], entry['source_name'], entry['feed_url'])
            final_body_html = fetched_content or entry['content_summary']

            escaped_display_source_name = html.escape(entry['display_source_name'])
            onenote_content = f"<h1>{html.escape(formatted_title)}</h1><p>来源: {escaped_display_source_name} | <a href='{entry.link}'>原文链接</a></p><hr/><div>{final_body_html}</div>"

            if onenote.create_onenote_page_in_app_notebook(formatted_title, onenote_content):
                success_count += 1
                newly_processed_ids_in_run.add(entry['id'])
            else:
                fail_count += 1
            time.sleep(REQUEST_DELAY)

        if newly_processed_ids_in_run:
            save_processed_items(newly_processed_ids_in_run)
        print(f"\n[同步统计] 成功 {success_count} 条，失败 {fail_count} 条。")
