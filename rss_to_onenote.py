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
# 引入 urlencode 的 quote 函数
from urllib.parse import urljoin, quote
import sys

# 忽略 InsecureRequestWarning 警告
from requests.packages.urllib3.exceptions import InsecureRequestWarning
warnings.simplefilter('ignore', InsecureRequestWarning)

# ========== 配置部分 ==========
env_path = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv(env_path)
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
ONENOTE_SECTION_NAME = os.getenv("ONENOTE_SECTION_NAME")
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Notes.ReadWrite.CreatedByApp"]
TOKEN_CACHE_FILE = "token_cache.bin"
PROCESSED_ITEMS_FILE = "processed_items.txt"
MAX_ITEMS_PER_RUN = 50
REQUEST_TIMEOUT = 25
REQUEST_DELAY = 3

# (RSS源列表和映射保持不变, 此处省略)
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
        if accounts:
            print(f"[认证] 找到账户缓存: {accounts[0].get('username', '未知用户')}")
            token_result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
        if not token_result:
            print("[认证] 缓存未命中/无效，需要认证...")
            try:
                if os.getenv('CI') == 'true':
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
                    token_result = self.app.acquire_token_interactive(scopes=SCOPES)
            except Exception as e:
                print(f"[认证错误] 认证流程中出现异常: {e}")
                return None
        if token_result and "access_token" in token_result:
            print("[认证] 成功获取访问令牌。")
            return token_result["access_token"]
        else:
            print(f"[认证失败] 未能获取访问令牌: {token_result.get('error_description', '未知错误') if token_result else '流程返回空'}")
            return None
    def _make_api_call(self, method, url, headers=None, json_data=None, data=None):
        token = self.get_token()
        if not token:
            print("[API] 因无法获取令牌而跳过API调用。")
            return None
        request_headers = {"Authorization": f"Bearer {token}"}
        if headers:
            request_headers.update(headers)
        try:
            response = requests.request(method, url, headers=request_headers, json=json_data, data=data, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
            return response
        except requests.exceptions.HTTPError as e:
            print(f"[API 调用 HTTP 错误] {e.response.status_code} {e.response.reason} for url {url}")
            try:
                print(f"  错误详情: {e.response.json()}")
            except:
                print(f"  错误响应 (非JSON): {e.response.text}")
            return None
        except Exception as e:
            print(f"[API 调用未知错误] {e}")
            return None
    def create_onenote_page_in_app_notebook(self, title, content_html):
        if ONENOTE_SECTION_NAME:
            section = ONENOTE_SECTION_NAME.strip()
            # 使用 quote 函数对分区名进行编码
            url = f"https://graph.microsoft.com/v1.0/me/onenote/pages?sectionName={quote(section, safe='')}"
            print(f"[OneNote] 将尝试保存到指定分区: {section}")
        else:
            url = "https://graph.microsoft.com/v1.0/me/onenote/pages"
            print("[OneNote] 未指定分区，将保存到默认位置。")

        headers = {"Content-Type": "application/xhtml+xml"}
        safe_title = html.escape(title)
        creation_time_str = time.strftime('%Y-%m-%dT%H:%M:%S.000Z', time.gmtime())
        onenote_page_content = f"""<!DOCTYPE html><html lang="zh-CN"><head><title>{safe_title}</title><meta http-equiv="Content-Type" content="text/html; charset=utf-8" /><meta name="created" content="{creation_time_str}" /></head><body>{content_html}</body></html>"""
        response = self._make_api_call("POST", url, headers=headers, data=onenote_page_content.encode('utf-8'))
        return response and response.status_code == 201
def get_user_agent():
    return {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
def clean_extracted_html(html_content):
    if not html_content or not isinstance(html_content, str): return ""
    try:
        soup = BeautifulSoup(html_content, 'lxml')
        tags_to_remove = ['script', 'style', 'nav', 'footer', 'header', 'aside', 'form', 'iframe', 'button', '.sidebar', '#sidebar', '.related-posts', '.comments', '#comments']
        for tag_selector in tags_to_remove:
            for element in soup.select(tag_selector):
                element.decompose()
        for a_tag in soup.find_all('a'):
            a_tag.unwrap()
        return str(soup)
    except Exception as e:
        print(f"  >> HTML 清理出错: {e}")
        return html_content
def get_full_content_from_link(url, feed_url):
    if not url or not url.startswith('http'): return None, None
    try:
        response = requests.get(url, headers=get_user_agent(), timeout=REQUEST_TIMEOUT, allow_redirects=True)
        response.raise_for_status()
        final_url = response.url
        response.encoding = response.apparent_encoding or 'utf-8'
        soup = BeautifulSoup(response.text, 'lxml')
        article_body = None
        selectors = [
            'div.rich_media_content', 'div#js_content',
            'div.art_contextBox', 'div.art_context',
            'div.content', 'div#zoom',
            'article', '.article-content', '.entry-content',
        ]
        for selector in selectors:
            target = soup.select_one(selector)
            if target and len(target.get_text(strip=True)) > 100:
                article_body = target
                break
        if not article_body: return None, None
        for img in article_body.find_all('img'):
            src = img.get('src') or img.get('data-src')
            if src and not src.startswith(('http', 'data:')):
                img['src'] = urljoin(final_url, src)
        return str(article_body), None
    except Exception:
        return None, None
def fetch_rss_feeds():
    all_entries = []
    for feed_url in ORIGINAL_FEEDS:
        source_name = FEED_SOURCES.get(feed_url, "未知来源")
        print(f"  正在处理: {source_name}")
        feed_data = feedparser.parse(feed_url, agent=get_user_agent()['User-Agent'])
        if feed_data.bozo: continue
        for entry in feed_data.entries:
            all_entries.append({
                'id': entry.get('id', entry.link),
                'title': entry.title,
                'link': entry.link,
                'published_time_rss': datetime.fromtimestamp(time.mktime(entry.published_parsed)) if hasattr(entry, 'published_parsed') else datetime.now(),
                'content_summary': entry.get('summary', ''),
                'source_name': source_name,
                'feed_url': feed_url
            })
    all_entries.sort(key=lambda x: x['published_time_rss'], reverse=True)
    return all_entries
def load_processed_items(filename=PROCESSED_ITEMS_FILE):
    filepath = os.path.join(os.path.dirname(__file__), filename)
    if os.path.exists(filepath):
        with open(filepath, 'r', encoding='utf-8') as f:
            return set(line.strip() for line in f)
    return set()
def save_processed_items(new_ids, filename=PROCESSED_ITEMS_FILE):
    filepath = os.path.join(os.path.dirname(__file__), filename)
    with open(filepath, 'a', encoding='utf-8') as f:
        f.writelines(f"{item_id}\n" for item_id in new_ids)
if __name__ == "__main__":
    print(f"=== RSS to OneNote Sync 开始于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
    if not CLIENT_ID:
        sys.exit("错误：AZURE_CLIENT_ID 未能加载。")
    processed_ids = load_processed_items()
    all_entries = fetch_rss_feeds()
    new_entries = [e for e in all_entries if e['id'] not in processed_ids]
    if not new_entries:
        print("没有需要同步的新条目。")
    else:
        entries_to_process = new_entries[:MAX_ITEMS_PER_RUN]
        print(f"发现 {len(new_entries)} 条新条目，本次将处理 {len(entries_to_process)} 条。")
        onenote = OneNoteSync()
        if not onenote.get_token():
            sys.exit("错误：无法获取 OneNote 访问令牌。")
        
        success_count, fail_count = 0, 0
        newly_processed_ids_in_run = set()
        for entry in entries_to_process:
            print(f"\n--- 正在处理: {entry['title'][:50]}...")
            
            date_prefix = entry['published_time_rss'].strftime('%y%m%d')
            formatted_title = f"{date_prefix}-{entry['title']}"
            
            full_content, _ = get_full_content_from_link(entry['link'], entry['feed_url'])
            final_body_html = clean_extracted_html(full_content or entry['content_summary'])
            
            onenote_content = f"""
                <h1>{html.escape(formatted_title)}</h1>
                <p style="font-size:9pt; color:gray;">
                    来源: {html.escape(entry['source_name'])} | <a href="{html.escape(entry['link'])}">原文链接</a>
                </p><hr/>
                <div>{final_body_html}</div>
            """
            
            if onenote.create_onenote_page_in_app_notebook(formatted_title, onenote_content):
                success_count += 1
                newly_processed_ids_in_run.add(entry['id'])
                print(f"  >> 成功同步到 OneNote。")
            else:
                fail_count += 1
                print(f"  >> 同步失败。")
            time.sleep(REQUEST_DELAY)
            
        if newly_processed_ids_in_run:
            save_processed_items(newly_processed_ids_in_run)
        print(f"\n[同步统计] 成功 {success_count} 条，失败 {fail_count} 条。")
