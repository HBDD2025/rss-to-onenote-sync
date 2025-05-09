#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import time
import requests
from dotenv import load_dotenv
import msal
from msal import PublicClientApplication
# --- 导入 msal_extensions ---
from msal_extensions import PersistedTokenCache, FilePersistence
# ---
import feedparser
from bs4 import BeautifulSoup
import html # 用于HTML转义
from datetime import datetime, timezone # 导入 timezone
import re # 用于检查标题开头的数字
import warnings # 用于处理 SSL 警告
from urllib.parse import urljoin
import sys # 用于退出脚本

# 忽略 InsecureRequestWarning 警告
from requests.packages.urllib3.exceptions import InsecureRequestWarning
warnings.simplefilter('ignore', InsecureRequestWarning)

# ========== 配置部分 ==========
# 加载环境变量
env_path = os.path.join(os.path.dirname(__file__), '.env')
load_dotenv(env_path)
# Azure AD 配置
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Notes.ReadWrite.CreatedByApp"] # 权限范围
TOKEN_CACHE_FILE = "token_cache.bin" # 缓存文件名
PROCESSED_ITEMS_FILE = "processed_items.txt"
# --- 设置为 50 ---
MAX_ITEMS_PER_RUN = 50
REQUEST_TIMEOUT = 25
REQUEST_DELAY = 3 # 处理每个条目后的基本延迟（秒）

# --- RSS 源列表和来源映射 ---
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
    "http://www.jintiankansha.me/rss/GM2TSMRUGZ6DAMZSGY4DQNJRGZSDKMJUMQ4GEY3GG4ZGINDFMJRGGYRTMMYGGOBSHBTDGNDBMM4A====", # 国寿研究声 (名称已修改)
    "http://www.jintiankansha.me/rss/GEYDANZXHF6DCMDBMZSDGMDDMRTDMZRYGIYDANTEGA3TMYRTGAYGINLEGYZDIMLBGQYWEMBTHAZQ====", # 欣琦看金融
    "https://hbdd2025.github.io/my-hexun-rss/hexun_insurance_rss.xml", # 和讯保险 (原已存在)
    "https://hbdd2025.github.io/nfra-rss-feed/nfra_rss.xml", # 总局官网
    "http://www.jintiankansha.me/rss/GEYTGOBYPQZDKNDGMQZDQMRWGA2TQZRRGU2DGMTEG4ZWMMRXME3TSODBGFSTENRQGBT", # <--- 新添加: 中金点睛
    "http://www.jintiankansha.me/rss/GMZTONZRHF6DQZJSMY4GIMJQG5TGMZBWHA4WIMJWGUYTKNBYG5QTKYJVMQYGINJYGAYDEODEGFTA====", # <--- 新添加: 保煎烩
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
    "http://www.jintiankansha.me/rss/GEYTGOBYPQZDKNDGMQZDQMRWGA2TQZRRGU2DGMTEG4ZWMMRXME3TSODBGFSTENRQGBT": "中金点睛", # <--- 新添加
    "http://www.jintiankansha.me/rss/GMZTONZRHF6DQZJSMY4GIMJQG5TGMZBWHA4WIMJWGUYTKNBYG5QTKYJVMQYGINJYGAYDEODEGFTA====": "保煎烩", # <--- 新添加
}
# --- END RSS 配置 ---

# ========== OneNote 操作类 ==========
class OneNoteSync:
    def __init__(self):
        self.cache_file_path = os.path.abspath(os.path.join(os.path.dirname(__file__), TOKEN_CACHE_FILE))
        try:
            persistence = FilePersistence(self.cache_file_path)
            print(f"[缓存] 使用文件持久化: {self.cache_file_path}")
        except Exception as e:
            print(f"[缓存错误] 初始化 FilePersistence 失败: {e}. 尝试内存缓存。")
            persistence = None # 回退到内存缓存
        self.token_cache = PersistedTokenCache(persistence) if persistence else msal.SerializableTokenCache()
        self.app = PublicClientApplication(
            client_id=CLIENT_ID,
            authority=AUTHORITY,
            token_cache=self.token_cache
        )

    def get_token(self):
        """获取令牌，优先缓存，失败则根据环境选择交互式或设备代码流"""
        token_result = None
        accounts = self.app.get_accounts()
        if accounts:
            print(f"[认证] 找到账户缓存: {accounts[0].get('username', '未知用户')}")
            token_result = self.app.acquire_token_silent(SCOPES, account=accounts[0])

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
                    print("[认证] 等待用户在浏览器中完成验证 (通常有几分钟的等待时间)...")
                    sys.stdout.flush() # 确保信息立刻显示在 Actions 日志中
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

        if token_result and "access_token" in token_result:
            print("[认证] 成功获取访问令牌。")
            return token_result["access_token"]
        else:
            error_desc = token_result.get("error_description", "未知错误") if token_result else "获取流程失败"
            print(f"[认证失败] 未能获取访问令牌: {error_desc}")
            if token_result and "error" in token_result:
                print(f"  错误代码: {token_result.get('error')}")
                print(f"  Correlation ID: {token_result.get('correlation_id')}")
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
        except requests.exceptions.SSLError as ssl_err:
            print(f"[API 调用 SSL 错误] {ssl_err}")
            print("  >> 检查服务器 SSL/TLS 配置或更新客户端库。")
            return None
        except requests.exceptions.HTTPError as http_err:
            if http_err.response.status_code == 401:
                print("[API 调用错误] 401 未授权。")
            elif http_err.response.status_code == 403:
                print("[API 调用错误] 403 禁止访问。")
            elif http_err.response.status_code == 429:
                print("[API 调用错误] 429 请求过于频繁。")
            elif http_err.args and isinstance(http_err.args[0], requests.exceptions.SSLError):
                print(f"[API 调用 SSL 错误] {http_err}")
                print("  >> 检查服务器 SSL 配置或更新库。")
            else:
                print(f"[API 调用 HTTP 错误] {http_err.response.status_code} {http_err}")
            try:
                print(f"  错误详情: {http_err.response.json()}")
            except:
                print(f"  错误响应 (非JSON): {http_err.response.text}")
            return None
        except requests.exceptions.ConnectionError as conn_err:
            print(f"[API 调用连接错误] {conn_err}")
            return None
        except requests.exceptions.RequestException as req_err:
            print(f"[API 调用请求错误] {req_err}")
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
        if response and response.status_code == 201:
            try:
                response_data = response.json()
                page_link = response_data.get('links', {}).get('oneNoteWebUrl', {}).get('href')
            except Exception as e:
                print(f"  >> 解析 API 响应时出错: {e}")
            return True
        else:
            return False

# ========== RSS 处理函数 ==========
def get_user_agent():
    return {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/115.0'}

def clean_extracted_html(html_content):
    if not html_content or not isinstance(html_content, str) or html_content.startswith("[错误："):
        return html_content
    cleaned_html_str = html_content
    try:
        try:
            soup = BeautifulSoup(html_content, 'lxml')
        except Exception:
            soup = BeautifulSoup(html_content, 'html.parser')
        tags_to_remove = ['script', 'style', 'nav', 'footer', 'header', 'aside', 'form', 'iframe', 'button', '.sidebar', '#sidebar', '.related-posts', '.comments', '#comments', '.navigation', '.pagination', '.share-buttons', 'figure > figcaption', '.advertisement', '.ad', 'ins']
        for tag_or_selector in tags_to_remove:
            try:
                elements = soup.select(tag_or_selector) if tag_or_selector.startswith(('.', '#')) else soup.find_all(tag_or_selector)
                [el.decompose() for el in elements]
            except Exception:
                pass
        attributes_to_remove = ['style', 'class', 'id', 'onclick', 'onerror', 'onload', 'align', 'valign', 'bgcolor']
        tags_to_process = ['p', 'div', 'span', 'section', 'article', 'main', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'table', 'tr', 'td', 'th', 'strong', 'em', 'b', 'i', 'u', 'blockquote']
        for tag_name in tags_to_process:
            for tag in soup.find_all(tag_name):
                current_attrs = list(tag.attrs.keys())
                for attr in current_attrs:
                    if attr in attributes_to_remove:
                        del tag[attr]
        links_removed_count = 0
        for a_tag in soup.find_all('a'):
            a_tag.unwrap()
            links_removed_count += 1
        for p in soup.find_all('p'):
            if not p.get_text(strip=True) and not p.find(['img']):
                p.decompose()
        cleaned_html_str = str(soup)
        return cleaned_html_str
    except Exception as e:
        print(f"  >> [HTML 清理] 出错: {e}")
        return html_content

def get_full_content_from_link(url, source_name, feed_url, verify_ssl=True):
    print(f"  >> [内容/时间抓取] 开始处理链接: {url[:100]}...")
    extracted_date = None
    cleaned_html = None
    final_url = url
    is_jintiankansha_source = feed_url and 'jintiankansha.me' in feed_url
    if not url or not url.startswith('http'):
        return "[错误：无效的文章链接]", None
    try:
        print(f"  >> [网络请求] 尝试连接... (SSL 验证: {verify_ssl})")
        response = requests.get(url, headers=get_user_agent(), timeout=REQUEST_TIMEOUT, allow_redirects=True, verify=verify_ssl)
        response.raise_for_status()
        final_url = response.url
        print(f"  >> [内容/时间抓取] 最终访问 URL: {final_url[:100]}...")
        response.encoding = response.apparent_encoding if response.apparent_encoding else 'utf-8'
        print(f"  >> [内容/时间抓取] 链接抓取成功 (状态码: {response.status_code}, 编码: {response.encoding})。")
        try:
            soup = BeautifulSoup(response.text, 'lxml')
        except Exception as e:
            print(f"  >> [内容/时间抓取] 警告：lxml 解析失败 ({e})，回退 html.parser。")
            soup = BeautifulSoup(response.text, 'html.parser')

        # --- 时间提取 ---
        if source_name and (source_name == '和讯保险' or source_name.startswith('和讯-')): # 规则保留, 涵盖修改后的和讯来源
            print(f"  >> [时间抓取] 来源 '{source_name}'，尝试和讯规则...")
            try:
                date_span = soup.select_one('span.pr20')
                if not date_span:
                    tip_div = soup.select_one('div.tip.fl.gray')
                    date_span = tip_div.find('span') if tip_div else None
                if date_span:
                    date_text = date_span.get_text(strip=True)
                    match = re.search(r'(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})', date_text)
                    if match:
                        date_str = match.group(1)
                        try:
                            extracted_date = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
                            print(f"  >> [时间抓取] **成功** (和讯规则): {extracted_date}")
                        except ValueError:
                            print(f"  >> [时间抓取] 警告：找到日期 '{date_str}' 但格式不匹配 '%Y-%m-%d %H:%M:%S'")
                    else: print(f"  >> [时间抓取] 警告：未匹配到日期格式: {date_text}")
                else: print("  >> [时间抓取] 未能在和讯页面找到日期元素。")
            except Exception as e: print(f"  >> [时间抓取] 尝试从和讯页面提取时间时出错: {e}")
        elif source_name and source_name.startswith('总局'): # 规则保留 (例如 "总局公众号", "总局官网")
            print(f"  >> [时间抓取] 来源 '{source_name}'，尝试总局规则...")
            try:
                info_div = soup.select_one('div.pages-detail-info')
                found_date_str = None
                if info_div:
                    for span in info_div.find_all('span', recursive=False):
                        span_text = span.get_text(strip=True)
                        if "发布日期：" in span_text:
                            date_part = span_text.split("发布日期：")[-1].strip()
                            match = re.search(r'^(\d{4}-\d{2}-\d{2})', date_part)
                            if not match: match = re.search(r'^(\d{4}年\d{1,2}月\d{1,2}日)', date_part)
                            if match:
                                found_date_str = match.group(1); break
                    if found_date_str:
                        try:
                            if '-' in found_date_str: extracted_date = datetime.strptime(found_date_str, '%Y-%m-%d')
                            else: extracted_date = datetime.strptime(found_date_str, '%Y年%m月%d日')
                            print(f"  >> [时间抓取] **成功** (总局规则): {extracted_date}")
                        except ValueError: print(f"  >> [时间抓取] 警告：找到日期 '{found_date_str}' 但格式无法解析")
                    else: print("  >> [时间抓取] 未在信息区找到日期。")
                else: print("  >> [时间抓取] 未找到日期信息区域。")
            except Exception as e: print(f"  >> [时间抓取] 尝试从总局页面提取时间时出错: {e}")
        elif is_jintiankansha_source:
            print(f"  >> [时间抓取] 检测到来源 '{source_name}' (jintiankansha)，尝试微信规则...")
            try:
                time_em = soup.select_one('em#publish_time')
                if time_em:
                    time_text = time_em.get_text(strip=True)
                    try:
                        if ' ' in time_text: extracted_date = datetime.strptime(time_text, '%Y-%m-%d %H:%M')
                        else: extracted_date = datetime.strptime(time_text, '%Y-%m-%d')
                        print(f"  >> [时间抓取] **成功** (Jintiankansha/微信规则): {extracted_date}")
                    except ValueError: print(f"  >> [时间抓取] 警告：找到微信时间 '{time_text}' 但格式无法解析")
                else: print("  >> [时间抓取] 未找到微信时间元素 (em#publish_time)。")
            except Exception as e: print(f"  >> [时间抓取] 尝试从 jintiankansha/微信页面提取时间时出错: {e}")
        else: print(f"  >> [时间抓取] 来源 '{source_name}'，非特定来源，不尝试网页提取。")

        # --- 内容提取 ---
        selectors = []
        if source_name and source_name.startswith('总局'):
            selectors.extend(['div.content', 'div#zoom', '.pages_content', 'div.view.TRS_UEDITOR.trs_paper_default.trs_word'])
        elif source_name and (source_name == '和讯保险' or source_name.startswith('和讯-')): # 同样更新和讯规则适用条件
            selectors.extend(['div.art_contextBox', 'div.art_context'])
        elif is_jintiankansha_source:
            selectors.extend(['div.rich_media_content', 'div#js_content', 'div.wx-content'])
        selectors.extend([ 'article', '.article-content', '.entry-content', '.post-content', '.post-body', '#article-body', '#entry-content', 'div[itemprop="articleBody"]', 'div.main-content', 'div.entry', 'div[class*=content i]', 'div[class*=article i]', 'div[class*=post i]', 'div[class*=body i]', 'div[id*=content i]', 'div[id*=article i]', 'div[id*=post i]', 'div[id*=body i]', 'main' ])
        article_body = None; found_selector = None
        for selector in selectors:
            try:
                target = soup.select_one(selector)
                if target:
                    text_len = len(target.get_text(strip=True))
                    if text_len > 30 or target.find('img'):
                        article_body = target
                        found_selector = selector
                        print(f"  >> [内容抓取] 找到内容区域 (选择器: {selector})。")
                        break
            except Exception as e: continue
        if article_body:
            print(f"  >> [内容抓取] 最终选用选择器: {found_selector}")
            print("  >> [图片路径修复] 开始处理图片链接...")
            img_fixed_count = 0
            if article_body:
                for img in article_body.find_all('img'):
                    try:
                        src = img.get('src') or img.get('data-src')
                        if src:
                            if not src.startswith(('http://', 'https://', 'data:')):
                                try:
                                    absolute_src = urljoin(final_url, src)
                                    img['src'] = absolute_src
                                    img_fixed_count += 1
                                except Exception as url_e:
                                    pass
                            elif 'data-src' in img.attrs and img.get('src') != img['data-src']:
                                data_src = img['data-src']
                                if data_src and not data_src.startswith(('http://', 'https://', 'data:')):
                                    try:
                                        absolute_src = urljoin(final_url, data_src)
                                        img['src'] = absolute_src
                                        img_fixed_count += 1
                                    except Exception as url_e:
                                        pass
                                else:
                                    img['src'] = data_src
                    except Exception as img_e:
                        pass
            if img_fixed_count > 0: print(f"  >> [图片路径修复] 完成，共修复/处理 {img_fixed_count} 个。")
            else: print("  >> [图片路径修复] 未找到需修复或处理的图片链接。")
            content_before_cleaning = str(article_body)
            cleaned_html = clean_extracted_html(content_before_cleaning)
            print(f"  >> [内容抓取] 清理完成 (最终长度: {len(cleaned_html)})。")
            return cleaned_html, extracted_date
        else: print(f"  >> [内容抓取] **失败**：未能找到主要内容区域。"); return None, extracted_date
    except requests.exceptions.SSLError as ssl_err: return "[错误：SSL 连接失败]", None
    except requests.exceptions.ConnectionError as conn_err: return "[错误：连接失败]", None
    except requests.exceptions.Timeout: return "[错误：抓取原文链接超时]", None
    except requests.exceptions.RequestException as e: return f"[错误：抓取原文链接失败 - {e}]", None
    except Exception as e: return f"[错误：解析原文 HTML 时出错 - {e}]", None

def fetch_rss_feeds():
    print("\n[开始抓取所有 RSS Feeds...]")
    all_entries = []
    processed_urls_in_run = set()
    for feed_url in ORIGINAL_FEEDS:
        if feed_url in processed_urls_in_run:
            continue
        source_name_from_dict = FEED_SOURCES.get(feed_url, feed_url)
        print(f"  处理中: {source_name_from_dict} ({feed_url})")
        actual_url = feed_url
        try:
            feed_data = feedparser.parse(feed_url, agent=get_user_agent()['User-Agent'], request_headers=get_user_agent())
            
            if hasattr(feed_data, 'href') and feed_data.href != feed_url:
                print(f"  >> Feed URL 重定向到: {feed_data.href}")
                actual_url = feed_data.href
                FEED_SOURCES.setdefault(actual_url, source_name_from_dict) 
                processed_urls_in_run.add(actual_url) 

            if feed_data.bozo:
                print(f"  警告: 解析 Feed '{source_name_from_dict}' 时问题: {feed_data.bozo_exception}")
            
            if not feed_data.entries:
                print("  警告: 未获取到任何条目。")
                continue

            for entry in feed_data.entries:
                entry_id = entry.get('id', entry.get('link')) or f"{entry.get('title', '无标题')}_{entry.get('published', entry.get('updated', time.time()))}"
                
                published_time_rss = None
                pub_parsed = entry.get('published_parsed') or entry.get('updated_parsed')
                if pub_parsed:
                    try:
                        published_time_rss = datetime.fromtimestamp(time.mktime(pub_parsed))
                    except (ValueError, OverflowError) as e:
                        print(f"  警告: 条目 '{entry.get('title', '')[:30]}...' RSS 时间戳无效: {pub_parsed}, Error: {e}")
                        published_time_rss = None
                
                title = entry.get('title', '无标题').strip()
                link = entry.get('link', '').strip()

                content_summary = ""
                if 'content' in entry and entry.content:
                    html_content = next((c.value for c in entry.content if c.type and 'html' in c.type.lower()), None)
                    content_summary = html_content if html_content else next((f"<p>{html.escape(c.value)}</p>" for c in entry.content if c.type and 'text' in c.type.lower()), "")
                if not content_summary and 'summary' in entry:
                    content_summary = entry.summary
                elif not content_summary and 'description' in entry:
                    content_summary = entry.description

                if not entry_id or not title:
                    continue
                
                # Determine the source name, special handling for Hexun
                current_source_name = FEED_SOURCES.get(actual_url, source_name_from_dict)
                display_source_name = current_source_name # Default display name
                
                # --- MODIFICATION FOR HEXUN SOURCE ---
                if current_source_name == "和讯保险" and " | " in title:
                    try:
                        channel_part = title.split(" | ")[0].strip()
                        if channel_part: # Make sure channel_part is not empty
                             display_source_name = f"和讯-{channel_part}"
                    except Exception as e:
                        print(f"  >> 警告: 尝试为和讯保险提取频道名称时出错: {e}")
                        # display_source_name remains "和讯保险"
                # --- END MODIFICATION ---

                all_entries.append({
                    'id': entry_id,
                    'title': title, # Store original title for potential channel extraction later
                    'link': link,
                    'published_time_rss': published_time_rss,
                    'content_summary': content_summary,
                    'source_name': current_source_name, # This is the name from FEED_SOURCES
                    'display_source_name': display_source_name, # This is for OneNote page
                    'feed_url': actual_url 
                })
        except Exception as e:
            print(f"  错误: 抓取 Feed '{source_name_from_dict}' 时出错: {e}")
            import traceback
            traceback.print_exc()
            continue
        finally:
            processed_urls_in_run.add(feed_url)

    print(f"\n[RSS抓取完成] 共找到 {len(all_entries)} 条有效记录。")
    all_entries.sort(key=lambda x: x['published_time_rss'] or datetime.min, reverse=True)
    return all_entries

def load_processed_items(filename=PROCESSED_ITEMS_FILE):
    processed = set()
    filepath = os.path.join(os.path.dirname(__file__), filename)
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                processed.update(line.strip() for line in f if line.strip())
            print(f"[状态] 已加载 {len(processed)} 个已处理条目 ID 从 {filename}")
        except Exception as e:
            print(f"警告：读取已处理条目文件 {filename} 失败: {e}")
    else:
        print(f"[状态] 未找到已处理条目文件 {filename}。")
    return processed

def save_processed_items(new_ids, filename=PROCESSED_ITEMS_FILE):
    if not new_ids:
        return
    filepath = os.path.join(os.path.dirname(__file__), filename)
    try:
        with open(filepath, 'a', encoding='utf-8') as f:
            f.writelines(f"{item_id}\n" for item_id in new_ids if item_id)
        print(f"[状态] 已将 {len(new_ids)} 个新处理条目 ID 保存到 {filename}")
    except Exception as e:
        print(f"错误：保存已处理条目文件 {filename} 失败: {e}")

# ========== 主程序 ==========
if __name__ == "__main__":
    print(f"=== RSS to OneNote Sync (Consumer Scope - Full Run) 开始于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
    run_start_time = time.time()

    if not CLIENT_ID:
        exit("错误：AZURE_CLIENT_ID 未能加载。请检查 .env 文件或环境变量。")
    print("\n[环境变量验证通过]")

    processed_ids = load_processed_items()
    newly_processed_ids_in_run = set()

    all_entries = fetch_rss_feeds()

    new_entries = [entry for entry in all_entries if entry['id'] not in processed_ids]

    print(f"\n[筛选结果] 发现 {len(new_entries)} 条新条目。")

    if not new_entries:
        print("没有需要同步的新条目。")
    else:
        entries_to_process = new_entries[:MAX_ITEMS_PER_RUN]
        print(f"本次计划同步最多 {len(entries_to_process)} 条新条目。 (上限: {MAX_ITEMS_PER_RUN})")

        print("\n[初始化 OneNote 连接...]")
        onenote = OneNoteSync()
        initial_token = onenote.get_token()
        if not initial_token:
            exit("错误：无法获取 OneNote 访问令牌。同步终止。")
        print("OneNote 连接初始化成功。")

        print("\n[开始同步新条目到 OneNote...]")
        success_count = 0
        fail_count = 0
        total_to_process = len(entries_to_process)

        for i, entry in enumerate(entries_to_process):
            item_start_time = time.time()
            # Use 'display_source_name' for logging if available, otherwise 'source_name'
            log_source_name = entry.get('display_source_name', entry['source_name'])
            print(f"\n--- 处理第 {i+1}/{total_to_process} 条: {entry['title'][:60]} ({log_source_name}) ---")
            print(f"  原始链接 (Link): {entry['link']}")

            print("  * 正在准备页面内容和时间...")
            VERIFY_SSL_TEMP = True
            # Pass the general source_name for get_full_content_from_link logic
            fetched_content, fetched_date = get_full_content_from_link(entry['link'], entry['source_name'], entry['feed_url'], verify_ssl=VERIFY_SSL_TEMP)

            final_time = fetched_date or entry.get('published_time_rss')
            time_source = "网页提取" if fetched_date else ("RSS Feed" if entry.get('published_time_rss') else "无")
            print(f"  >> 决定使用时间来源: {time_source}")
            if not final_time:
                print("  >> 警告：最终无有效发布时间。")
            
            # OneNote Page Title Preparation
            # For Hexun, if channel was extracted, the entry['title'] is still "Channel | Actual Title"
            # We need to decide if we want to remove the channel prefix from the OneNote title
            onenote_page_title_content = entry['title']
            if entry['source_name'] == "和讯保险" and " | " in entry['title']:
                try:
                    # Use the part after " | " as the actual title for OneNote page
                    onenote_page_title_content = entry['title'].split(" | ", 1)[1].strip()
                except IndexError:
                    # If split fails or no second part, use original title
                    onenote_page_title_content = entry['title']


            date_prefix = (final_time or datetime.now()).strftime('%y%m%d')
            separator = "-" if onenote_page_title_content and onenote_page_title_content[0].isdigit() else ""
            formatted_title = f"{date_prefix}{separator}{onenote_page_title_content}"[:200]
            print(f"  * 目标标题: {formatted_title}")

            final_body_html = ""
            if fetched_content and isinstance(fetched_content, str) and fetched_content.startswith("[错误："):
                print(f"  >> 获取原文内容时发生错误 ({fetched_content})，尝试使用 RSS 摘要。")
                final_body_html = clean_extracted_html(entry['content_summary'])
            elif fetched_content:
                final_body_html = fetched_content
            else:
                print(f"  >> 获取原文内容失败或无效，尝试使用 RSS 摘要。")
                final_body_html = clean_extracted_html(entry['content_summary'])
            
            if not final_body_html or (isinstance(final_body_html, str) and final_body_html.startswith("[错误：")):
                print("  >> RSS 摘要内容也为空或清理失败。")
                final_body_html = "<p><i>[无法自动提取正文，请访问原文链接查看]</i></p>"

            publish_time_str = final_time.strftime('%Y-%m-%d %H:%M:%S') if final_time else "未知"
            escaped_link = html.escape(entry['link'] or '')
            # Use the 'display_source_name' for the OneNote page
            escaped_display_source_name = html.escape(entry.get('display_source_name', entry['source_name']))
            
            onenote_content = f"""
            <h1>{html.escape(formatted_title)}</h1>
            <p style="font-size:9pt; color:gray;">发布时间: {publish_time_str} | 来源: {escaped_display_source_name} | <a href="{escaped_link}">原文链接</a></p>
            <hr/>
            <div>
                {final_body_html}
            </div>
            """

            creation_success = onenote.create_onenote_page_in_app_notebook(title=formatted_title, content_html=onenote_content)

            current_entry_id = entry['id']
            if creation_success:
                success_count += 1
                newly_processed_ids_in_run.add(current_entry_id)
                print(f"  >> 处理成功 (耗时: {time.time() - item_start_time:.2f} 秒)")
            else:
                fail_count += 1
                print(f"  !! 处理失败: {entry['title'][:60]}")

            time.sleep(REQUEST_DELAY)

        print(f"\n[同步统计] 本次成功 {success_count} 条，失败 {fail_count} 条。")
        if newly_processed_ids_in_run:
            save_processed_items(newly_processed_ids_in_run)

    run_end_time = time.time()
    print(f"\n=== 同步完成于: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} (总耗时: {run_end_time - run_start_time:.2f} 秒) ===")