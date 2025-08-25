# ... (所有 import 保持不变) ...
# ... (所有配置保持不变) ...

# ========== OneNote 操作类 ==========
class OneNoteSync:
    # ... (__init__ 保持不变) ...
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
        """获取令牌，优先缓存，失败则根据环境选择交互式或设备代码流"""
        token_result = None
        accounts = self.app.get_accounts()
        was_silent_successful = False # 标记是否静默获取成功
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
                    print("[认证] 等待用户在浏览器中完成验证 (通常有几分钟的等待时间)...")
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
        
        # --- 关键修改：认证成功后，如果不是静默获取的，就立即保存一次缓存 ---
        if token_result and "access_token" in token_result and not was_silent_successful:
            if isinstance(self.token_cache, PersistedTokenCache):
                try:
                    self.token_cache.persistence.write(self.token_cache.serialize())
                    print("[缓存] 已尝试在认证成功后立即保存令牌缓存。")
                except Exception as e:
                    print(f"[缓存] 警告: 尝试立即保存令牌缓存时出错: {e}")
        # --- 结束修改 ---

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

    # ... (_make_api_call 和 create_onenote_page_in_app_notebook 保持不变) ...
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

# ... (所有 RSS 处理函数和主程序 main 逻辑保持不变) ...
# ... (从 def get_user_agent(): 到脚本结尾的所有代码都和您之前的文件一样，此处省略) ...
