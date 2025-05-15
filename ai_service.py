import os
import json
import urllib.request
import re
import logging
import datetime
import traceback
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK


class AIService:    
    def __init__(self, ctx):
        self.ctx = ctx
        # 存儲上次回應的token數量 (統一使用這個變數)
        self.previous_token = None
        # 長度調整因數，默認為1.0
        self.length_adjustment_factor = 1.0
        # 初始化日誌系統
        self.setup_logging()
        
    def setup_logging(self):
        """設定日誌系統"""
        try:
            # 創建 .libreoffice 目錄（如果不存在）
            libreoffice_dir = os.path.join(os.path.expanduser("~"), ".libreoffice")
            if not os.path.exists(libreoffice_dir):
                os.makedirs(libreoffice_dir)
                
            # 創建日誌目錄
            logs_dir = os.path.join(libreoffice_dir, "logs")
            if not os.path.exists(logs_dir):
                os.makedirs(logs_dir)
                
            # 設定日誌文件名（按日期）
            today = datetime.datetime.now().strftime("%Y-%m-%d")
            log_file = os.path.join(logs_dir, f"ai_query_{today}.log")
            
            # 配置日誌處理器
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(formatter)
            
            # 獲取日誌記錄器並設置級別
            self.logger = logging.getLogger('ai_service')
            self.logger.setLevel(logging.INFO)
            
            # 清除可能存在的處理器
            if self.logger.handlers:
                for handler in self.logger.handlers:
                    self.logger.removeHandler(handler)
            
            # 添加處理器
            self.logger.addHandler(file_handler)
            
            self.logger.info("====== AI 服務日誌系統已初始化 ======")
        except Exception as e:
            print(f"無法設置日誌系統: {str(e)}")
            self.logger = None

    def show_message(self, message, title="Information", message_type=INFOBOX):
        """顯示訊息對話框"""
        toolkit = self.ctx.ServiceManager.createInstance("com.sun.star.awt.Toolkit")
        parent = toolkit.getActiveTopWindow()
        mb = toolkit.createMessageBox(
            parent, message_type, BUTTONS_OK, title, str(message))
        mb.execute()

    def extract_length_adjustment(self, prompt):
        """解析提示詞中的長度調整參數"""
        # 定義可能的長度調整表達方式
        patterns = [
            # 中文表達
            r'(擴展|增加|加長|延長)\s*(\d+)%',
            r'(縮減|減少|縮短|減短)\s*(\d+)%',
            # 英文表達
            r'(expand|increase|lengthen|extend)\s*(\d+)%',
            r'(reduce|decrease|shorten)\s*(\d+)%'
        ]
        
        # 遍歷所有模式嘗試匹配
        for pattern in patterns:
            matches = re.search(pattern, prompt, re.IGNORECASE)
            if matches:
                action_type = matches.group(1)
                percentage = int(matches.group(2))
                # 判斷是增加還是減少
                increase_keywords = ['擴展', '增加', '加長', '延長', 'expand', 'increase', 'lengthen', 'extend']
                if any(keyword in action_type.lower() for keyword in increase_keywords):
                    return f"+{percentage}%"
                else:
                    return f"-{percentage}%"
        
        return None
        
    def estimate_token_count(self, text, provider="gemini"):
        """更准确地估算文本的token数量"""
        # 先尝试调用API获取精确的token数量
        try:
            token_count = self.get_token_count_from_api(text, provider)
            if token_count:
                if hasattr(self, 'logger') and self.logger:
                    self.logger.info(f"从API获取到精确的token数量: {token_count}")
                return token_count
        except Exception as e:
            if hasattr(self, 'logger') and self.logger:
                self.logger.warning(f"无法从API获取token数量: {str(e)}, 使用本地估算方法")
        
        # 如果API方法失败，使用改进的本地估算方法
        
        # 分析文本组成
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
        # 数字
        numbers = len(re.findall(r'\d', text))
        # 英文单词
        english_words = len(re.findall(r'\b[a-zA-Z]+\b', text))
        # 标点符号
        punct_chars = len(re.findall(r'[^\w\s\u4e00-\u9fff]', text))
        # 空白字符
        whitespace = len(re.findall(r'\s', text))
        # 其他字符
        other_chars = len(text) - chinese_chars - numbers - punct_chars - whitespace
        
        # 使用更准确的权重计算
        # 中文字符权重 - 根据Gemini/GPT等模型，中文每个字符约为0.5-0.7个token
        chinese_weight = 0.7
        # 英文单词权重 - 平均每个英文单词约1.3个token
        english_word_weight = 1.3
        # 数字权重 - 数字通常比字母更高效编码
        number_weight = 0.5
        # 标点符号权重
        punct_weight = 0.5
        # 空白字符权重
        whitespace_weight = 0.3
        # 其他字符权重
        other_weight = 1.0
        
        # 计算总token数
        estimated_tokens = (
            chinese_chars * chinese_weight +
            english_words * english_word_weight +
            numbers * number_weight +
            punct_chars * punct_weight +
            whitespace * whitespace_weight +
            other_chars * other_weight
        )
        
        # 向上取整
        token_count = int(estimated_tokens + 0.5)
        
        if hasattr(self, 'logger') and self.logger:
            self.logger.info(f"文本组成分析: 中文字符 {chinese_chars}, 英文单词 {english_words}, 数字 {numbers}, 标点 {punct_chars}, 空白字符 {whitespace}, 其他字符 {other_chars}")
            self.logger.info(f"Token 估算: 文本长度 {len(text)} 字符, 估算 {token_count} tokens")
        
        return token_count

    def get_token_count_from_api(self, text, provider="gemini"):
        """尝试从API获取精确的token数量"""
        if not text:
            return 0
            
        try:
            # 加载API设置
            libreoffice_dir = os.path.join(os.path.expanduser("~"), ".libreoffice")
            env_path = os.path.join(libreoffice_dir, ".env")
            
            api_key = ""
            
            if os.path.exists(env_path):
                with open(env_path, 'r') as f:
                    content = f.read()
                    for line in content.splitlines():
                        if line.startswith("DEFAULT_PROVIDER="):
                            provider = line.split("=")[1].strip('"\'')
                        elif "API_KEY=" in line:
                            api_key = line.split("=")[1].strip('"\'')
            
            if not api_key:
                return None
                
            # 根据提供商选择不同的API调用方式
            if provider == "gemini":
                # Google Gemini API的countTokens端点
                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:countTokens?key={api_key}"
                data = {
                    "contents": [{"parts": [{"text": text}]}]
                }
                headers = {'Content-Type': 'application/json'}
                
                data_bytes = json.dumps(data).encode('utf-8')
                req = urllib.request.Request(url, data=data_bytes, headers=headers)
                
                with urllib.request.urlopen(req) as response:
                    result = json.loads(response.read().decode('utf-8'))
                    return result.get('totalTokens', None)
                    
            elif provider == "openai":
                # OpenAI API - 使用tiktoken库模拟，此处不实现
                # 若需要实际实现，需要在服务端集成tiktoken库
                return None
                
            elif provider == "claude":
                # Anthropic Claude API - 目前无直接的计数API
                return None
                
        except Exception as e:
            if hasattr(self, 'logger') and self.logger:
                self.logger.error(f"获取API token数量失败: {str(e)}")
            return None

    def get_target_token_count(self, length_adjustment, current_token_count=None):
        length_adjustments = {
            "-75%": 0.4, "-50%": 0.6, "-25%": 0.8,
            "+25%": 2.0, "+50%": 3.0, "+75%": 4.0
        }
        
        adjustment_factor = length_adjustments.get(length_adjustment)
        if not adjustment_factor and length_adjustment:
            match = re.match(r'([+-])(\d+)%', length_adjustment)
            if match:
                sign, percentage = match.group(1), int(match.group(2)) / 100
                adjustment_factor = 1 + percentage if sign == '+' else 1 - percentage
                
        if current_token_count is None:
            current_token_count = self.previous_token
            if hasattr(self, 'logger') and self.logger:
                self.logger.info(f"使用 previous_token: {current_token_count}")
                
        if current_token_count and adjustment_factor:
            target_count = int(current_token_count * adjustment_factor)
            if hasattr(self, 'logger') and self.logger:
                self.logger.info(f"計算目標 token 數: {current_token_count} * {adjustment_factor} = {target_count}")
            return target_count

        if hasattr(self, 'logger') and self.logger:
            self.logger.warning("無法計算目標 token 數（缺少 token 或調整因數）")
        return None

    def create_adjustment_prompt(self, original_text, current_token_count, target_token_count):
        """創建用於調整文本長度的提示"""
        # 计算目标范围 - 使用更严格的范围，确保调整更精确
        lower_bound = int(target_token_count * 0.9)
        upper_bound = int(target_token_count * 1.1)
        
        # 计算与目标的百分比差异
        diff_percentage = int(abs(current_token_count - target_token_count) / target_token_count * 100)
        
        # 如果當前token數超過了上限
        if current_token_count > upper_bound:
            # 計算需要減少的百分比
            reduction_percentage = int((current_token_count - target_token_count) / current_token_count * 100)
            return f"""
【精確token調整】當前內容有{current_token_count}個tokens，目標為{target_token_count}個tokens (容許範圍:{lower_bound}-{upper_bound})。
您需要將内容縮減約{reduction_percentage}%，確保最終token數在目標範圍內。

縮減指南：
1. 删除重复或冗余信息
2. 简化句子结构，减少修饰词
3. 合并相似段落
4. 保留核心观点和关键信息
5. 确保逻辑连贯性不受影响

原始内容：
{original_text}
"""
        # 如果當前token數低於下限
        elif current_token_count < lower_bound:
            # 計算需要增加的百分比
            increase_percentage = int((target_token_count - current_token_count) / current_token_count * 100)
            
            return f"""
【精確token調整】當前內容有{current_token_count}個tokens，目標為{target_token_count}個tokens (容許範圍:{lower_bound}-{upper_bound})。
您需要將内容擴展約{increase_percentage}%，確保最終token數在目標範圍內。

擴展指南：
1. 增加相关的例子、证据或细节
2. 扩展对关键概念的解释
3. 添加相关的背景信息
4. 保持连贯性，不添加无关内容
5. 文风和语气保持一致

原始内容：
{original_text}
"""
        # 如果已經在範圍內，不需要調整
        return None    
    
    def ask_ai_with_length_adjustment(self, question, length_adjustment=None, max_attempts=3):
        """使用長度調整功能發送請求到AI服務"""
        try:
            # 在方法開始時就保存當前的previous_token，整個方法中都使用這個值
            previous_token_value = self.previous_token
            
            # 記錄長度調整參數
            if hasattr(self, 'logger') and self.logger:
                self.logger.info(f"====== 開始長度調整流程 ======")
                self.logger.info(f"長度調整參數: {length_adjustment}, 調用前token數: {previous_token_value}")
            
            # 從提示中提取長度調整指示
            if not length_adjustment:
                length_adjustment = self.extract_length_adjustment(question)
                
            # 如果沒有長度調整指示，不進行調整
            if not length_adjustment:
                if hasattr(self, 'logger') and self.logger:
                    self.logger.info(f"未找到長度調整參數，不進行調整")
                initial_response = self.ask_ai(question)
                current_token_count = self.estimate_token_count(initial_response)
                self.previous_token = current_token_count
                return initial_response
            
            # 獲取初始回應
            initial_response = self.ask_ai(question)
                
            # 估算當前回應的token數
            current_token_count = self.estimate_token_count(initial_response)

            # 計算目標token數（一律使用previous_token_value而不是current_token_count）
            target_token_count = self.get_target_token_count(length_adjustment, previous_token_value)
            
            # 如果無法計算目標token數，直接返回初始結果
            if not target_token_count:
                if hasattr(self, 'logger') and self.logger:
                    self.logger.warning("無法計算目標 token 數量，返回初始回應")
                return initial_response
                
            # 檢查初始回應的token數是否已經接近目標
            token_diff = abs(current_token_count - target_token_count)
            token_diff_percent = (token_diff / target_token_count) * 100
            
            # 如果已經足夠接近目標（差距小於10%），直接返回初始回應
            if token_diff_percent < 10:
                if hasattr(self, 'logger') and self.logger:
                    self.logger.info(f"初始回應的 token 數已接近目標 (差距 {token_diff_percent:.2f}% < 10%)，不需調整")
                return initial_response
                
            # 準備進行長度調整
            best_response = initial_response
            best_token_count = current_token_count
            best_token_diff = token_diff
            
            # 多次嘗試調整長度
            for attempt in range(1, max_attempts):
                # 計算與目標的差距
                token_diff_percent = (abs(best_token_count - target_token_count) / target_token_count) * 100
                
                if hasattr(self, 'logger') and self.logger:
                    self.logger.info(f"調整嘗試 #{attempt}, 當前 token 數: {best_token_count}, 與目標 {target_token_count} 的差距: {token_diff_percent:.2f}%")
                
                # 如果差距在可接受範圍內（10%以內），停止調整
                if token_diff_percent <= 10:
                    if hasattr(self, 'logger') and self.logger:
                        self.logger.info(f"差距在可接受範圍內 ({token_diff_percent:.2f}% ≤ 10%), 停止調整")
                    break
                    
                # 創建調整提示
                adjustment_prompt = self.create_adjustment_prompt(
                    best_response, best_token_count, target_token_count
                )
                
                # 如果沒有需要調整的提示，停止調整
                if not adjustment_prompt:
                    if hasattr(self, 'logger') and self.logger:
                        self.logger.info("無需調整提示，停止調整")
                    break
                    
                # 發送調整請求
                try:
                    if hasattr(self, 'logger') and self.logger:
                        self.logger.info(f"發送第 {attempt} 次調整請求")
                        
                    adjusted_response = self.ask_ai(adjustment_prompt)
                    adjusted_token_count = self.estimate_token_count(adjusted_response)
                    adjusted_token_diff = abs(adjusted_token_count - target_token_count)
                    
                    if hasattr(self, 'logger') and self.logger:
                        self.logger.info(f"調整後 token 數: {adjusted_token_count}, 與目標差距: {adjusted_token_diff} tokens ({(adjusted_token_diff / target_token_count * 100):.2f}%)")
                    
                    # 如果調整後的結果更接近目標，採用這個結果
                    if adjusted_token_diff < best_token_diff:
                        if hasattr(self, 'logger') and self.logger:
                            self.logger.info(f"採用新調整結果: 從 {best_token_count} tokens → {adjusted_token_count} tokens")
                            self.logger.info(f"改善程度: {best_token_diff - adjusted_token_diff} tokens ({(best_token_diff - adjusted_token_diff) / target_token_count * 100:.2f}%)")
                            
                        best_token_diff = adjusted_token_diff
                        best_response = adjusted_response
                        best_token_count = adjusted_token_count
                    else:
                        if hasattr(self, 'logger') and self.logger:
                            self.logger.info(f"保留先前結果，新調整沒有改善（{best_token_count} tokens 優於 {adjusted_token_count} tokens）")
                except Exception as e:
                    # 調整失敗，繼續使用最佳結果
                    if hasattr(self, 'logger') and self.logger:
                        self.logger.error(f"調整請求失敗: {str(e)}")
                    pass
            
            if hasattr(self, 'logger') and self.logger:
                self.logger.info(f"長度調整完成，最終 token 數: {best_token_count}")
                self.logger.info(f"目標 token 數: {target_token_count}, 最終差異: {best_token_diff} tokens ({(best_token_diff / target_token_count * 100):.2f}%)")
                if current_token_count != best_token_count:
                    change_percentage = ((best_token_count - current_token_count) / current_token_count * 100)
                    direction = "增加" if change_percentage > 0 else "減少"
                    self.logger.info(f"與初始回應相比: {direction} {abs(change_percentage):.2f}% ({abs(best_token_count - current_token_count)} tokens)")
                self.logger.info("====== 長度調整流程完成 ======")
            
            # 只在方法最後更新previous_token
            self.previous_token = best_token_count  # 確保這一行只在方法末尾出現一次

            return best_response
        except Exception as e:
            if hasattr(self, 'logger') and self.logger:
                self.logger.error(f"長度調整過程出錯: {str(e)}")
            raise Exception(f"長度調整過程出錯: {str(e)}")
            
    def validate_api_key(self, api_key, provider):
        """驗證API金鑰是否有效"""
        try:
            # 根據不同提供商設置測試URL和請求數據
            if provider == "gemini":
                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash?key={api_key}"
                headers = {'Content-Type': 'application/json'}
                req = urllib.request.Request(url, headers=headers, method='GET')
            elif provider == "openai":
                url = "https://api.openai.com/v1/models"
                headers = {'Authorization': f'Bearer {api_key}'}
                req = urllib.request.Request(url, headers=headers, method='GET')
            elif provider == "claude":
                url = "https://api.anthropic.com/v1/models"
                headers = {
                    'x-api-key': api_key,
                    'anthropic-version': '2023-06-01'
                }
                req = urllib.request.Request(url, headers=headers, method='GET')
            elif provider == "mistral":
                url = "https://api.mistral.ai/v1/models"
                headers = {'Authorization': f'Bearer {api_key}'}
                req = urllib.request.Request(url, headers=headers, method='GET')
            else:
                return False, f"不支援的AI提供商: {provider}"
                
            # 發送請求檢查API金鑰有效性
            try:
                with urllib.request.urlopen(req, timeout=5) as response:
                    if response.getcode() == 200:
                        return True, "API金鑰有效"
                    else:
                        return False, f"API請求返回錯誤狀態碼: {response.getcode()}"
            except urllib.error.URLError as e:
                return False, f"API金鑰驗證失敗: {str(e)}"
                
        except Exception as e:
            if hasattr(self, 'logger') and self.logger:
                self.logger.error(f"API金鑰驗證過程出錯: {str(e)}")
            return False, f"API金鑰驗證過程出錯: {str(e)}"
            
    def ask_ai(self, question, dialog=None, generate_prompt=False, selected_options=None, config_manager=None):
        """
        直接發送請求到AI服務API

        Args:
            question: 問題或提示詞
            dialog: 對話框引用
            generate_prompt: 是否僅生成提示詞模板
            selected_options: 選擇的選項字典
            config_manager: 配置管理器實例
            
        Returns:
            str: AI的回應文本或在錯誤情況下的錯誤訊息
        """
        # 在類中添加一個屬性來存儲上次的token信息
        self.last_token_info = None
        self.token_info_str = None  # 新增一個字符串版本的token信息
        try:
            # 記錄API請求
            if hasattr(self, 'logger') and self.logger:
                # 限制日誌中的提示詞長度
                log_question = question[:200] + "..." if len(question) > 200 else question
                self.logger.info(f"API 請求: {log_question}")
                
            # 如果是生成提示詞模式且傳入了config_manager
            if generate_prompt and selected_options and config_manager:
                return config_manager.generate_adjustment_prompt(
                    selected_options=selected_options,
                    text=""  # 空文本，因為我們只是生成模板
                )
        
            # 載入API設定
            libreoffice_dir = os.path.join(os.path.expanduser("~"), ".libreoffice")
            env_path = os.path.join(libreoffice_dir, ".env")
        
            provider = "gemini"  # 預設
            api_key = ""
            model = ""  # 添加模型變數
        
            if os.path.exists(env_path):
                with open(env_path, 'r') as f:
                    content = f.read()
                    for line in content.splitlines():
                        if line.startswith("DEFAULT_PROVIDER="):
                            provider = line.split("=")[1].strip('"\'')
                        elif "API_KEY=" in line:
                            api_key = line.split("=")[1].strip('"\'')
                        elif line.startswith(f"{provider.upper()}_MODEL="):  # 讀取模型配置
                            model = line.split("=")[1].strip('"\'')
        
            if not api_key:
                error_msg = "未設定API金鑰，請前往設定頁面設定"
                if dialog:
                    dialog.show_error("API金鑰錯誤", error_msg)
                if hasattr(self, 'logger') and self.logger:
                    self.logger.error(error_msg)
                return error_msg

            # 設置默認模型（如果未在配置中指定）
            if not model:
                if provider == "gemini":
                    model = "gemini-1.5-flash"
                elif provider == "openai":
                    model = "gpt-3.5-turbo"
                elif provider == "claude":
                    model = "claude-3-opus-20240229"
                elif provider == "mistral":
                    model = "mistral-large-latest"
            
            # 根據不同的AI提供商建立API請求
            headers = {'Content-Type': 'application/json'}
            
            if provider == "gemini":
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
                data = {
                    "contents": [{"parts": [{"text": question}]}],
                    "generationConfig": {
                        "temperature": 0.7,
                        "maxOutputTokens": 2048
                    }
                }
            elif provider == "openai":
                url = "https://api.openai.com/v1/chat/completions"
                data = {
                    "model": model,
                    "messages": [{"role": "user", "content": question}],
                    "temperature": 0.7,
                    "max_tokens": 2048
                }
                headers['Authorization'] = f'Bearer {api_key}'
            elif provider == "claude":
                url = "https://api.anthropic.com/v1/messages"
                data = {
                    "model": model,
                    "max_tokens": 2048,
                    "messages": [{"role": "user", "content": question}],
                    "temperature": 0.7
                }
                headers.update({
                    'anthropic-version': '2023-06-01', 
                    'x-api-key': api_key
                })
            elif provider == "mistral":
                url = "https://api.mistral.ai/v1/chat/completions"
                data = {
                    "model": model,
                    "messages": [{"role": "user", "content": question}],
                    "temperature": 0.7,
                    "max_tokens": 2048
                }
                headers['Authorization'] = f'Bearer {api_key}'
            else:
                error_msg = f"不支援的AI提供商: {provider}"
                if dialog:
                    dialog.show_error("不支援的提供商", error_msg)
                if hasattr(self, 'logger') and self.logger:
                    self.logger.error(error_msg)
                return error_msg
                
            # 準備HTTP請求
            data_bytes = json.dumps(data).encode('utf-8')
            req = urllib.request.Request(url, data=data_bytes, headers=headers)
            
            # 增加超時處理
            try:
                with urllib.request.urlopen(req, timeout=30) as response:
                    result = json.loads(response.read().decode('utf-8'))
            except urllib.error.URLError as e:
                error_msg = f"API請求失敗: {str(e)}"
                if dialog:
                    dialog.show_error("API請求錯誤", error_msg)
                if hasattr(self, 'logger') and self.logger:
                    self.logger.error(error_msg)
                return error_msg
            except json.JSONDecodeError:
                error_msg = "解析API回應失敗"
                if dialog:
                    dialog.show_error("解析錯誤", error_msg)
                if hasattr(self, 'logger') and self.logger:
                    self.logger.error(error_msg)
                return error_msg
            
            # 根據不同的AI提供商解析回應
            response_text = ""
            token_info = None
            
            if provider == "gemini":
                if 'candidates' in result and result['candidates']:
                    response_text = result.get('candidates', [{}])[0].get('content', {}).get('parts', [{}])[0].get('text', '')
                    # 提取token使用情況
                    if 'usageMetadata' in result:
                        token_info = {
                            'prompt_tokens': result['usageMetadata'].get('promptTokenCount', 0),
                            'completion_tokens': result['usageMetadata'].get('candidatesTokenCount', 0),
                            'total_tokens': result['usageMetadata'].get('totalTokenCount', 0)
                        }
                else:
                    error_msg = f"Gemini API錯誤: {result.get('error', {}).get('message', '未知錯誤')}"
                    if dialog:
                        dialog.show_error("Gemini錯誤", error_msg)
                    if hasattr(self, 'logger') and self.logger:
                        self.logger.error(error_msg)
                    return error_msg
            elif provider == "openai":
                if 'choices' in result and result['choices']:
                    response_text = result.get('choices', [{}])[0].get('message', {}).get('content', '')
                    # 提取token使用情況
                    if 'usage' in result:
                        token_info = {
                            'prompt_tokens': result['usage'].get('prompt_tokens', 0),
                            'completion_tokens': result['usage'].get('completion_tokens', 0),
                            'total_tokens': result['usage'].get('total_tokens', 0)
                        }
                else:
                    error_msg = f"OpenAI API錯誤: {result.get('error', {}).get('message', '未知錯誤')}"
                    if dialog:
                        dialog.show_error("OpenAI錯誤", error_msg)
                    if hasattr(self, 'logger') and self.logger:
                        self.logger.error(error_msg)
                    return error_msg
            elif provider == "claude":
                if 'content' in result:
                    response_text = result.get('content', [{}])[0].get('text', '')
                    # 提取token使用情況
                    if 'usage' in result:
                        token_info = {
                            'prompt_tokens': result['usage'].get('input_tokens', 0),
                            'completion_tokens': result['usage'].get('output_tokens', 0),
                            'total_tokens': result['usage'].get('input_tokens', 0) + result['usage'].get('output_tokens', 0)
                        }
                else:
                    error_msg = f"Claude API錯誤: {result.get('error', {}).get('message', '未知錯誤')}"
                    if dialog:
                        dialog.show_error("Claude錯誤", error_msg)
                    if hasattr(self, 'logger') and self.logger:
                        self.logger.error(error_msg)
                    return error_msg
            elif provider == "mistral":
                if 'choices' in result and result['choices']:
                    response_text = result.get('choices', [{}])[0].get('message', {}).get('content', '')
                    # 提取token使用情況
                    if 'usage' in result:
                        token_info = {
                            'prompt_tokens': result['usage'].get('prompt_tokens', 0),
                            'completion_tokens': result['usage'].get('completion_tokens', 0),
                            'total_tokens': result['usage'].get('total_tokens', 0)
                        }
                else:
                    error_msg = f"Mistral API錯誤: {result.get('error', {}).get('message', '未知錯誤')}"
                    if dialog:
                        dialog.show_error("Mistral錯誤", error_msg)
                    if hasattr(self, 'logger') and self.logger:
                        self.logger.error(error_msg)
                    return error_msg            # 記錄API回應
            if hasattr(self, 'logger') and self.logger:
                log_response = response_text[:200] + "..." if len(response_text) > 200 else response_text
                self.logger.info(f"API 回應: {log_response}")
                
                if token_info:
                    self.logger.info(f"Token使用: {token_info}")  # 新增紀錄-當前token數
                    if 'completion_tokens' in token_info:
                        self.logger.info(f"當前token數: {token_info['completion_tokens']}")
                        # 保存當前的completion_tokens到previous_token
                        self.previous_token = token_info['completion_tokens']
                        
                        # 計算目標token數（使用previous_token而不是當前token）
                        if hasattr(self, 'previous_token') and hasattr(self, 'length_adjustment_factor') and self.length_adjustment_factor:
                            target_tokens = int(self.previous_token * self.length_adjustment_factor)
                            self.logger.info(f"目標token數: {target_tokens} = 上次保存的token數 {self.previous_token} × 長度調整因數 {self.length_adjustment_factor}")
                            
            # 存儲token信息並轉換為字符串
            if token_info:
                self.last_token_info = token_info
                self.token_info_str = str(token_info)  # 轉為字符串
                
            return response_text  # 只返回回應文本，不返回token信息
            
        except Exception as e:
            error_msg = f"API請求過程中發生錯誤: {str(e)}"
            if dialog:
                dialog.show_error("API錯誤", error_msg)
            if hasattr(self, 'logger') and self.logger:
                self.logger.error(error_msg)
            import traceback
            if hasattr(self, 'logger') and self.logger:
                self.logger.error(traceback.format_exc())
            return error_msg
