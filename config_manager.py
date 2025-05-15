import os
import json
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK


class ConfigManager:
    def __init__(self, ctx):
        self.ctx = ctx
        self.config = None
    
    def show_message(self, message, title="Information", message_type=INFOBOX):
        """顯示訊息對話框"""
        toolkit = self.ctx.ServiceManager.createInstance("com.sun.star.awt.Toolkit")
        parent = toolkit.getActiveTopWindow()
        mb = toolkit.createMessageBox(
            parent, message_type, BUTTONS_OK, title, str(message))
        mb.execute()
    
    def load_config(self):
        """
        從外部配置文件加載下拉選單配置
        """
        try:
            # 創建 .libreoffice 目錄（如果不存在）
            libreoffice_dir = os.path.join(os.path.expanduser("~"), ".libreoffice")
            if not os.path.exists(libreoffice_dir):
                os.makedirs(libreoffice_dir)
            
            # 嘗試從 .libreoffice 目錄加載配置文件
            config_file_path = os.path.join(libreoffice_dir, "libreoffice_ai_config.json")
            
            if os.path.exists(config_file_path):
                with open(config_file_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            else:
                # 如果找不到文件，使用默認配置
                self.config = {
                    "dropdowns": [
                        {
                            "id": "reading_level",
                            "display_name": "閱讀程度",
                            "position": 1,
                            "options": ["閱讀程度", "幼稚園", "國小", "初中", "高中", "大學", "研究所"],
                            "default_option": 0,
                            "prompt_template": "將文本調整為{option}閱讀程度，{prompt_value}。",
                            "prompt_values": {
                                "幼稚園": "使用非常簡單的詞彙和短句，適合5-6歲兒童理解",
                                "國小": "使用簡單的詞彙和句子結構，適合6-12歲兒童理解",
                                "初中": "使用中等難度的詞彙和句子，適合12-15歲青少年理解",
                                "高中": "使用較為複雜的詞彙和句子，適合15-18歲青少年理解",
                                "大學": "使用專業詞彙和複雜句式，適合大學生理解",
                                "研究所": "使用高度專業的詞彙和學術性表達，適合研究生水平"
                            }
                        },
                        {
                            "id": "length_adjustment",
                            "display_name": "長度調整",
                            "position": 2,
                            "options": ["長度調整", "-75%", "-50%", "-25%", "+25%", "+50%", "+75%"],
                            "default_option": 0,
                            "prompt_templates": {
                                "decrease": "將文本精簡，縮減{percentage}的內容，但保留核心信息。",
                                "increase": "將文本擴展{percentage}，添加更多細節和解釋。"
                            }
                        },
                        {
                            "id": "language",
                            "display_name": "語言",
                            "position": 3,
                            "options": ["語言", "繁體中文", "英文"],
                            "default_option": 0,
                            "prompt_template": "將文本轉換為{option}。",
                            "target_option": "英文"  # 只有選擇這個選項時才生成提示詞
                        },
                        {
                            "id": "emotion",
                            "display_name": "情緒",
                            "position": 4,
                            "options": ["情緒", "活潑", "穩重", "成熟"],
                            "default_option": 0,
                            "prompt_template": "以{option}的情緒風格撰寫，{prompt_value}。",
                            "prompt_values": {
                                "活潑": "使用生動、活潑的語調，加入更多感嘆詞和生動的形容詞，表現出輕鬆愉快的態度",
                                "穩重": "使用平穩、客觀的語調，避免過度誇張的表達，保持邏輯清晰和事實導向",
                                "成熟": "使用成熟、深思熟慮的語調，加入適當的專業術語和深度分析，表現出沉穩和專業"
                            }
                        }
                    ],
                    "prompt_header": "請按照以下要求修改文本：",
                    "original_text_label": "\n原始文本：",
                    "modified_text_label": "\n修改後的文本："
                }
                
                # 將默認配置寫入文件作為範例
                with open(config_file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.config, f, ensure_ascii=False, indent=2)
                    
                self.show_message(f"已在 {config_file_path} 建立默認配置檔案", "信息")
            return self.config
        except Exception as e:
            self.show_message(f"載入配置檔案時出錯: {str(e)}\n將使用內建默認配置", "警告", MESSAGEBOX)
            return None
    
    def reload_configuration(self):
        """
        重新載入配置文件
        """
        try:
            old_config = self.config
            self.config = None  # 清除舊配置
            self.load_config()  # 重新載入
            return True
        except Exception as e:
            self.config = old_config  # 恢復舊配置
            self.show_message(f"重新載入配置失敗: {str(e)}", "錯誤", MESSAGEBOX)
            return False
    
    def save_env_file(self, provider, api_key):
        """
        儲存設定至 .env 檔案並創建啟動腳本（Windows 與 Linux）
        """
        try:
            # 創建 .libreoffice 目錄（如果不存在）
            libreoffice_dir = os.path.join(os.path.expanduser("~"), ".libreoffice")
            if not os.path.exists(libreoffice_dir):
                os.makedirs(libreoffice_dir)
        
            # 定義檔案路徑
            env_path = os.path.join(libreoffice_dir, ".env")
        
            provider_map = {
                "Gemini": "gemini",
                "GPT (OpenAI)": "openai",
                "Claude": "claude"
            }
    
            provider_key = provider_map.get(provider, "gemini")
    
            # 組織 .env 檔案內容
            env_content = f"DEFAULT_PROVIDER={provider_key}\n"
        
            # 針對不同供應商設置對應的 API 金鑰
            if provider_key == "gemini":
                env_content += f"GOOGLE_API_KEY={api_key}\n"
            elif provider_key == "openai":
                env_content += f"OPENAI_API_KEY={api_key}\n"
            elif provider_key == "claude":
                env_content += f"ANTHROPIC_API_KEY={api_key}\n"
    
            # 寫入 .env 檔案
            with open(env_path, "w") as f:
                f.write(env_content)
        
            # 依據作業系統創建啟動腳本
            if os.name == 'nt':  # Windows
                batch_path = os.path.join(libreoffice_dir, "ai_service.bat")
                batch_content = """@echo off
pip install -r requirements.txt
python gemini_service.py
pause"""
                with open(batch_path, "w") as f:
                    f.write(batch_content)
            else:  # Linux/Mac
                shell_path = os.path.join(libreoffice_dir, "ai_service.sh")
                shell_content = """#!/bin/bash
pip install -r requirements.txt
python gemini_service.py
read -p "Press Enter to continue..."
"""
                with open(shell_path, "w") as f:
                    f.write(shell_content)
            
                # 設置 shell 腳本的執行權限
                os.chmod(shell_path, 0o755)
        
            return True
        except Exception as e:
            self.show_message(f"儲存設定失敗: {str(e)}", "錯誤", MESSAGEBOX)
            return False
            
    def generate_adjustment_prompt(self, selected_options, text=""):
        """
        根據獲取的選項映射，使用配置文件中的模板生成提示詞
        
        Args:
            selected_options: 字典，格式為 {'dropdown_id': 'selected_value'}
            text: 要調整的原始文本
            
        Returns:
            生成的完整提示詞
        """
        # 檢查配置是否已加載
        if self.config is None:
            self.load_config()
            
        # 從配置獲取提示詞頭部和標籤
        prompt_parts = [self.config.get("prompt_header", "請按照以下要求修改文本：")]
        
        # 對於每個下拉選單配置
        for dropdown in self.config["dropdowns"]:
            dropdown_id = dropdown["id"]
            
            # 獲取選擇的選項
            selected_option = selected_options.get(dropdown_id)
            
            # 如果選擇了非默認選項
            if selected_option and selected_option != dropdown["options"][dropdown["default_option"]]:
                # 特殊處理長度調整
                if dropdown_id == "length_adjustment":
                    if "-" in selected_option:
                        percentage = selected_option.replace("-", "")
                        template = dropdown["prompt_templates"]["decrease"]
                        prompt_parts.append(template.format(percentage=percentage))
                    else:
                        percentage = selected_option.replace("+", "")
                        template = dropdown["prompt_templates"]["increase"]
                        prompt_parts.append(template.format(percentage=percentage))
                
                # 特殊處理只針對特定選項生成提示詞
                elif "target_option" in dropdown and selected_option == dropdown["target_option"]:
                    template = dropdown.get("prompt_template", "將文本轉換為{option}。")
                    prompt_parts.append(template.format(option=selected_option))
                
                # 標準處理具有提示值的下拉選單
                elif "prompt_values" in dropdown and selected_option in dropdown["prompt_values"]:
                    template = dropdown.get("prompt_template", "")
                    if template:
                        prompt_value = dropdown["prompt_values"][selected_option]
                        prompt_parts.append(template.format(option=selected_option, prompt_value=prompt_value))
                
                # 一般處理不需要額外提示值的下拉選單
                elif "prompt_template" in dropdown:
                    template = dropdown["prompt_template"]
                    prompt_parts.append(template.format(option=selected_option))
        
        # 添加原始文本和修改後的文本提示
        prompt_parts.append(self.config.get("original_text_label", "\n原始文本："))
        prompt_parts.append(text)
        prompt_parts.append(self.config.get("modified_text_label", "\n修改後的文本："))
        
        # 組合最終的提示詞，使用換行符連接非空部分
        return "\n".join(prompt_parts)
