import unohelper
from com.sun.star.awt import XActionListener
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX
import uno


class EventHandlers:
    def __init__(self, ctx, ai_service, config_manager, utils):
        self.ctx = ctx
        self.ai_service = ai_service
        self.config_manager = config_manager
        self.utils = utils
    
    def get_dialog_listeners(self, dialog, current_response):
        """
        獲取所有對話框按鈕的監聽器
        
        Args:
            dialog: 對話框實例
            current_response: 用於存儲當前回應的列表
        
        Returns:
            具有所有監聽器的字典
        """
        listeners = {
            "AskButtonListener": self.create_ask_button_listener(dialog, current_response),
            "InsertButtonListener": self.create_insert_button_listener(current_response),
            "ClearButtonListener": self.create_clear_button_listener(dialog),
            "CloseButtonListener": self.create_close_button_listener(dialog),
            "ReloadConfigButtonListener": self.create_reload_config_button_listener(dialog),
            "ResetDropdownsButtonListener": self.create_reset_dropdowns_button_listener(dialog),
            "PreviewPromptsButtonListener": self.create_preview_prompts_button_listener(dialog),
            "AdjustResponseButtonListener": self.create_adjust_response_button_listener(dialog, current_response),
            "SettingsButtonListener": self.create_settings_button_listener(dialog)
        }
        
        return listeners
    
    def create_ask_button_listener(self, dialog, current_response):
        """創建詢問按鈕監聽器"""
        
        class AskButtonListener(unohelper.Base, XActionListener):
            def __init__(self, parent, dialog, current_response, ai_service, utils):
                self.parent = parent
                self.dialog = dialog
                self.current_response = current_response
                self.ai_service = ai_service
                self.utils = utils
            
            def actionPerformed(self, event):
                try:
                    text_field = self.dialog.getControl("TextField1")
                    response_field = self.dialog.getControl("ResponseField")
                    
                    question = text_field.getText()
                    if question.strip():
                        response = self.ai_service.ask_ai(question)
                        current_text = response_field.getText()
                        if current_text.strip():
                            new_text = current_text + "\n\n-------------------\n\n" + response
                        else:
                            new_text = response
                        response_field.setText(new_text)
                        self.current_response[0] = new_text  # 更新列表
                    else:
                        self.utils.show_message("Please enter a question", "Warning", MESSAGEBOX)
                except Exception as e:
                    self.utils.show_message(f"Error: {str(e)}", "Error", MESSAGEBOX)
            
            def disposing(self, event):
                pass
                
        return AskButtonListener(self, dialog, current_response, self.ai_service, self.utils)
        
    def create_insert_button_listener(self, current_response):
        """創建插入按鈕監聽器"""
        
        class InsertButtonListener(unohelper.Base, XActionListener):
            def __init__(self, parent, current_response, utils):
                self.parent = parent
                self.current_response = current_response
                self.utils = utils
            
            def actionPerformed(self, event):
                if self.current_response[0]:
                    self.utils.insert_text_at_cursor(self.current_response[0])
                else:
                    self.utils.show_message("No response to insert", "Warning", MESSAGEBOX)
            
            def disposing(self, event):
                pass
                
        return InsertButtonListener(self, current_response, self.utils)
    
    def create_clear_button_listener(self, dialog):
        """創建清除按鈕監聽器"""
        
        class ClearButtonListener(unohelper.Base, XActionListener):
            def __init__(self, dialog, current_response):
                self.dialog = dialog
                self.current_response = current_response
                
            def actionPerformed(self, event):
                response_field = self.dialog.getControl("ResponseField")
                response_field.setText("")
                self.current_response[0] = ""
            
            def disposing(self, event):
                pass
        
        return ClearButtonListener(dialog, [])
    
    def create_close_button_listener(self, dialog):
        """創建關閉按鈕監聽器"""
        
        class CloseButtonListener(unohelper.Base, XActionListener):
            def __init__(self, dialog):
                self.dialog = dialog
                
            def actionPerformed(self, event):
                self.dialog.endExecute()
            
            def disposing(self, event):
                pass
        
        return CloseButtonListener(dialog)
    
    def create_reload_config_button_listener(self, dialog):
        """創建重載配置按鈕監聽器"""
        
        class ReloadConfigButtonListener(unohelper.Base, XActionListener):
            def __init__(self, parent, dialog, config_manager, utils):
                self.parent = parent
                self.dialog = dialog
                self.config_manager = config_manager
                self.utils = utils
                
            def actionPerformed(self, event):
                try:
                    # 重新載入配置
                    success = self.config_manager.reload_configuration()
                    if success:
                        self.utils.show_message("配置已重新載入", "成功", INFOBOX)
                        
                        # 關閉當前對話框
                        self.dialog.endExecute()
                        
                        # 通知外部重新開啟主對話框
                        self.parent.reload_requested = True
                    else:
                        self.utils.show_message("配置重載失敗", "錯誤", MESSAGEBOX)
                except Exception as e:
                    self.utils.show_message(f"重載配置時發生錯誤: {str(e)}", "錯誤", MESSAGEBOX)
                    
            def disposing(self, event):
                pass
                
        return ReloadConfigButtonListener(self, dialog, self.config_manager, self.utils)
    
    def create_reset_dropdowns_button_listener(self, dialog):
        """創建重置下拉選單按鈕監聽器"""
        
        class ResetDropdownsButtonListener(unohelper.Base, XActionListener):
            def __init__(self, parent, dialog, config_manager, utils):
                self.parent = parent
                self.dialog = dialog
                self.config_manager = config_manager
                self.utils = utils
                
            def actionPerformed(self, event):
                try:
                    # 遍歷配置中的所有下拉選單
                    for dropdown in self.config_manager.config["dropdowns"]:
                        dropdown_id = dropdown["id"]
                        default_option = dropdown["default_option"]
                        
                        try:
                            # 獲取下拉選單控制項
                            dropdown_control = self.dialog.getControl(f"{dropdown_id}List")
                            # 重置為默認選項
                            dropdown_control.selectItemPos(default_option, True)
                        except Exception as e:
                            print(f"無法重置 {dropdown_id} 的選擇: {str(e)}")
                    
                    self.utils.show_message("所有下拉選單已重置為默認設置", "成功", INFOBOX)
                except Exception as e:
                    self.utils.show_message(f"重置下拉選單時發生錯誤: {str(e)}", "錯誤", MESSAGEBOX)
                    
            def disposing(self, event):
                pass
                
        return ResetDropdownsButtonListener(self, dialog, self.config_manager, self.utils)
    
    def create_preview_prompts_button_listener(self, dialog):
        """創建預覽提示詞按鈕監聽器"""
        
        class PreviewPromptsButtonListener(unohelper.Base, XActionListener):
            def __init__(self, parent, dialog, config_manager, utils):
                self.parent = parent
                self.dialog = dialog
                self.config_manager = config_manager
                self.utils = utils
                
            def actionPerformed(self, event):
                try:
                    # 獲取每個下拉選單的選擇
                    selected_options = {}
                    
                    # 遍歷配置中的所有下拉選單
                    for dropdown in self.config_manager.config["dropdowns"]:
                        dropdown_id = dropdown["id"]
                        try:
                            # 獲取選擇的選項索引
                            dropdown_control = self.dialog.getControl(f"{dropdown_id}List")
                            selected_index = dropdown_control.getSelectedItemPos()
                            # 獲取選擇的選項值
                            selected_value = dropdown_control.getItem(selected_index)
                            # 添加到選項字典
                            selected_options[dropdown_id] = selected_value
                        except Exception as e:
                            print(f"無法獲取 {dropdown_id} 的選擇: {str(e)}")
                            
                    # 獲取當前的回應文本
                    response_field = self.dialog.getControl("ResponseField")
                    current_text = response_field.getText().strip()
                    
                    # 生成包含當前文本的提示詞
                    prompt_template = self.config_manager.generate_adjustment_prompt(
                        selected_options=selected_options,
                        text=current_text  # 加入當前回應文本
                    )
                    
                    # 顯示生成的提示詞到提示詞欄位
                    prompts_field = self.dialog.getControl("PromptsField")
                    prompts_field.setText(prompt_template.strip())
                    
                except Exception as e:
                    self.utils.show_message(f"預覽提示詞錯誤: {str(e)}", "錯誤", MESSAGEBOX)
                    
            def disposing(self, event):
                pass
                
        return PreviewPromptsButtonListener(self, dialog, self.config_manager, self.utils)
    
    def create_adjust_response_button_listener(self, dialog, current_response):
        """創建調整回應按鈕監聽器"""
        
        class AdjustResponseButtonListener(unohelper.Base, XActionListener):
            def __init__(self, parent, dialog, current_response, config_manager, ai_service, utils):
                self.parent = parent
                self.dialog = dialog
                self.current_response = current_response
                self.config_manager = config_manager
                self.ai_service = ai_service
                self.utils = utils
                
            def actionPerformed(self, event):
                try:
                    prompts_field = self.dialog.getControl("PromptsField")
                    response_field = self.dialog.getControl("ResponseField")
                    prompts_text = prompts_field.getText().strip()
                    
                    # 初始化長度調整參數
                    length_adjustment = None
                    
                    # 判斷是否有手動編輯的提示詞
                    if prompts_text:
                        # 使用提示詞欄位的內容
                        complete_prompt = prompts_text
                        # 從手動提示詞中提取長度調整指示
                        length_adjustment = self.ai_service.extract_length_adjustment(complete_prompt)
                    else:
                        # 如果提示詞欄位為空，則自動生成提示詞
                        selected_options = {}
                        
                        # 遍歷配置中的所有下拉選單
                        for dropdown in self.config_manager.config["dropdowns"]:
                            dropdown_id = dropdown["id"]
                            try:
                                dropdown_control = self.dialog.getControl(f"{dropdown_id}List")
                                selected_index = dropdown_control.getSelectedItemPos()
                                selected_value = dropdown_control.getItem(selected_index)
                                selected_options[dropdown_id] = selected_value
                                
                                # 檢查是否有長度調整選項
                                if dropdown_id == "length_adjustment" and selected_value != dropdown["options"][dropdown["default_option"]]:
                                    length_adjustment = selected_value
                            except Exception as e:
                                print(f"無法獲取 {dropdown_id} 的選擇: {str(e)}")
                                
                        # 獲取當前的回應文本
                        current_text = response_field.getText().strip()
                        
                        # 生成包含當前文本的提示詞
                        complete_prompt = self.config_manager.generate_adjustment_prompt(
                            selected_options=selected_options,
                            text=current_text
                        )
                        
                    # 使用新的長度調整功能發送請求
                    if length_adjustment:
                        # 使用帶長度調整的高級方法
                        adjusted_response = self.ai_service.ask_ai_with_length_adjustment(
                            question=complete_prompt,
                            length_adjustment=length_adjustment,
                            max_attempts=3
                        )
                    else:
                        # 使用標準方法（不帶長度調整）
                        adjusted_response = self.ai_service.ask_ai(question=complete_prompt)
                    
                    # 更新回應欄位
                    response_field.setText(adjusted_response)
                    
                    # 更新 current_response 列表的第一個元素
                    self.current_response[0] = adjusted_response
                    
                except Exception as e:
                    self.utils.show_message(f"調整回應錯誤: {str(e)}", "錯誤", MESSAGEBOX)
                    
            def disposing(self, event):
                pass
                
        return AdjustResponseButtonListener(self, dialog, current_response, self.config_manager, self.ai_service, self.utils)
    
    def create_settings_button_listener(self, dialog):
        """創建設定按鈕監聽器"""
        
        class SettingsButtonListener(unohelper.Base, XActionListener):
            def __init__(self, parent, dialog, ctx, config_manager, ai_service, utils):
                self.parent = parent
                self.main_dialog = dialog
                self.ctx = ctx
                self.config_manager = config_manager
                self.ai_service = ai_service
                self.utils = utils
                
            def actionPerformed(self, event):
                try:
                    # 引入對話框建立器
                    from dialog_builder import DialogBuilder
                    
                    # 顯示設置對話框
                    dialog_builder = DialogBuilder(self.ctx)
                    settings_dialog = dialog_builder.create_settings_dialog()
                    
                    # 建立設置對話框按鈕監聽器
                    settings_listeners = self.parent.get_settings_dialog_listeners(settings_dialog)
                    
                    # 添加按鈕事件處理
                    settings_dialog.getControl("SaveButton").addActionListener(settings_listeners["SaveSettingsListener"])
                    settings_dialog.getControl("CancelButton").addActionListener(settings_listeners["CancelSettingsListener"])
                    
                    # 顯示設置對話框
                    settings_dialog.execute()
                except Exception as e:
                    self.utils.show_message(f"打開設置對話框時出錯: {str(e)}", "錯誤", MESSAGEBOX)
                    
            def disposing(self, event):
                pass
                
        return SettingsButtonListener(self, dialog, self.ctx, self.config_manager, self.ai_service, self.utils)
    
    def get_settings_dialog_listeners(self, settings_dialog):
        """獲取設定對話框的事件監聽器"""
        
        class SaveSettingsListener(unohelper.Base, XActionListener):
            def __init__(self, parent, dialog, config_manager, ai_service, utils):
                self.parent = parent
                self.dialog = dialog
                self.config_manager = config_manager
                self.ai_service = ai_service
                self.utils = utils
                
            def actionPerformed(self, event):
                try:
                    # 獲取所選模型和API密鑰
                    model_list = self.dialog.getControl("ModelDropdown")
                    model_index = model_list.getSelectedItemPos()
                    model_name = model_list.getItem(model_index)
                    api_key = self.dialog.getControl("ApiKeyField").getText()
                    
                    # 顯示驗證中的訊息
                    self.utils.show_message("正在驗證API金鑰，請稍候...", "驗證中")
                    
                    # 驗證API金鑰有效性
                    is_valid = self.ai_service.validate_api_key(model_name, api_key)
                    
                    if is_valid:
                        # 保存設置
                        if self.config_manager.save_env_file(model_name, api_key):
                            self.utils.show_message("API金鑰驗證成功，設置已保存", "成功")
                            self.dialog.endExecute()
                        else:
                            self.utils.show_message("保存設置失敗", "錯誤", MESSAGEBOX)
                    else:
                        self.utils.show_message("API金鑰無效，請檢查後重試", "錯誤", MESSAGEBOX)
                except Exception as e:
                    self.utils.show_message(f"保存設置時出錯: {str(e)}", "錯誤", MESSAGEBOX)
                    
            def disposing(self, event):
                pass
                
        class CancelSettingsListener(unohelper.Base, XActionListener):
            def __init__(self, dialog):
                self.dialog = dialog
                
            def actionPerformed(self, event):
                self.dialog.endExecute()
                
            def disposing(self, event):
                pass
                
        return {
            "SaveSettingsListener": SaveSettingsListener(self, settings_dialog, self.config_manager, self.ai_service, self.utils),
            "CancelSettingsListener": CancelSettingsListener(settings_dialog)
        }
