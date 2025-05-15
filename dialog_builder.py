import uno
import unohelper


class DialogBuilder:
    def __init__(self, ctx):
        self.ctx = ctx
        
    def create_settings_dialog(self):
        """創建設定對話框"""
        # Get the component context
        smgr = self.ctx.getServiceManager()
        
        # Create dialog
        dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", self.ctx)
        dialog_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", self.ctx)
        dialog.setModel(dialog_model)
        
        # Set dialog properties
        dialog_model.setPropertyValues(
            ("Width", "Height", "Title"),
            (300, 160, " AI Settings")
        )
        
        # Model label
        model_label = dialog.getModel().createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        model_label.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label"),
            (80, 15, 20, 23, "AI Model:")
        )
        dialog_model.insertByName("ModelLabel", model_label)
        
        # Model dropdown
        model_dropdown = dialog.getModel().createInstance("com.sun.star.awt.UnoControlListBoxModel")
        model_dropdown.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Dropdown"),
            (180, 15, 100, 20, True)
        )
        model_dropdown.StringItemList = tuple(["Gemini", "GPT (OpenAI)", "Claude"])
        model_dropdown.SelectedItems = [0]  # Default to Gemini
        dialog_model.insertByName("ModelDropdown", model_dropdown)
        
        # API Key label
        api_key_label = dialog.getModel().createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        api_key_label.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label"),
            (80, 15, 20, 63, "API Key:")
        )
        dialog_model.insertByName("ApiKeyLabel", api_key_label)
        
        # API Key input field
        api_key_field = dialog.getModel().createInstance("com.sun.star.awt.UnoControlEditModel")
        api_key_field.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "EchoChar"),
            (180, 15, 100, 60, 42)  # EchoChar 42 = "*" for password masking
        )
        dialog_model.insertByName("ApiKeyField", api_key_field)
        
        # Help text
        help_text = dialog.getModel().createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        help_text.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label", "MultiLine"),
            (280, 35, 10, 90, "設定會儲存至 ~/.libreoffice 目錄。\n按下儲存後將創建啟動 AI 服務的批次檔。", True)
        )
        dialog_model.insertByName("HelpText", help_text)
        
        # Save button
        save_button = dialog.getModel().createInstance("com.sun.star.awt.UnoControlButtonModel")
        save_button.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label"),
            (80, 20, 110, 125, "Save")
        )
        dialog_model.insertByName("SaveButton", save_button)
        
        # Cancel button
        cancel_button = dialog.getModel().createInstance("com.sun.star.awt.UnoControlButtonModel")
        cancel_button.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label"),
            (80, 20, 200, 125, "Cancel")
        )
        dialog_model.insertByName("CancelButton", cancel_button)
        
        # Create window peer
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        dialog.createPeer(toolkit, None)
        
        # Try to load existing .env settings
        try:
            import os
            libreoffice_dir = os.path.join(os.path.expanduser("~"), ".libreoffice")
            env_path = os.path.join(libreoffice_dir, ".env")
            
            if os.path.exists(env_path):
                with open(env_path, 'r') as f:
                    content = f.read()
                    # Find key-value pairs
                    for line in content.splitlines():
                        if line.startswith("DEFAULT_PROVIDER="):
                            provider = line.split("=")[1].strip('"\'')
                            if provider == "gemini":
                                model_dropdown.SelectedItems = [0]
                            elif provider == "openai":
                                model_dropdown.SelectedItems = [1]
                            elif provider == "claude":
                                model_dropdown.SelectedItems = [2]
                        elif "API_KEY=" in line:
                            api_key = line.split("=")[1].strip('"\'')
                            if api_key:
                                api_key_field.Text = api_key
        except Exception as e:
            print(f"Error loading .env: {str(e)}")
            
        return dialog

    def create_simple_dialog(self, config):
        """創建主要對話框"""
        # Get the component context
        smgr = self.ctx.getServiceManager()
        
        # Create dialog
        dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", self.ctx)
        dialog_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", self.ctx)
        dialog.setModel(dialog_model)
        
        # Set dialog properties
        dialog_model.setPropertyValues(
            ("Width", "Height", "Title"),
            (350, 360, " AI Query")
        )
        
        # Question label
        question_label = dialog.getModel().createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        question_label.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label"),
            (170, 15, 10, 10, "Your question:")
        )
        dialog_model.insertByName("QuestionLabel", question_label)
        
        # Question input field
        text_field = dialog.getModel().createInstance("com.sun.star.awt.UnoControlEditModel")
        text_field.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "MultiLine", "VScroll"),
            (330, 50, 10, 30, True, True)
        )
        dialog_model.insertByName("TextField1", text_field)
        
        # Response label
        response_label = dialog.getModel().createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        response_label.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label"),
            (170, 15, 10, 100, "AI Response:")
        )
        dialog_model.insertByName("ResponseLabel", response_label)
        
        # Response display area
        response_field = dialog.getModel().createInstance("com.sun.star.awt.UnoControlEditModel")
        response_field.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "MultiLine", "ReadOnly", "VScroll"),
            (330, 90, 10, 120, True, True, True)
        )
        dialog_model.insertByName("ResponseField", response_field)

        # Adjust Response label
        adjust_response_label = dialog.getModel().createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        adjust_response_label.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label"),
            (55, 10, 10, 228, "Adjust Response:")
        )
        dialog_model.insertByName("AdjustResponseLabel", adjust_response_label)

        # 根據配置排序下拉選單
        sorted_dropdowns = sorted(config["dropdowns"], key=lambda x: x["position"])
        
        # 定義下拉選單的位置
        dropdown_positions = [70, 120, 170, 220]
        
        # 動態創建所有的下拉選單
        for i, dropdown in enumerate(sorted_dropdowns):
            if i < len(dropdown_positions):  # 確保不超過預定義的位置數量
                list_box = dialog.getModel().createInstance("com.sun.star.awt.UnoControlListBoxModel")
                list_box.setPropertyValues(
                    ("Width", "Height", "PositionX", "PositionY", "Dropdown"),
                    (40, 15, dropdown_positions[i], 225, True)
                )
                # 設置選項
                list_box.StringItemList = tuple(dropdown["options"])
                # 設置默認選項
                list_box.SelectedItems = [dropdown["default_option"]]
                # 插入到對話框模型
                dialog_model.insertByName(f"{dropdown['id']}List", list_box)

        # Settings Button - 設定按鈕
        settings_button = dialog.getModel().createInstance("com.sun.star.awt.UnoControlButtonModel")
        settings_button.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label", "HelpText"),
            (10, 10, 330, 335, "⚙️", "AI 模型設定")
        )
        dialog_model.insertByName("SettingsButton", settings_button)

        # Reload Config Button - 重載配置按鈕
        reload_config_button = dialog.getModel().createInstance("com.sun.star.awt.UnoControlButtonModel")
        reload_config_button.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label", "HelpText"),
            (10, 10, 270, 220, "♻️", "重載配置選單")
        )
        dialog_model.insertByName("ReloadConfigButton", reload_config_button)

        # Reset Dropdowns Button - 清除下拉選單按鈕
        reset_dropdowns_button = dialog.getModel().createInstance("com.sun.star.awt.UnoControlButtonModel")
        reset_dropdowns_button.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label", "HelpText"),
            (10, 10, 270, 235, "🧹", "還原下拉選單")
        )
        dialog_model.insertByName("ResetDropdownsButton", reset_dropdowns_button)

        # Adjust Response Button
        adjust_response_button = dialog.getModel().createInstance("com.sun.star.awt.UnoControlButtonModel")
        adjust_response_button.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label", "HelpText"),
            (50, 15, 290, 225, "Adjust", "調整 AI 回應")
        )
        dialog_model.insertByName("AdjustResponseButton", adjust_response_button)

        # Adjust Prompts display area
        prompts_field = dialog.getModel().createInstance("com.sun.star.awt.UnoControlEditModel")
        prompts_field.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "MultiLine", "VScroll"),
            (270, 50, 10, 255, True, True)
        )
        dialog_model.insertByName("PromptsField", prompts_field)

        # Preview Prompts Button
        preview_prompts_button = dialog.getModel().createInstance("com.sun.star.awt.UnoControlButtonModel")
        preview_prompts_button.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label", "HelpText"),
            (50, 15, 290, 273, "Preview", "預覽調整提示詞")
        )
        dialog_model.insertByName("PreviewPromptsButton", preview_prompts_button)

        # Ask button
        ask_button = dialog.getModel().createInstance("com.sun.star.awt.UnoControlButtonModel")
        ask_button.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label"),
            (60, 20, 10, 330, "Ask AI")
        )
        dialog_model.insertByName("AskButton", ask_button)
        
        # Insert button
        insert_button = dialog.getModel().createInstance("com.sun.star.awt.UnoControlButtonModel")
        insert_button.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label"),
            (80, 20, 80, 330, "Insert to Doc")
        )
        dialog_model.insertByName("InsertButton", insert_button)
        
        # Clear Response button
        clear_button = dialog.getModel().createInstance("com.sun.star.awt.UnoControlButtonModel")
        clear_button.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label"),
            (80, 20, 170, 330, "Clear Response")
        )
        dialog_model.insertByName("ClearButton", clear_button)
        
        # Close button
        close_button = dialog.getModel().createInstance("com.sun.star.awt.UnoControlButtonModel")
        close_button.setPropertyValues(
            ("Width", "Height", "PositionX", "PositionY", "Label"),
            (60, 20, 260, 330, "Close")
        )
        dialog_model.insertByName("CloseButton", close_button)
        
        # Create window peer
        toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        dialog.createPeer(toolkit, None)
        
        return dialog
