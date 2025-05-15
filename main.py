import uno
import unohelper
import urllib.request
import json
import officehelper
import sys
import os
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import XActionListener
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK

sys.path.append(os.path.dirname(__file__))

# Implement the UNO component
class AIQueryJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

        # 初始化配置
        self.config = None

        try:
            self.sm = ctx.getServiceManager()
            self.desktop = XSCRIPTCONTEXT.getDesktop()
        except NameError:
            self.sm = ctx.ServiceManager
            self.desktop = self.ctx.getServiceManager().createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx
            )
    
    def trigger(self, args):
        try:
            # 不管傳入什麼參數，都執行 AI 查詢
            self.main()
        except Exception as e:
            from utils import Utils
            utils = Utils(self.ctx)
            utils.show_message(f"Error: {str(e)}", "Error", MESSAGEBOX)
            
    def main(self, *args):
        try:
            # 導入模組
            from config_manager import ConfigManager
            from ai_service import AIService
            from utils import Utils
            from dialog_builder import DialogBuilder
            from event_handlers import EventHandlers

            # 初始化各個服務
            utils = Utils(self.ctx)
            config_manager = ConfigManager(self.ctx)
            ai_service = AIService(self.ctx)
            dialog_builder = DialogBuilder(self.ctx)
            
            # 確保配置已加載
            config = config_manager.load_config()

            # 在創建對話框前先獲取選取的文字
            selected_text = utils.get_selected_text()

            # Create dialog
            dialog = dialog_builder.create_simple_dialog(config)
            text_field = dialog.getControl("TextField1")
            response_field = dialog.getControl("ResponseField")
            prompts_field = dialog.getControl("PromptsField")

            current_response = [""]

            # 如果有選取的文字，設置到輸入框
            if selected_text:
                # 使用 Model 來設置文字
                text_field.getModel().Text = selected_text

                # 將游標移到文字末尾
                text_field.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", len(selected_text), len(selected_text)))

            # 創建事件處理器
            event_handler = EventHandlers(self.ctx, ai_service, config_manager, utils)
            event_handler.reload_requested = False
            
            # 獲取所有對話框監聽器
            listeners = event_handler.get_dialog_listeners(dialog, current_response)
            
            # 綁定按鈕事件
            dialog.getControl("AskButton").addActionListener(listeners["AskButtonListener"])
            dialog.getControl("InsertButton").addActionListener(listeners["InsertButtonListener"])
            dialog.getControl("ClearButton").addActionListener(listeners["ClearButtonListener"])
            dialog.getControl("CloseButton").addActionListener(listeners["CloseButtonListener"])
            dialog.getControl("ReloadConfigButton").addActionListener(listeners["ReloadConfigButtonListener"])
            dialog.getControl("ResetDropdownsButton").addActionListener(listeners["ResetDropdownsButtonListener"])
            dialog.getControl("PreviewPromptsButton").addActionListener(listeners["PreviewPromptsButtonListener"])
            dialog.getControl("AdjustResponseButton").addActionListener(listeners["AdjustResponseButtonListener"])
            dialog.getControl("SettingsButton").addActionListener(listeners["SettingsButtonListener"])
            
            # Execute dialog
            dialog.execute()
            
            # 檢查是否需要重新載入
            if hasattr(event_handler, 'reload_requested') and event_handler.reload_requested:
                # 重新執行主函數來重新創建對話框
                self.main()
            
        except Exception as e:
            from utils import Utils
            utils = Utils(self.ctx)
            utils.show_message(f"Error: {str(e)}", "Error", MESSAGEBOX)

# Starting from Python IDE
def main():
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
    except NameError:
        ctx = officehelper.bootstrap()
        if ctx is None:
            print("ERROR: Could not bootstrap default Office.")
            sys.exit(1)
    
    job = AIQueryJob(ctx)
    job.trigger("StartAIQuery")  # 使用與 XML 中相同的命令

# Starting from command line
if __name__ == "__main__":
    main()

# 註冊 UNO 組件
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    AIQueryJob,
    "org.openoffice.comp.pyuno.AIQueryJob",  # 確保與 XML 完全匹配
    ("com.sun.star.task.Job",),
)