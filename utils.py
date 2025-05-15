class Utils:
    def __init__(self, ctx):
        self.ctx = ctx

    def show_message(self, message, title="Information", message_type=None):
        """顯示訊息對話框"""
        # Import here to avoid circular imports
        from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
        
        # Use default message type if not specified
        if message_type is None:
            message_type = INFOBOX
            
        toolkit = self.ctx.ServiceManager.createInstance("com.sun.star.awt.Toolkit")
        parent = toolkit.getActiveTopWindow()
        mb = toolkit.createMessageBox(
            parent, message_type, BUTTONS_OK, title, str(message))
        mb.execute()
        
    def get_selected_text(self):
        """獲取選中的文字"""
        try:
            # 獲取當前文件
            desktop = self.ctx.ServiceManager.createInstance("com.sun.star.frame.Desktop")
            doc = desktop.getCurrentComponent()
        
            # 獲取選取內容
            selection = doc.getCurrentController().getSelection()
        
            # 如果有選取內容
            if selection and selection.getCount() > 0:
                # 獲取第一個選取區域
                selected_range = selection.getByIndex(0)
            
                # 獲取選取的文字
                if hasattr(selected_range, 'getString'):
                    return selected_range.getString().strip()
            
                # 如果是表格選取，嘗試獲取儲存格文字
                elif hasattr(selected_range, 'getFormula'):
                    return selected_range.getFormula().strip()
                
            return ""
        except Exception as e:
            print(f"Error getting selected text: {str(e)}")  # 用於除錯
            return ""
            
    def insert_text_at_cursor(self, text):
        """在游標位置插入文字"""
        desktop = self.ctx.ServiceManager.createInstance("com.sun.star.frame.Desktop")
        doc = desktop.getCurrentComponent()
        cursor = doc.getCurrentController().getViewCursor()
    
        # 首先移動到選取區域的結束位置
        if cursor.isCollapsed() == False:  # 如果有選取文字
            cursor.gotoRange(cursor.getEnd(), False)  # 移動到選取區域的結尾
    
        # 插入換行符號，然後在新行插入文字
        cursor.Text.insertControlCharacter(cursor, 0, False)  # 插入換行符號 (0 = 換行)
        cursor.Text.insertString(cursor, text, False)
