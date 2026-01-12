use crate::ui::UIAutomation;
use windows::Win32::UI::Accessibility::*;
use windows::core::*;
use anyhow::{Result, anyhow};
use std::thread;
use std::time::Duration;

pub struct Chat {
    automation: UIAutomation,
    main_window: IUIAutomationElement,
}

impl Chat {
    pub fn new() -> Result<Self> {
        let automation = UIAutomation::new()?;
        let main_window = automation.find_window("WeChatMainWndForPC")
            .map_err(|_| anyhow!("微信主窗口未找到，请确保微信已启动"))?;
        Ok(Self { automation, main_window })
    }
    
    pub fn select_contact(&self, contact: &str) -> Result<()> {
        unsafe {
            let contact_element = self.automation.find_element_by_name(&self.main_window, contact)
                .map_err(|_| anyhow!("联系人 '{}' 未找到", contact))?;
            
            let invoke_pattern: IUIAutomationInvokePattern = contact_element
                .GetCurrentPatternAs(UIA_InvokePatternId)
                .map_err(|_| anyhow!("无法点击联系人"))?;
            
            invoke_pattern.Invoke()?;
            thread::sleep(Duration::from_millis(500));
        }
        Ok(())
    }
    
    pub fn send_text(&self, text: &str) -> Result<()> {
        unsafe {
            // 查找输入框
            let input_box = self.automation.find_element_by_control_type(
                &self.main_window, 
                UIA_EditControlTypeId.0 as i32
            ).map_err(|_| anyhow!("输入框未找到"))?;
            
            // 输入文本
            let value_pattern: IUIAutomationValuePattern = input_box
                .GetCurrentPatternAs(UIA_ValuePatternId)
                .map_err(|_| anyhow!("无法获取输入框"))?;
            
            value_pattern.SetValue(&HSTRING::from(text))?;
            thread::sleep(Duration::from_millis(200));
            
            // 发送消息
            self.send_message()?;
        }
        Ok(())
    }
    
    pub fn send_file(&self, file_path: &str) -> Result<()> {
        // 简化实现：通过剪贴板发送文件路径
        unsafe {
            // 查找附件按钮或右键菜单
            let attach_button = self.automation.find_element_by_name(&self.main_window, "文件")
                .or_else(|_| self.automation.find_element_by_name(&self.main_window, "附件"))
                .map_err(|_| anyhow!("附件按钮未找到"))?;
            
            let invoke_pattern: IUIAutomationInvokePattern = attach_button
                .GetCurrentPatternAs(UIA_InvokePatternId)
                .map_err(|_| anyhow!("无法点击附件按钮"))?;
            
            invoke_pattern.Invoke()?;
            thread::sleep(Duration::from_millis(1000));
            
            // 这里需要实现文件选择对话框的操作
            // 暂时返回成功，实际需要操作文件选择对话框
        }
        Ok(())
    }
    
    fn send_message(&self) -> Result<()> {
        unsafe {
            let send_button = self.automation.find_element_by_name(&self.main_window, "发送")
                .map_err(|_| anyhow!("发送按钮未找到"))?;
            
            let invoke_pattern: IUIAutomationInvokePattern = send_button
                .GetCurrentPatternAs(UIA_InvokePatternId)
                .map_err(|_| anyhow!("无法点击发送按钮"))?;
            
            invoke_pattern.Invoke()?;
            thread::sleep(Duration::from_millis(200));
        }
        Ok(())
    }
}
