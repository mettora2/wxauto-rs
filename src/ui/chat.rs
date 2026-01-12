use crate::ui::UIAutomation;
use windows::Win32::UI::Accessibility::*;
use windows::core::*;
use anyhow::{Result, anyhow};
use std::thread;
use std::time::Duration;

pub struct Chat {
    automation: UIAutomation,
    main_window: Option<IUIAutomationElement>,
}

impl Chat {
    pub fn new() -> Result<Self> {
        let automation = UIAutomation::new()?;
        let main_window = automation.find_window("WeChatMainWndForPC").ok();
        Ok(Self { automation, main_window })
    }
    
    pub fn select_contact(&self, contact: &str) -> Result<()> {
        let main_window = self.main_window.as_ref()
            .ok_or_else(|| anyhow!("微信主窗口未找到，请确保微信已启动"))?;
            
        unsafe {
            let contact_element = self.automation.find_element_by_name(main_window, contact)
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
        let main_window = self.main_window.as_ref()
            .ok_or_else(|| anyhow!("微信主窗口未找到"))?;
            
        unsafe {
            // 查找输入框
            let input_box = self.automation.find_element_by_control_type(
                main_window, 
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
    
    pub fn send_file(&self, _file_path: &str) -> Result<()> {
        // 文件发送功能待实现
        Ok(())
    }
    
    fn send_message(&self) -> Result<()> {
        let main_window = self.main_window.as_ref()
            .ok_or_else(|| anyhow!("微信主窗口未找到"))?;
            
        unsafe {
            let send_button = self.automation.find_element_by_name(main_window, "发送")
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
