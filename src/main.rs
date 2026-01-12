use windows::Win32::UI::Accessibility::*;
use windows::Win32::System::Com::*;
use windows::core::*;
use anyhow::Result;
use std::thread;
use std::time::Duration;

fn main() -> Result<()> {
    let args: Vec<String> = std::env::args().collect();
    
    if args.len() < 3 {
        println!("用法: wxauto-rs <联系人> <消息>");
        println!("示例: wxauto-rs \"文件传输助手\" \"Hello from Rust!\"");
        return Ok(());
    }
    
    let contact = &args[1];
    let message = &args[2];
    
    println!("初始化微信自动化...");
    
    unsafe {
        CoInitialize(None)?;
        
        let automation: IUIAutomation = CoCreateInstance(
            &CUIAutomation,
            None,
            CLSCTX_INPROC_SERVER,
        )?;
        
        // 查找微信主窗口
        let desktop = automation.GetRootElement()?;
        let condition = automation.CreatePropertyCondition(
            UIA_ClassNamePropertyId,
            &VARIANT::from("WeChatMainWndForPC"),
        )?;
        
        let main_window = desktop.FindFirst(TreeScope_Children, &condition);
        
        match main_window {
            Ok(window) => {
                println!("找到微信窗口");
                
                // 查找联系人
                let contact_condition = automation.CreatePropertyCondition(
                    UIA_NamePropertyId,
                    &VARIANT::from(contact),
                )?;
                
                if let Ok(contact_element) = window.FindFirst(TreeScope_Descendants, &contact_condition) {
                    println!("找到联系人: {}", contact);
                    
                    // 点击联系人
                    if let Ok(invoke_pattern) = contact_element.GetCurrentPatternAs::<IUIAutomationInvokePattern>(UIA_InvokePatternId) {
                        invoke_pattern.Invoke()?;
                        thread::sleep(Duration::from_millis(500));
                        
                        // 查找输入框
                        let edit_condition = automation.CreatePropertyCondition(
                            UIA_ControlTypePropertyId,
                            &VARIANT::from(UIA_EditControlTypeId.0 as i32),
                        )?;
                        
                        if let Ok(input_box) = window.FindFirst(TreeScope_Descendants, &edit_condition) {
                            println!("找到输入框");
                            
                            // 输入文本
                            if let Ok(value_pattern) = input_box.GetCurrentPatternAs::<IUIAutomationValuePattern>(UIA_ValuePatternId) {
                                value_pattern.SetValue(&HSTRING::from(message))?;
                                thread::sleep(Duration::from_millis(200));
                                
                                // 查找发送按钮
                                let send_condition = automation.CreatePropertyCondition(
                                    UIA_NamePropertyId,
                                    &VARIANT::from("发送"),
                                )?;
                                
                                if let Ok(send_button) = window.FindFirst(TreeScope_Descendants, &send_condition) {
                                    if let Ok(send_invoke) = send_button.GetCurrentPatternAs::<IUIAutomationInvokePattern>(UIA_InvokePatternId) {
                                        send_invoke.Invoke()?;
                                        println!("消息发送成功!");
                                    }
                                } else {
                                    println!("未找到发送按钮");
                                }
                            } else {
                                println!("无法操作输入框");
                            }
                        } else {
                            println!("未找到输入框");
                        }
                    } else {
                        println!("无法点击联系人");
                    }
                } else {
                    println!("未找到联系人: {}", contact);
                }
            }
            Err(_) => {
                println!("未找到微信窗口，请确保微信已启动");
            }
        }
    }
    
    Ok(())
}
