use windows::Win32::UI::Accessibility::*;
use windows::Win32::System::Com::*;
use windows::Win32::Foundation::*;
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
        
        println!("查找微信窗口...");
        let desktop = automation.GetRootElement()?;
        
        let class_name = HSTRING::from("WeChatMainWndForPC");
        let condition = automation.CreatePropertyCondition(
            UIA_ClassNamePropertyId,
            &VARIANT::from(&class_name),
        )?;
        
        match desktop.FindFirst(TreeScope_Children, &condition) {
            Ok(main_window) => {
                println!("找到微信窗口");
                
                // 查找联系人
                println!("查找联系人: {}", contact);
                let contact_name = HSTRING::from(contact);
                let contact_condition = automation.CreatePropertyCondition(
                    UIA_NamePropertyId,
                    &VARIANT::from(&contact_name),
                )?;
                
                match main_window.FindFirst(TreeScope_Descendants, &contact_condition) {
                    Ok(contact_element) => {
                        println!("找到联系人，点击选择...");
                        
                        if let Ok(invoke_pattern) = contact_element.GetCurrentPatternAs::<IUIAutomationInvokePattern>(UIA_InvokePatternId) {
                            invoke_pattern.Invoke()?;
                            thread::sleep(Duration::from_millis(1000));
                            
                            // 查找输入框
                            println!("查找输入框...");
                            let edit_type = VARIANT::from(UIA_EditControlTypeId.0 as i32);
                            let edit_condition = automation.CreatePropertyCondition(
                                UIA_ControlTypePropertyId,
                                &edit_type,
                            )?;
                            
                            match main_window.FindFirst(TreeScope_Descendants, &edit_condition) {
                                Ok(input_box) => {
                                    println!("找到输入框，输入消息...");
                                    
                                    if let Ok(value_pattern) = input_box.GetCurrentPatternAs::<IUIAutomationValuePattern>(UIA_ValuePatternId) {
                                        let msg = HSTRING::from(message);
                                        value_pattern.SetValue(&msg)?;
                                        thread::sleep(Duration::from_millis(500));
                                        
                                        // 查找发送按钮
                                        println!("查找发送按钮...");
                                        let send_name = HSTRING::from("发送");
                                        let send_condition = automation.CreatePropertyCondition(
                                            UIA_NamePropertyId,
                                            &VARIANT::from(&send_name),
                                        )?;
                                        
                                        match main_window.FindFirst(TreeScope_Descendants, &send_condition) {
                                            Ok(send_button) => {
                                                if let Ok(send_invoke) = send_button.GetCurrentPatternAs::<IUIAutomationInvokePattern>(UIA_InvokePatternId) {
                                                    send_invoke.Invoke()?;
                                                    println!("✅ 消息发送成功!");
                                                } else {
                                                    println!("❌ 无法点击发送按钮");
                                                }
                                            }
                                            Err(_) => {
                                                println!("❌ 未找到发送按钮");
                                            }
                                        }
                                    } else {
                                        println!("❌ 无法操作输入框");
                                    }
                                }
                                Err(_) => {
                                    println!("❌ 未找到输入框");
                                }
                            }
                        } else {
                            println!("❌ 无法点击联系人");
                        }
                    }
                    Err(_) => {
                        println!("❌ 未找到联系人: {}", contact);
                        println!("提示: 请确保联系人名称完全匹配");
                    }
                }
            }
            Err(_) => {
                println!("❌ 未找到微信窗口");
                println!("请确保微信已启动并登录");
            }
        }
    }
    
    Ok(())
}
