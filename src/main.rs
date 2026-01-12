use std::process::Command;
use anyhow::Result;

fn main() -> Result<()> {
    let args: Vec<String> = std::env::args().collect();
    
    if args.len() < 3 {
        println!("用法: wxauto-rs <联系人> <消息>");
        println!("示例: wxauto-rs \"文件传输助手\" \"Hello from Rust!\"");
        return Ok(());
    }
    
    let contact = &args[1];
    let message = &args[2];
    
    println!("微信自动化工具 v1.0");
    println!("联系人: {}", contact);
    println!("消息: {}", message);
    
    // 使用PowerShell调用UIAutomation
    let ps_script = format!(r#"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$automation = [System.Windows.Automation.AutomationElement]::RootElement
$wechatCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ClassNameProperty, "WeChatMainWndForPC")
$wechatWindow = $automation.FindFirst([System.Windows.Automation.TreeScope]::Children, $wechatCondition)

if ($wechatWindow) {{
    Write-Host "找到微信窗口"
    
    $contactCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "{}")
    $contactElement = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $contactCondition)
    
    if ($contactElement) {{
        Write-Host "找到联系人，点击选择..."
        $contactElement.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
        Start-Sleep -Milliseconds 1000
        
        $editCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
        $inputBox = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $editCondition)
        
        if ($inputBox) {{
            Write-Host "找到输入框，输入消息..."
            $inputBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue("{}")
            Start-Sleep -Milliseconds 500
            
            $sendCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "发送")
            $sendButton = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $sendCondition)
            
            if ($sendButton) {{
                $sendButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
                Write-Host "✅ 消息发送成功!"
            }} else {{
                Write-Host "❌ 未找到发送按钮"
            }}
        }} else {{
            Write-Host "❌ 未找到输入框"
        }}
    }} else {{
        Write-Host "❌ 未找到联系人: {}"
    }}
}} else {{
    Write-Host "❌ 未找到微信窗口，请确保微信已启动"
}}
"#, contact, message, contact);
    
    println!("执行微信自动化...");
    
    let output = Command::new("powershell")
        .args(&["-Command", &ps_script])
        .output()?;
    
    if output.status.success() {
        println!("{}", String::from_utf8_lossy(&output.stdout));
    } else {
        println!("执行失败: {}", String::from_utf8_lossy(&output.stderr));
        println!("请确保:");
        println!("1. 在Windows系统上运行");
        println!("2. 微信已启动并登录");
        println!("3. 有PowerShell执行权限");
    }
    
    Ok(())
}
