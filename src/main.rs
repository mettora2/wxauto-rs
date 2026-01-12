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
    
    println!("微信自动化工具 v1.2");
    println!("联系人: {}", contact);
    println!("消息: {}", message);
    
    // 简化的PowerShell脚本，专注于核心功能
    let ps_script = format!(r#"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Write-Host "正在查找微信窗口..."
$automation = [System.Windows.Automation.AutomationElement]::RootElement

# 尝试多种微信窗口查找方式
$wechatWindow = $null

# 方式1: 通过类名查找
$classNames = @("WeChatMainWndForPC", "ChatWnd", "WeChat")
foreach ($className in $classNames) {{
    $condition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ClassNameProperty, $className)
    $wechatWindow = $automation.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
    if ($wechatWindow) {{
        Write-Host "找到微信窗口 (类名: $className)"
        break
    }}
}}

# 方式2: 通过窗口标题查找
if (-not $wechatWindow) {{
    $nameCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "微信")
    $wechatWindow = $automation.FindFirst([System.Windows.Automation.TreeScope]::Children, $nameCondition)
    if ($wechatWindow) {{
        Write-Host "找到微信窗口 (通过标题)"
    }}
}}

if ($wechatWindow) {{
    Write-Host "查找联系人: {}"
    $contactCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "{}")
    $contactElement = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $contactCondition)
    
    if ($contactElement) {{
        Write-Host "找到联系人，点击..."
        $contactElement.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
        Start-Sleep -Seconds 2
        
        Write-Host "查找输入框..."
        $editCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
        $inputBox = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $editCondition)
        
        if ($inputBox) {{
            Write-Host "输入消息..."
            $inputBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue("{}")
            Start-Sleep -Seconds 1
            
            Write-Host "查找发送按钮..."
            $sendCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "发送")
            $sendButton = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $sendCondition)
            
            if ($sendButton) {{
                $sendButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
                Write-Host "✅ 消息发送成功!"
            }} else {{
                Write-Host "⚠️ 未找到发送按钮，消息已输入到输入框"
                Write-Host "请手动按Enter键发送"
            }}
        }} else {{
            Write-Host "❌ 未找到输入框"
        }}
    }} else {{
        Write-Host "❌ 未找到联系人: {}"
        Write-Host "请确保联系人在聊天列表中可见"
    }}
}} else {{
    Write-Host "❌ 未找到微信窗口"
    Write-Host "请确保:"
    Write-Host "1. 微信已启动并登录"
    Write-Host "2. 微信窗口未最小化"
}}
"#, contact, contact, message, contact);
    
    println!("执行微信自动化...");
    
    let output = Command::new("powershell")
        .args(&["-ExecutionPolicy", "Bypass", "-Command", &ps_script])
        .output()?;
    
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    
    println!("{}", stdout);
    
    if !stderr.is_empty() {
        println!("错误信息: {}", stderr);
    }
    
    Ok(())
}
