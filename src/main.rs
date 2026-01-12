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
    
    println!("微信自动化工具 v1.1");
    println!("联系人: {}", contact);
    println!("消息: {}", message);
    
    // 首先检查微信进程
    let ps_check = r#"
$wechatProcess = Get-Process -Name "WeChat" -ErrorAction SilentlyContinue
if ($wechatProcess) {
    Write-Host "✅ 找到微信进程"
} else {
    Write-Host "❌ 微信进程未运行，请先启动微信"
    exit 1
}
"#;
    
    println!("检查微信进程...");
    let check_output = Command::new("powershell")
        .args(&["-Command", ps_check])
        .output()?;
    
    if !check_output.status.success() {
        println!("{}", String::from_utf8_lossy(&check_output.stdout));
        return Ok(());
    }
    
    // 查找所有可能的微信窗口类名
    let ps_script = format!(r#"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$automation = [System.Windows.Automation.AutomationElement]::RootElement

# 尝试多种可能的微信窗口类名
$classNames = @("WeChatMainWndForPC", "ChatWnd", "WeChat")
$wechatWindow = $null

foreach ($className in $classNames) {{
    Write-Host "尝试查找窗口类名: $className"
    $condition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ClassNameProperty, $className)
    $wechatWindow = $automation.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
    if ($wechatWindow) {{
        Write-Host "✅ 找到微信窗口: $className"
        break
    }}
}}

if (-not $wechatWindow) {{
    # 尝试通过进程名查找窗口
    Write-Host "尝试通过进程名查找微信窗口..."
    $nameCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "微信")
    $wechatWindow = $automation.FindFirst([System.Windows.Automation.TreeScope]::Children, $nameCondition)
    if ($wechatWindow) {{
        Write-Host "✅ 通过名称找到微信窗口"
    }}
}}

if ($wechatWindow) {{
    Write-Host "微信窗口信息:"
    Write-Host "  名称: $($wechatWindow.Current.Name)"
    Write-Host "  类名: $($wechatWindow.Current.ClassName)"
    Write-Host "  控件类型: $($wechatWindow.Current.ControlType.LocalizedControlType)"
    
    # 查找联系人
    Write-Host "查找联系人: {}"
    $contactCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "{}")
    $contactElement = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $contactCondition)
    
    if ($contactElement) {{
        Write-Host "✅ 找到联系人，点击选择..."
        try {{
            $contactElement.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
            Start-Sleep -Milliseconds 1500
            
            # 查找输入框
            Write-Host "查找输入框..."
            $editCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
            $inputBoxes = $wechatWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, $editCondition)
            
            Write-Host "找到 $($inputBoxes.Count) 个输入框"
            
            $inputBox = $null
            foreach ($box in $inputBoxes) {{
                if ($box.Current.IsEnabled -and -not $box.Current.IsReadOnly) {{
                    $inputBox = $box
                    break
                }}
            }}
            
            if ($inputBox) {{
                Write-Host "✅ 找到可用输入框，输入消息..."
                $inputBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue("{}")
                Start-Sleep -Milliseconds 800
                
                # 尝试多种发送方式
                Write-Host "尝试发送消息..."
                
                # 方式1: 查找发送按钮
                $sendCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "发送")
                $sendButton = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $sendCondition)
                
                if ($sendButton) {{
                    $sendButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
                    Write-Host "✅ 通过发送按钮发送成功!"
                }} else {{
                    # 方式2: 发送Enter键
                    Write-Host "未找到发送按钮，尝试发送Enter键..."
                    $inputBox.SetFocus()
                    [System.Windows.Forms.SendKeys]::SendWait("{{ENTER}}")
                    Write-Host "✅ 通过Enter键发送成功!"
                }}
            }} else {{
                Write-Host "❌ 未找到可用的输入框"
            }}
        }} catch {{
            Write-Host "❌ 操作失败: $($_.Exception.Message)"
        }}
    }} else {{
        Write-Host "❌ 未找到联系人: {}"
        Write-Host "提示: 请确保联系人在聊天列表中可见"
    }}
}} else {{
    Write-Host "❌ 未找到微信窗口"
    Write-Host "请确保:"
    Write-Host "1. 微信已启动并完全加载"
    Write-Host "2. 微信窗口未最小化"
    Write-Host "3. 微信已登录"
}}
"#, contact, contact, message, contact);
    
    println!("执行微信自动化...");
    
    let output = Command::new("powershell")
        .args(&["-Command", ps_script])
        .output()?;
    
    println!("{}", String::from_utf8_lossy(&output.stdout));
    
    if !output.status.success() {
        println!("错误信息: {}", String::from_utf8_lossy(&output.stderr));
    }
    
    Ok(())
}
