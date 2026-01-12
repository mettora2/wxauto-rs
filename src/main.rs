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
    
    println!("微信自动化工具 v1.3");
    println!("联系人: {}", contact);
    println!("消息: {}", message);
    
    let ps_script = format!(r#"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Write-Host "正在查找微信窗口..."
$automation = [System.Windows.Automation.AutomationElement]::RootElement

$wechatWindow = $null
$classNames = @("WeChatMainWndForPC", "ChatWnd", "WeChat")
foreach ($className in $classNames) {{
    $condition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ClassNameProperty, $className)
    $wechatWindow = $automation.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
    if ($wechatWindow) {{
        Write-Host "找到微信窗口 (类名: $className)"
        break
    }}
}}

if (-not $wechatWindow) {{
    $nameCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "微信")
    $wechatWindow = $automation.FindFirst([System.Windows.Automation.TreeScope]::Children, $nameCondition)
    if ($wechatWindow) {{
        Write-Host "找到微信窗口 (通过标题)"
    }}
}}

if ($wechatWindow) {{
    # 先尝试使用搜索功能
    Write-Host "尝试使用搜索功能查找联系人..."
    
    # 查找搜索框 (通常在顶部)
    $searchCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
    $searchBoxes = $wechatWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, $searchCondition)
    
    $searchBox = $null
    foreach ($box in $searchBoxes) {{
        $name = $box.Current.Name
        if ($name -like "*搜索*" -or $name -eq "" -or $name -like "*查找*") {{
            $searchBox = $box
            break
        }}
    }}
    
    if ($searchBox) {{
        Write-Host "找到搜索框，搜索联系人..."
        $searchBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue("{}")
        Start-Sleep -Seconds 2
        
        # 查找搜索结果中的联系人
        $contactCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "{}")
        $contactElement = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $contactCondition)
        
        if ($contactElement) {{
            Write-Host "在搜索结果中找到联系人，点击..."
            $contactElement.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
            Start-Sleep -Seconds 2
        }} else {{
            Write-Host "搜索结果中未找到联系人，清空搜索框..."
            $searchBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue("")
            Start-Sleep -Seconds 1
        }}
    }}
    
    # 如果搜索没找到，直接在聊天列表中查找
    if (-not $contactElement) {{
        Write-Host "在聊天列表中查找联系人: {}"
        $contactCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "{}")
        $contactElement = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $contactCondition)
        
        if ($contactElement) {{
            Write-Host "在聊天列表中找到联系人，点击..."
            $contactElement.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
            Start-Sleep -Seconds 2
        }}
    }}
    
    if ($contactElement) {{
        Write-Host "查找输入框..."
        Start-Sleep -Seconds 1
        
        # 查找消息输入框
        $editCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
        $inputBoxes = $wechatWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, $editCondition)
        
        $inputBox = $null
        foreach ($box in $inputBoxes) {{
            $name = $box.Current.Name
            # 排除搜索框，找消息输入框
            if ($name -notlike "*搜索*" -and $name -notlike "*查找*" -and $box.Current.IsEnabled) {{
                $inputBox = $box
                break
            }}
        }}
        
        if ($inputBox) {{
            Write-Host "找到输入框，输入消息..."
            $inputBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue("{}")
            Start-Sleep -Seconds 1
            
            Write-Host "查找发送按钮..."
            $sendCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "发送")
            $sendButton = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $sendCondition)
            
            if ($sendButton) {{
                $sendButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
                Write-Host "✅ 消息发送成功!"
            }} else {{
                Write-Host "⚠️ 未找到发送按钮，消息已输入，请手动按Enter发送"
            }}
        }} else {{
            Write-Host "❌ 未找到消息输入框"
        }}
    }} else {{
        Write-Host "❌ 未找到联系人: {}"
        Write-Host "请确保:"
        Write-Host "1. 联系人名称完全正确"
        Write-Host "2. 该联系人存在于您的微信中"
        Write-Host "3. 尝试先手动搜索该联系人"
    }}
}} else {{
    Write-Host "❌ 未找到微信窗口"
    Write-Host "请确保微信已启动并登录"
}}
"#, contact, contact, contact, contact, message, contact);
    
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
