param(
    [string]$Contact,
    [string]$Message
)

if (-not $Contact -or -not $Message) {
    Write-Host "Usage: .\wxauto.ps1 -Contact '联系人' -Message '消息'"
    Write-Host "Example: .\wxauto.ps1 -Contact '文件传输助手' -Message 'Hello!'"
    exit
}

# Set UTF-8 encoding to fix Chinese display issues
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Write-Host "Finding WeChat window..."
$automation = [System.Windows.Automation.AutomationElement]::RootElement

$wechatWindow = $null
$classNames = @("WeChatMainWndForPC", "ChatWnd", "WeChat")
foreach ($className in $classNames) {
    $condition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ClassNameProperty, $className)
    $wechatWindow = $automation.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
    if ($wechatWindow) {
        Write-Host "Found WeChat window (class: $className)"
        break
    }
}

if (-not $wechatWindow) {
    $nameCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "微信")
    $wechatWindow = $automation.FindFirst([System.Windows.Automation.TreeScope]::Children, $nameCondition)
    if ($wechatWindow) {
        Write-Host "Found WeChat window (by title)"
    }
}

if ($wechatWindow) {
    Write-Host "Searching for contact: $Contact"
    
    # Try search function first
    $searchCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
    $searchBoxes = $wechatWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, $searchCondition)
    
    $searchBox = $null
    foreach ($box in $searchBoxes) {
        $name = $box.Current.Name
        if ($name -like "*搜索*" -or $name -eq "" -or $name -like "*查找*") {
            $searchBox = $box
            break
        }
    }
    
    $contactElement = $null
    if ($searchBox) {
        Write-Host "Using search box to find contact..."
        $searchBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($Contact)
        Start-Sleep -Seconds 2
        
        $contactCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, $Contact)
        $contactElement = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $contactCondition)
        
        if ($contactElement) {
            Write-Host "Found contact in search results, clicking..."
            $contactElement.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
            Start-Sleep -Seconds 2
        } else {
            Write-Host "Contact not found in search, clearing search box..."
            $searchBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue("")
            Start-Sleep -Seconds 1
        }
    }
    
    # If search didn't work, try direct lookup in chat list
    if (-not $contactElement) {
        Write-Host "Looking for contact in chat list: $Contact"
        $contactCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, $Contact)
        $contactElement = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $contactCondition)
        
        if ($contactElement) {
            Write-Host "Found contact in chat list, clicking..."
            $contactElement.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
            Start-Sleep -Seconds 2
        }
    }
    
    if ($contactElement) {
        Write-Host "Looking for message input box..."
        Start-Sleep -Seconds 1
        
        $editCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
        $inputBoxes = $wechatWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, $editCondition)
        
        $inputBox = $null
        foreach ($box in $inputBoxes) {
            $name = $box.Current.Name
            if ($name -notlike "*搜索*" -and $name -notlike "*查找*" -and $box.Current.IsEnabled) {
                $inputBox = $box
                break
            }
        }
        
        if ($inputBox) {
            Write-Host "Found input box, typing message..."
            $inputBox.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($Message)
            Start-Sleep -Seconds 1
            
            Write-Host "Looking for send button..."
            $sendCondition = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, "发送")
            $sendButton = $wechatWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $sendCondition)
            
            if ($sendButton) {
                $sendButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
                Write-Host "SUCCESS: Message sent!"
            } else {
                Write-Host "WARNING: Send button not found, message typed. Please press Enter manually."
            }
        } else {
            Write-Host "ERROR: Message input box not found"
        }
    } else {
        Write-Host "ERROR: Contact not found: $Contact"
        Write-Host "Please ensure:"
        Write-Host "1. Contact name is exactly correct"
        Write-Host "2. Contact exists in your WeChat"
        Write-Host "3. Try manually searching for the contact first"
    }
} else {
    Write-Host "ERROR: WeChat window not found"
    Write-Host "Please ensure WeChat is running and logged in"
}
