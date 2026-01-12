# wxauto-rs

Rust版本的微信自动化工具，专注于消息和文件发送功能。

## 快速使用 (推荐)

**直接使用PowerShell脚本 (无需编译)：**

```powershell
# 下载脚本
curl -o wxauto.ps1 https://raw.githubusercontent.com/mettora2/wxauto-rs/master/wxauto.ps1

# 使用方法
.\wxauto.ps1 -Contact "文件传输助手" -Message "Hello from PowerShell!"

# 或者一行命令
powershell -ExecutionPolicy Bypass -File wxauto.ps1 -Contact "联系人名称" -Message "消息内容"
```

## 编译版本

**下载编译好的exe：**
- 访问 [Releases](https://github.com/mettora2/wxauto-rs/releases) 页面
- 或从 [Actions](https://github.com/mettora2/wxauto-rs/actions) 下载最新构建

**使用方法：**
```bash
wxauto-rs.exe "联系人名称" "消息内容"
```

## 功能特点

- ✅ 发送文本消息
- ✅ 智能搜索联系人
- ✅ 支持不在最近聊天列表中的联系人
- ✅ 基于Windows UIAutomation
- ✅ 支持Windows 10/11

## 环境要求

- Windows 10/11
- 微信 PC版 (已启动并登录)
- PowerShell 5.0+ (Windows自带)

## 故障排除

1. **找不到联系人**：确保联系人名称完全正确
2. **编码问题**：使用PowerShell脚本版本
3. **权限问题**：运行 `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`

## 开发

```bash
# 克隆项目
git clone https://github.com/mettora2/wxauto-rs.git

# 编译 (需要Rust环境)
cargo build --release
```
