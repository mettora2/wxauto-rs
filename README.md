# wxauto-rs

Rust版本的微信自动化工具，专注于消息和文件发送功能。

## 功能

- 发送文本消息
- 发送文件（开发中）
- 基于Windows UIAutomation

## 使用

```rust
use wxauto_rs::WeChat;

let wx = WeChat::new()?;
wx.send_text("联系人名称", "Hello from Rust!")?;
```

## 环境要求

- Windows 10/11
- 微信 3.9.x
- Rust 1.70+
