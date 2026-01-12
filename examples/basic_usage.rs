use wxauto_rs::WeChat;

fn main() -> anyhow::Result<()> {
    println!("初始化微信自动化...");
    let wx = WeChat::new()?;
    
    // 发送文本消息
    println!("发送文本消息...");
    wx.send_text("文件传输助手", "Hello from Rust!")?;
    
    // 发送文件 (暂未完全实现)
    // wx.send_file("文件传输助手", "C:\\path\\to\\file.txt")?;
    
    println!("操作完成!");
    Ok(())
}
