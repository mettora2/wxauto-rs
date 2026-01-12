use wxauto_rs::WeChat;
use std::env;

fn main() -> anyhow::Result<()> {
    let args: Vec<String> = env::args().collect();
    
    if args.len() < 3 {
        println!("用法: wxauto-rs <联系人> <消息>");
        println!("示例: wxauto-rs \"文件传输助手\" \"Hello from Rust!\"");
        return Ok(());
    }
    
    let contact = &args[1];
    let message = &args[2];
    
    println!("初始化微信自动化...");
    let wx = WeChat::new()?;
    
    println!("发送消息给: {}", contact);
    wx.send_text(contact, message)?;
    
    println!("消息发送成功!");
    Ok(())
}
