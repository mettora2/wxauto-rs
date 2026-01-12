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
    
    println!("微信自动化工具 v0.1.0");
    println!("联系人: {}", contact);
    println!("消息: {}", message);
    println!("注意: 此版本为演示版本，需要在Windows环境下配合微信使用");
    
    // 基础版本，仅显示参数
    // 实际的Windows UIAutomation功能需要在Windows环境下测试
    
    Ok(())
}
