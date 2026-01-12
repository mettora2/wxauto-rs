use anyhow::Result;
use crate::ui::Chat;

pub struct WeChat {
    chat: Chat,
}

impl WeChat {
    pub fn new() -> Result<Self> {
        let chat = Chat::new()?;
        Ok(Self { chat })
    }
    
    pub fn send_text(&self, contact: &str, text: &str) -> Result<()> {
        self.chat.select_contact(contact)?;
        self.chat.send_text(text)
    }
    
    pub fn send_file(&self, contact: &str, file_path: &str) -> Result<()> {
        self.chat.select_contact(contact)?;
        self.chat.send_file(file_path)
    }
}
