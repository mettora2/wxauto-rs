use windows::Win32::UI::Accessibility::*;
use windows::Win32::System::Com::*;
use windows::core::*;
use anyhow::Result;

pub struct UIAutomation {
    automation: IUIAutomation,
}

impl UIAutomation {
    pub fn new() -> Result<Self> {
        unsafe {
            CoInitialize(None)?;
            let automation: IUIAutomation = CoCreateInstance(
                &CUIAutomation,
                None,
                CLSCTX_INPROC_SERVER,
            )?;
            Ok(Self { automation })
        }
    }
    
    pub fn find_window(&self, class_name: &str) -> Result<IUIAutomationElement> {
        unsafe {
            let desktop = self.automation.GetRootElement()?;
            let condition = self.automation.CreatePropertyCondition(
                UIA_ClassNamePropertyId,
                &VARIANT::from(class_name),
            )?;
            desktop.FindFirst(TreeScope_Children, &condition)
        }
    }
    
    pub fn find_element_by_name(&self, parent: &IUIAutomationElement, name: &str) -> Result<IUIAutomationElement> {
        unsafe {
            let condition = self.automation.CreatePropertyCondition(
                UIA_NamePropertyId,
                &VARIANT::from(name),
            )?;
            parent.FindFirst(TreeScope_Descendants, &condition)
        }
    }
    
    pub fn find_element_by_control_type(&self, parent: &IUIAutomationElement, control_type: i32) -> Result<IUIAutomationElement> {
        unsafe {
            let condition = self.automation.CreatePropertyCondition(
                UIA_ControlTypePropertyId,
                &VARIANT::from(control_type),
            )?;
            parent.FindFirst(TreeScope_Descendants, &condition)
        }
    }
}
