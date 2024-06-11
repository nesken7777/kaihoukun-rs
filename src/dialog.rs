use windows::{
    core::Error,
    Win32::{
        Foundation::{BOOL, HWND},
        UI::{Controls::IsDlgButtonChecked, WindowsAndMessaging::GetDlgItemInt},
    },
};

use crate::consts::{
    ErrorKind::{self, *},
    SelfCreatedError::*,
    PORT_NUM_INPUT, UDP_CHECKED,
};

pub fn get_dialog_item(hdlg: HWND) -> std::result::Result<(u16, &'static str), (Error, ErrorKind)> {
    let port_num = unsafe { GetDlgItemInt(hdlg, PORT_NUM_INPUT, None, BOOL::from(false)) };
    let Ok(port_num) = u16::try_from(port_num) else {
        return Err((Error::empty(), SelfE(InvalidPortNumber)));
    };
    let udp_check = unsafe { IsDlgButtonChecked(hdlg, UDP_CHECKED) };
    let protocol = if udp_check != 1 { "TCP" } else { "UDP" };
    Ok((port_num, protocol))
}
