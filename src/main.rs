#![windows_subsystem = "windows"]

mod consts;
mod dialog;
mod display_err;
mod port_mapping;

use consts::{CLOSE_PORT, DIALOG, OPEN_PORT, OUT_TEXT, TCP_CHECKED};
use dialog::get_dialog_item;
use display_err::display_err;
use port_mapping::{close_port, open_port};
use windows::{
    core::*,
    Win32::{
        Foundation::{HWND, LPARAM, WPARAM},
        System::LibraryLoader::GetModuleHandleW,
        UI::{
            Controls::{CheckDlgButton, BST_CHECKED},
            WindowsAndMessaging::{
                DialogBoxParamW, EndDialog, SetDlgItemTextW, IDCANCEL, MESSAGEBOX_RESULT,
                WM_COMMAND, WM_INITDIALOG,
            },
        },
    },
};

fn main() -> Result<()> {
    unsafe {
        let instance = GetModuleHandleW(None)?;
        DialogBoxParamW(
            instance,
            PCWSTR::from_raw(DIALOG as *const u16),
            HWND(0),
            Some(dlg_proc),
            LPARAM(0),
        );
        Ok(())
    }
}

unsafe extern "system" fn dlg_proc(
    window_handle: HWND,
    message: u32,
    wparam: WPARAM,
    _: LPARAM,
) -> isize {
    match message {
        WM_INITDIALOG => {
            let _ = CheckDlgButton(window_handle, TCP_CHECKED, BST_CHECKED);
            0
        }
        WM_COMMAND => match wparam.0 & 0xffff {
            command @ (OPEN_PORT | CLOSE_PORT) => match get_dialog_item(window_handle) {
                Ok((port_num, protocol)) => match command {
                    OPEN_PORT => match open_port(port_num, protocol) {
                        Ok(external_ip) => {
                            let _ = SetDlgItemTextW(
                                window_handle,
                                OUT_TEXT,
                                PCWSTR::from_raw(
                                    HSTRING::from(format!(
                                        "ポート開放に成功しました。\r\n外部IPアドレス: {}",
                                        external_ip
                                    ))
                                    .as_ptr(),
                                ),
                            );
                            1
                        }
                        Err((error, kind)) => {
                            display_err(error, kind, window_handle);
                            0
                        }
                    },
                    CLOSE_PORT => match close_port(port_num, protocol) {
                        Ok(_) => {
                            let _ = SetDlgItemTextW(
                                window_handle,
                                OUT_TEXT,
                                w!("ポートを閉じました。"),
                            );
                            1
                        }
                        Err((error, kind)) => {
                            display_err(error, kind, window_handle);
                            0
                        }
                    },
                    _ => unreachable!(),
                },
                Err((error, kind)) => {
                    display_err(error, kind, window_handle);
                    0
                }
            },
            x if MESSAGEBOX_RESULT(x as i32) == IDCANCEL => {
                let _ = EndDialog(window_handle, 2);
                0
            }
            _ => 0,
        },
        _ => 0,
    }
}
