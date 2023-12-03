#![windows_subsystem = "windows"]

mod close;
mod consts;
mod display_err;
mod open;

use close::close_port;
use consts::{CLOSE_PORT, DIALOG, G_HDLG, OPEN_PORT, OUT_TEXT, TCP_CHECKED};
use display_err::display_err;
use open::open_port;

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
            CheckDlgButton(window_handle, TCP_CHECKED, BST_CHECKED);
            G_HDLG = window_handle;
            0
        }
        WM_COMMAND => match wparam.0 & 0xffff {
            OPEN_PORT => match open_port() {
                Ok(_) => {
                    SetDlgItemTextW(G_HDLG, OUT_TEXT, w!("ポート開放に成功しました。"));
                    1
                }
                Err((win_error, error_kind)) => {
                    display_err(win_error, error_kind);
                    0
                }
            },
            CLOSE_PORT => match close_port() {
                Ok(_) => {
                    SetDlgItemTextW(G_HDLG, OUT_TEXT, w!("ポートを閉じました。"));
                    1
                }
                Err((win_error, error_kind)) => {
                    display_err(win_error, error_kind);
                    0
                }
            },
            x if MESSAGEBOX_RESULT(x as i32) == IDCANCEL => {
                EndDialog(window_handle, 2);
                0
            }
            _ => 0,
        },
        _ => 0,
    }
}
