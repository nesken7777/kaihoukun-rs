#![windows_subsystem = "windows"]

use windows::{
    core::*,
    Win32::{
        Foundation::{BOOL, BSTR, HWND, LPARAM, WPARAM},
        NetworkManagement::WindowsFirewall::{IUPnPNAT, UPnPNAT},
        Networking::WinSock::{
            gethostbyname, gethostname, inet_ntoa, WSAData, WSAGetLastError, WSAStartup, IN_ADDR,
            IN_ADDR_0,
        },
        System::{
            Com::{
                CoCreateInstance, CoInitialize, CoUninitialize, CLSCTX_INPROC_HANDLER,
                CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, CLSCTX_REMOTE_SERVER,
            },
            LibraryLoader::GetModuleHandleW,
        },
        UI::{
            Controls::{CheckDlgButton, IsDlgButtonChecked, BST_CHECKED},
            WindowsAndMessaging::{
                DialogBoxParamW, EndDialog, GetDlgItemInt, SetDlgItemTextW, WM_COMMAND,
                WM_INITDIALOG,
            },
        },
    },
};

use std::{ffi::CStr, ptr::null};

const DIALOG: usize = 4000;
const PORT_NUM_INPUT: i32 = 5001;
const TCP_CHECKED: i32 = 5002;
const UDP_CHECKED: i32 = 5003;
const OUT_TEXT: i32 = 5004;
const OPEN_PORT: usize = 5005;
const CLOSE_PORT: usize = 5006;
static mut G_HDLG: HWND = HWND(0);

fn main() -> Result<()> {
    unsafe {
        let instance = GetModuleHandleW(None)?;
        DialogBoxParamW(
            instance,
            PCWSTR(DIALOG as *const u16),
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
            2 => {
                EndDialog(window_handle, 2);
                0
            }
            OPEN_PORT => {
                open_port();
                0
            }
            CLOSE_PORT => {
                close_port();
                0
            }

            _ => 0,
        },
        _ => 0,
    }
}

fn open_port() {
    let mut get_port_check = BOOL::default();
    let port_num = unsafe {
        GetDlgItemInt(
            G_HDLG,
            PORT_NUM_INPUT,
            &mut get_port_check,
            BOOL::from(false),
        )
    };
    if get_port_check == BOOL::from(false) || port_num <= 0 || port_num > 65535 {
        unsafe {
            SetDlgItemTextW(G_HDLG, OUT_TEXT, "ポート番号が不正です");
        }
    } else {
        let udp_check = unsafe { IsDlgButtonChecked(G_HDLG, UDP_CHECKED) };
        let tcp_or_udp_string = if udp_check != 1 { "TCP" } else { "UDP" };
        if let Ok(()) = unsafe { CoInitialize(null()) } {
            if let Ok(p_upnp_nat) = unsafe {
                CoCreateInstance::<Option<_>, IUPnPNAT>(
                    &UPnPNAT,
                    None,
                    CLSCTX_INPROC_SERVER
                        | CLSCTX_INPROC_HANDLER
                        | CLSCTX_LOCAL_SERVER
                        | CLSCTX_REMOTE_SERVER,
                )
            } {
                if let Ok(p_static_port_mapping_collection) =
                    unsafe { p_upnp_nat.StaticPortMappingCollection() }
                {
                    let mut wsadata = WSAData::default();
                    let wsaresult = unsafe { WSAStartup(0x101, &mut wsadata) };
                    if wsaresult == 0 {
                        let mut localhost_name = [0u8; 260];
                        let gethostname_result =
                            unsafe { gethostname(PSTR(localhost_name.as_mut_ptr()), 256) };
                        if gethostname_result == 0 {
                            let hostent_a =
                                unsafe { gethostbyname(PCSTR(localhost_name.as_ptr())) };
                            if hostent_a.is_null() {
                                unsafe {
                                    SetDlgItemTextW(
                                        G_HDLG,
                                        OUT_TEXT,
                                        "gethostbyname関数失敗です。",
                                    );
                                }
                            } else {
                                let ipaddress_pstr = unsafe {
                                    inet_ntoa(IN_ADDR {
                                        S_un: IN_ADDR_0 {
                                            S_addr: *((*(*hostent_a).h_addr_list) as *mut u32),
                                        },
                                    })
                                };
                                let description_bstr = BSTR::from("kaihoukun");
                                let ipaddress_bstr = unsafe {
                                    BSTR::from(
                                        CStr::from_ptr(ipaddress_pstr.0 as *const i8)
                                            .to_str()
                                            .unwrap(),
                                    )
                                };
                                let tcpudp_bstr = BSTR::from(tcp_or_udp_string);
                                if let Ok(_) = unsafe {
                                    p_static_port_mapping_collection.Add(
                                        port_num as i32,
                                        tcpudp_bstr,
                                        port_num as i32,
                                        ipaddress_bstr,
                                        1,
                                        description_bstr,
                                    )
                                } {
                                    unsafe {
                                        SetDlgItemTextW(
                                            G_HDLG,
                                            OUT_TEXT,
                                            "ポート開放に成功しました。",
                                        );
                                    }
                                } else {
                                    unsafe {
                                        SetDlgItemTextW(
                                            G_HDLG,
                                            OUT_TEXT,
                                            "ポート開放に失敗しました。",
                                        );
                                    }
                                }
                            }
                        } else {
                            unsafe {
                                let lasterror = format!(
                                    "gethostname関数失敗です。\nエラーコード:{:?}",
                                    WSAGetLastError()
                                );

                                SetDlgItemTextW(G_HDLG, OUT_TEXT, lasterror);
                            }
                        }
                    } else {
                        unsafe {
                            SetDlgItemTextW(G_HDLG, OUT_TEXT, "WSAStartup関数失敗です。");
                        }
                    }
                } else {
                    unsafe {
                        SetDlgItemTextW(
                            G_HDLG,
                            OUT_TEXT,
                            "get_StaticPortMappingCollectionに失敗しました。",
                        );
                    }
                }
            } else {
                unsafe {
                    SetDlgItemTextW(G_HDLG, OUT_TEXT, "CoCreateInstanceに失敗しました。");
                }
            }
            unsafe {
                CoUninitialize();
            }
        } else {
            unsafe {
                SetDlgItemTextW(G_HDLG, OUT_TEXT, "CoInitializeに失敗しました。");
            }
        }
    }
}

fn close_port() {
    let mut get_port_check = BOOL::default();
    let port_num = unsafe {
        GetDlgItemInt(
            G_HDLG,
            PORT_NUM_INPUT,
            &mut get_port_check,
            BOOL::from(false),
        )
    };
    if get_port_check == BOOL::from(false) || port_num <= 0 || port_num > 65535 {
        unsafe {
            SetDlgItemTextW(G_HDLG, OUT_TEXT, "ポート番号が不正です");
        }
    } else {
        let udp_check = unsafe { IsDlgButtonChecked(G_HDLG, UDP_CHECKED) };
        let tcp_or_udp_string = if udp_check != 1 { "TCP" } else { "UDP" };
        if let Ok(()) = unsafe { CoInitialize(null()) } {
            if let Ok(p_upnp_nat) = unsafe {
                CoCreateInstance::<Option<_>, IUPnPNAT>(
                    &UPnPNAT,
                    None,
                    CLSCTX_INPROC_SERVER
                        | CLSCTX_INPROC_HANDLER
                        | CLSCTX_LOCAL_SERVER
                        | CLSCTX_REMOTE_SERVER,
                )
            } {
                if let Ok(p_static_port_mapping_collection) =
                    unsafe { p_upnp_nat.StaticPortMappingCollection() }
                {
                    if let Ok(()) = unsafe {
                        p_static_port_mapping_collection
                            .Remove(port_num as i32, BSTR::from(tcp_or_udp_string))
                    } {
                        unsafe {
                            SetDlgItemTextW(G_HDLG, OUT_TEXT, "ポートを閉じました。");
                        }
                    } else {
                        unsafe {
                            SetDlgItemTextW(
                                G_HDLG,
                                OUT_TEXT,
                                "ポートが閉じれませんでした。開いていますか？",
                            );
                        }
                    }
                } else {
                    unsafe {
                        SetDlgItemTextW(
                            G_HDLG,
                            OUT_TEXT,
                            "get_StaticPortMappingCollectionに失敗しました。",
                        );
                    }
                }
            } else {
                unsafe {
                    SetDlgItemTextW(G_HDLG, OUT_TEXT, "CoCreateInstanceに失敗しました。");
                }
            }
            unsafe {
                CoUninitialize();
            }
        } else {
            unsafe {
                SetDlgItemTextW(G_HDLG, OUT_TEXT, "CoInitializeに失敗しました。");
            }
        }
    }
}
