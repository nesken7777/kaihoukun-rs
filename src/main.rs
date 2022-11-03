#![windows_subsystem = "windows"]

use windows::{
    core::*,
    Win32::{
        Foundation::{BOOL, HWND, LPARAM, WPARAM},
        NetworkManagement::WindowsFirewall::{IUPnPNAT, UPnPNAT},
        Networking::WinSock::{
            GetAddrInfoExW, GetHostNameW, WSACleanup, WSAStartup, ADDRESS_FAMILY, ADDRINFOEXW,
            AF_INET, NS_DNS, SOCKADDR_IN, SOCK_RAW, WSADATA,
        },
        System::{
            Com::{
                CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_HANDLER,
                CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, CLSCTX_REMOTE_SERVER,
                COINIT_MULTITHREADED,
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

use std::ptr::null_mut;

const DIALOG: usize = 4000;
const PORT_NUM_INPUT: i32 = 5001;
const TCP_CHECKED: i32 = 5002;
const UDP_CHECKED: i32 = 5003;
const OUT_TEXT: i32 = 5004;
const OPEN_PORT: usize = 5005;
const CLOSE_PORT: usize = 5006;
static mut G_HDLG: HWND = HWND(0);

enum ErrorKind {
    InvalidPortNumber,
    WSAStartupFail,
    GetAddrInfoExWFail,
    GetHostNameWFail,
    CoInitializeFail,
    CoCreateInstanceFail,
    StaticPortMappingCollectionFail,
    AddFail,
    RemoveFail,
}

use ErrorKind::*;

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
            2 => {
                EndDialog(window_handle, 2);
                0
            }
            OPEN_PORT => open_port().map_or_else(
                |(win_error, error_kind)| {
                    display_err(win_error, error_kind);
                    0
                },
                |_| {
                    SetDlgItemTextW(G_HDLG, OUT_TEXT, w!("ポート開放に成功しました。"));
                    1
                },
            ),
            CLOSE_PORT => close_port().map_or_else(
                |(win_error, error_kind)| {
                    display_err(win_error, error_kind);
                    0
                },
                |_| {
                    SetDlgItemTextW(G_HDLG, OUT_TEXT, w!("ポートを閉じました。"));
                    1
                },
            ),

            _ => 0,
        },
        _ => 0,
    }
}

fn open_port() -> std::result::Result<(), (Error, ErrorKind)> {
    let mut get_port_check = BOOL::default();
    let port_num = unsafe {
        GetDlgItemInt(
            G_HDLG,
            PORT_NUM_INPUT,
            Some(&mut get_port_check),
            BOOL::from(false),
        )
    };
    if get_port_check == BOOL::from(false) || port_num <= 0 || port_num > 65535 {
        return Err((Error::OK, InvalidPortNumber));
    }
    let udp_check = unsafe { IsDlgButtonChecked(G_HDLG, UDP_CHECKED) };
    let tcp_or_udp_string = if udp_check != 1 { "TCP" } else { "UDP" };
    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED)
            .map_err(|co_init_err| (co_init_err, CoInitializeFail))?
    };
    let upnp_nat = unsafe {
        CoCreateInstance::<_, IUPnPNAT>(
            &UPnPNAT,
            None,
            CLSCTX_INPROC_SERVER
                | CLSCTX_INPROC_HANDLER
                | CLSCTX_LOCAL_SERVER
                | CLSCTX_REMOTE_SERVER,
        )
        .map_err(|co_create_instance_err| (co_create_instance_err, CoCreateInstanceFail))?
    };
    let static_port_mapping_collection = unsafe {
        upnp_nat
            .StaticPortMappingCollection()
            .map_err(|static_port_mapping_collection_err| {
                (
                    static_port_mapping_collection_err,
                    StaticPortMappingCollectionFail,
                )
            })?
    };

    let mut wsa_data = WSADATA::default();
    let wsaresult = unsafe { WSAStartup(0x202, &mut wsa_data) };
    if wsaresult != 0 {
        return Err((Error::OK, WSAStartupFail));
    }

    let mut localhost_name = [0u16; 260];
    let gethostname_result = unsafe { GetHostNameW(localhost_name.as_mut_slice()) };
    if gethostname_result != 0 {
        return Err((Error::OK, GetHostNameWFail));
    }

    let hints = ADDRINFOEXW {
        ai_family: AF_INET.0 as i32,
        ai_socktype: SOCK_RAW as i32,
        ..Default::default()
    };
    let mut addr_info: *mut ADDRINFOEXW = null_mut();
    let dw_retval = unsafe {
        GetAddrInfoExW(
            PCWSTR::from_raw(localhost_name.as_ptr()),
            w!("7"),
            NS_DNS,
            None,
            Some(&hints),
            &mut addr_info,
            None,
            None,
            None,
            None,
        )
    };
    if dw_retval != 0 {
        return Err((Error::OK, GetAddrInfoExWFail));
    }
    let mut ptr = addr_info;

    let ip_str = {
        let mut return_str = String::with_capacity(64);
        while !ptr.is_null() && return_str.is_empty() {
            match unsafe { ADDRESS_FAMILY(((*ptr).ai_family) as u32) } {
                AF_INET => {
                    return_str = unsafe {
                        std::net::Ipv4Addr::from((*((*ptr).ai_addr as *mut SOCKADDR_IN)).sin_addr)
                            .to_string()
                    };
                }
                _ => {}
            }
            ptr = unsafe { (*ptr).ai_next };
        }
        return_str
    };
    unsafe {
        static_port_mapping_collection
            .Add(
                port_num as i32,
                &BSTR::from(tcp_or_udp_string),
                port_num as i32,
                &BSTR::from(ip_str),
                1,
                &BSTR::from("kaihoukun"),
            )
            .map_err(|add_err| (add_err, AddFail))?
    };
    unsafe {
        WSACleanup();
    }

    unsafe {
        CoUninitialize();
    }
    Ok(())
}

fn close_port() -> std::result::Result<(), (Error, ErrorKind)> {
    let mut get_port_check = BOOL::default();
    let port_num = unsafe {
        GetDlgItemInt(
            G_HDLG,
            PORT_NUM_INPUT,
            Some(&mut get_port_check),
            BOOL::from(false),
        )
    };
    if get_port_check == BOOL::from(false) || port_num <= 0 || port_num > 65535 {
        return Err((Error::OK, InvalidPortNumber));
    }
    let udp_check = unsafe { IsDlgButtonChecked(G_HDLG, UDP_CHECKED) };
    let tcp_or_udp_string = if udp_check != 1 { "TCP" } else { "UDP" };
    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED)
            .map_err(|co_init_err| (co_init_err, CoInitializeFail))?;
    }
    let upnp_nat = unsafe {
        CoCreateInstance::<Option<_>, IUPnPNAT>(
            &UPnPNAT,
            None,
            CLSCTX_INPROC_SERVER
                | CLSCTX_INPROC_HANDLER
                | CLSCTX_LOCAL_SERVER
                | CLSCTX_REMOTE_SERVER,
        )
        .map_err(|co_create_instance_err| (co_create_instance_err, CoCreateInstanceFail))?
    };
    let static_port_mapping_collection = unsafe {
        upnp_nat
            .StaticPortMappingCollection()
            .map_err(|static_port_mapping_collection_err| {
                (
                    static_port_mapping_collection_err,
                    StaticPortMappingCollectionFail,
                )
            })?
    };
    unsafe {
        static_port_mapping_collection
            .Remove(port_num as i32, &BSTR::from(tcp_or_udp_string))
            .map_err(|remove_err| (remove_err, RemoveFail))?
    };

    unsafe {
        CoUninitialize();
    }
    Ok(())
}

fn display_err(win_error: windows::core::Error, error_kind: ErrorKind) {
    match error_kind {
        InvalidPortNumber | WSAStartupFail | GetHostNameWFail | GetAddrInfoExWFail => {
            display_err_without_code(error_kind);
        }
        _ => display_err_with_code(win_error, error_kind),
    }
}

fn display_err_without_code(error_kind: ErrorKind) {
    let display_string = match error_kind {
        InvalidPortNumber => w!("ポート番号が不正です"),
        WSAStartupFail => w!("WSAStartup関数失敗です。"),
        GetHostNameWFail => w!("GetHostNameW関数失敗です。"),
        GetAddrInfoExWFail => w!("GetAddrInfoExW関数失敗です。"),
        _ => {
            unreachable!(
                "コード付きエラーを予期しましたが、コード無しエラーの関数に到達しました。"
            );
        }
    };
    unsafe {
        SetDlgItemTextW(G_HDLG, OUT_TEXT, display_string);
    }
}

fn display_err_with_code(win_error: Error, error_kind: ErrorKind) {
    let display_string = format!(
        "{}\r\n\r\nエラーコード:{:?}\r\n({})\r\n詳細は{}を参照されたし。",
        match error_kind{
            AddFail=>"ポート開放に失敗しました。",
            RemoveFail=>"ポートが閉じれませんでした。開いていますか？",
            CoInitializeFail=>"CoInitializeに失敗しました。",
            CoCreateInstanceFail=>"CoCreateInstanceに失敗しました。",
            StaticPortMappingCollectionFail=>"StaticPortMappingCollectionに失敗しました。",
            _=>unreachable!("コード無しエラーを予期しましたが、コード付きエラーの関数に到達しました。"),
        },
        win_error.code(),
        win_error.message().to_string().trim(),
        match error_kind{
            AddFail=>"https://docs.microsoft.com/en-us/windows/win32/api/natupnp/nf-natupnp-istaticportmappingcollection-add",
            RemoveFail=>"https://docs.microsoft.com/en-us/windows/win32/api/natupnp/nf-natupnp-istaticportmappingcollection-remove",
            CoInitializeFail=>"https://docs.microsoft.com/en-us/windows/win32/api/objbase/nf-objbase-coinitialize",
            CoCreateInstanceFail=>"https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstance",
            StaticPortMappingCollectionFail=>"https://docs.microsoft.com/en-us/windows/win32/api/natupnp/nf-natupnp-iupnpnat-get_staticportmappingcollection",
            _=>unreachable!("コード無しエラーを予期しましたが、コード付きエラーの関数に到達しました。"),
        },
    );
    unsafe {
        SetDlgItemTextW(
            G_HDLG,
            OUT_TEXT,
            PCWSTR::from(&HSTRING::from(display_string)),
        );
    }
}
