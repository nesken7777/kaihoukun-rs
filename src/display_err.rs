use crate::consts::{
    ErrorKind::{self, *},
    G_HDLG, OUT_TEXT,
};
use windows::{core::*, Win32::UI::WindowsAndMessaging::SetDlgItemTextW};

pub fn display_err(win_error: windows::core::Error, error_kind: ErrorKind) {
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
        SetDlgItemTextW(*G_HDLG.get().unwrap(), OUT_TEXT, display_string);
    }
}

fn display_err_with_code(win_error: Error, error_kind: ErrorKind) {
    let mut display_string = format!(
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
            AddFail=>"https://learn.microsoft.com/en-us/windows/win32/api/natupnp/nf-natupnp-istaticportmappingcollection-add",
            RemoveFail=>"https://learn.microsoft.com/en-us/windows/win32/api/natupnp/nf-natupnp-istaticportmappingcollection-remove",
            CoInitializeFail=>"https://learn.microsoft.com/en-us/windows/win32/api/objbase/nf-objbase-coinitialize",
            CoCreateInstanceFail=>"https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstance",
            StaticPortMappingCollectionFail=>"https://learn.microsoft.com/en-us/windows/win32/api/natupnp/nf-natupnp-iupnpnat-get_staticportmappingcollection",
            _=>unreachable!("コード無しエラーを予期しましたが、コード付きエラーの関数に到達しました。"),
        },
    );
    if matches!(error_kind, StaticPortMappingCollectionFail) && win_error.code() == HRESULT(0) {
        display_string.push_str("\r\nもしくはルーター側でUPnPが許可されていない可能性があります。\r\nあとパブリックネットワーク上で試行されているとたまにエラーになります。");
    }
    unsafe {
        SetDlgItemTextW(
            *G_HDLG.get().unwrap(),
            OUT_TEXT,
            PCWSTR::from_raw(HSTRING::from(display_string).as_ptr()),
        );
    }
}
