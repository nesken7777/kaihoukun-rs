use crate::consts::{
    APICreatedError::{self, *},
    ErrorKind::{self, *},
    SelfCreatedError::{self, *},
    OUT_TEXT,
};
use windows::{
    core::*,
    Win32::{Foundation::HWND, UI::WindowsAndMessaging::SetDlgItemTextW},
};

pub fn display_err(error: windows::core::Error, kind: ErrorKind, window_handle: HWND) {
    match kind {
        SelfE(kind) => display_err_without_code(kind, window_handle),
        ApiE(kind) => display_err_with_code(error, kind, window_handle),
    }
}

fn display_err_without_code(kind: SelfCreatedError, window_handle: HWND) {
    let display_string = match kind {
        InvalidPortNumber => w!("ポート番号が不正です"),
        WSAStartupFail => w!("WSAStartup関数失敗です。"),
        GetHostNameWFail => w!("GetHostNameW関数失敗です。"),
        GetAddrInfoExWFail => w!("GetAddrInfoExW関数失敗です。"),
        IPNotFound => w!("IPアドレスが見つかりませんでした。"),
    };
    unsafe {
        let _ = SetDlgItemTextW(window_handle, OUT_TEXT, display_string);
    }
}

fn display_err_with_code(error: Error, kind: APICreatedError, window_handle: HWND) {
    let mut display_string = format!(
        "{}\r\n\r\nエラーコード:{:?}\r\n({})\r\n詳細は{}を参照されたし。",
        match kind{
            AddFail=>"ポート開放に失敗しました。",
            RemoveFail=>"ポートが閉じれませんでした。開いていますか？",
            CoInitializeFail=>"CoInitializeに失敗しました。",
            CoCreateInstanceFail=>"CoCreateInstanceに失敗しました。",
            StaticPortMappingCollectionFail=>"StaticPortMappingCollectionに失敗しました。",
        },
        error.code(),
        error.message().to_string().trim(),
        match kind{
            AddFail=>"https://learn.microsoft.com/en-us/windows/win32/api/natupnp/nf-natupnp-istaticportmappingcollection-add",
            RemoveFail=>"https://learn.microsoft.com/en-us/windows/win32/api/natupnp/nf-natupnp-istaticportmappingcollection-remove",
            CoInitializeFail=>"https://learn.microsoft.com/en-us/windows/win32/api/objbase/nf-objbase-coinitialize",
            CoCreateInstanceFail=>"https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstance",
            StaticPortMappingCollectionFail=>"https://learn.microsoft.com/en-us/windows/win32/api/natupnp/nf-natupnp-iupnpnat-get_staticportmappingcollection",
        },
    );
    if matches!(kind, StaticPortMappingCollectionFail) && error.code() == HRESULT(0) {
        display_string.push_str("\r\nルーター側でUPnPが許可されていない可能性があります。\r\nあとパブリックネットワーク上で試行されているとたまにエラーになります。");
    }
    unsafe {
        let _ = SetDlgItemTextW(
            window_handle,
            OUT_TEXT,
            PCWSTR::from_raw(HSTRING::from(display_string).as_ptr()),
        );
    }
}
