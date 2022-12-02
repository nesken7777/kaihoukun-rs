use windows::Win32::Foundation::HWND;

pub const DIALOG: usize = 4000;
pub const PORT_NUM_INPUT: i32 = 5001;
pub const TCP_CHECKED: i32 = 5002;
pub const UDP_CHECKED: i32 = 5003;
pub const OUT_TEXT: i32 = 5004;
pub const OPEN_PORT: usize = 5005;
pub const CLOSE_PORT: usize = 5006;
pub static mut G_HDLG: HWND = HWND(0);

#[derive(PartialEq)]
pub enum ErrorKind {
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