use std::ptr::null_mut;

use windows::{
    core::*,
    Win32::{
        Foundation::BOOL,
        NetworkManagement::WindowsFirewall::{IUPnPNAT, UPnPNAT},
        Networking::WinSock::{
            GetAddrInfoExW, GetHostNameW, WSACleanup, WSAStartup, ADDRESS_FAMILY, ADDRINFOEXW,
            AF_INET, NS_DNS, SOCKADDR_IN, SOCK_RAW, WSADATA,
        },
        System::Com::{
            CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_HANDLER,
            CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, CLSCTX_REMOTE_SERVER, COINIT_MULTITHREADED,
        },
        UI::{Controls::IsDlgButtonChecked, WindowsAndMessaging::GetDlgItemInt},
    },
};

use crate::consts::{
    ErrorKind::{self, *},
    G_HDLG, PORT_NUM_INPUT, UDP_CHECKED,
};

pub fn open_port() -> std::result::Result<(), (Error, ErrorKind)> {
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
                    //Hamachiの場合以下で飛ばすようにしているが、なんか手続き的でダサい。mut Stringの方をどうにかしたい
                    if &return_str[0..2] == "25" {
                        return_str.clear();
                    }
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
            .map_err(|add_err| (add_err, AddFail))?;
    }
    unsafe {
        WSACleanup();
        CoUninitialize();
    }
    Ok(())
}