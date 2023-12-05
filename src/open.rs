use std::{mem::transmute, net::Ipv4Addr, ptr::null_mut};

use windows::{
    core::*,
    Win32::{
        Foundation::VARIANT_TRUE,
        NetworkManagement::WindowsFirewall::{IUPnPNAT, UPnPNAT},
        Networking::WinSock::{
            GetAddrInfoExW, GetHostNameW, WSACleanup, WSAStartup, ADDRINFOEXW, AF_INET, NS_DNS,
            SOCKADDR_IN, SOCK_RAW, WSADATA,
        },
        System::Com::{
            CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_HANDLER,
            CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, CLSCTX_REMOTE_SERVER, COINIT_MULTITHREADED,
        },
    },
};

use crate::consts::{
    APICreatedError::*,
    ErrorKind::{self, *},
    SelfCreatedError::*,
};

pub fn open_port(
    port_num: u16,
    protocol: &'static str,
) -> std::result::Result<(), (Error, ErrorKind)> {
    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED)
            .map_err(|co_init_err| (co_init_err, APIE(CoInitializeFail)))?
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
        .map_err(|co_create_instance_err| (co_create_instance_err, APIE(CoCreateInstanceFail)))?
    };
    let static_port_mapping_collection = unsafe {
        upnp_nat
            .StaticPortMappingCollection()
            .map_err(|static_port_mapping_collection_err| {
                (
                    static_port_mapping_collection_err,
                    APIE(StaticPortMappingCollectionFail),
                )
            })?
    };

    let ip_str = determine_ip()?;

    unsafe {
        static_port_mapping_collection
            .Add(
                port_num.into(),
                &BSTR::from(protocol),
                port_num.into(),
                &BSTR::from(ip_str),
                VARIANT_TRUE,
                &BSTR::from("kaihoukun"),
            )
            .map_err(|add_err| (add_err, APIE(AddFail)))?;
    }
    unsafe {
        WSACleanup();
        CoUninitialize();
    }
    Ok(())
}

fn determine_ip() -> std::result::Result<String, (Error, ErrorKind)> {
    let mut wsa_data = WSADATA::default();
    let wsaresult = unsafe { WSAStartup(0x202, &mut wsa_data) };
    if wsaresult != 0 {
        return Err((Error::OK, SelfE(WSAStartupFail)));
    }

    let mut localhost_name = [0u16; 260];
    let gethostname_result = unsafe { GetHostNameW(localhost_name.as_mut_slice()) };
    if gethostname_result != 0 {
        return Err((Error::OK, SelfE(GetHostNameWFail)));
    }

    let hints = ADDRINFOEXW {
        ai_family: AF_INET.0 as i32,
        ai_socktype: SOCK_RAW.0,
        ..Default::default()
    };
    let mut addr_info: *mut ADDRINFOEXW = null_mut();
    let wsa_error_code = unsafe {
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
    if wsa_error_code != 0 {
        return Err((Error::OK, SelfE(GetAddrInfoExWFail)));
    }

    let ip_str = {
        fn determine_ip(
            ptr: Option<&ADDRINFOEXW>,
        ) -> std::result::Result<String, (Error, ErrorKind)> {
            match ptr {
                Some(addr_info) => {
                    let sockaddr = unsafe {
                        *addr_info
                            .ai_addr
                            .as_ref()
                            .expect("結果のai_addrがnullになることは無いと思う")
                    };
                    match sockaddr.sa_family {
                        AF_INET => {
                            let return_str = unsafe {
                                Ipv4Addr::from(transmute::<_, SOCKADDR_IN>(sockaddr).sin_addr)
                                    .to_string()
                            };
                            if &return_str[0..2] == "25" {
                                determine_ip(unsafe { (addr_info.ai_next).as_ref() })
                            } else {
                                Ok(return_str)
                            }
                        }
                        _ => Err((Error::OK, SelfE(IPNotFound))),
                    }
                }
                None => Err((Error::OK, SelfE(IPNotFound))),
            }
        }
        determine_ip(unsafe { addr_info.as_ref() })
    };
    ip_str
}
