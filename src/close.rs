use windows::{
    core::*,
    Win32::{
        Foundation::BOOL,
        NetworkManagement::WindowsFirewall::{IUPnPNAT, UPnPNAT},
        System::Com::{
            CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_HANDLER,
            CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, CLSCTX_REMOTE_SERVER, COINIT_MULTITHREADED,
        },
        UI::{Controls::IsDlgButtonChecked, WindowsAndMessaging::GetDlgItemInt},
    },
};

use crate::consts::{
    APICreatedError::*,
    ErrorKind::{self, *},
    SelfCreatedError::*,
    G_HDLG, PORT_NUM_INPUT, UDP_CHECKED,
};

pub fn close_port() -> std::result::Result<(), (Error, ErrorKind)> {
    let mut get_port_check = BOOL::default();
    let port_num = unsafe {
        GetDlgItemInt(
            *G_HDLG.get().unwrap(),
            PORT_NUM_INPUT,
            Some(&mut get_port_check),
            BOOL::from(false),
        )
    };
    if get_port_check == BOOL::from(false) || port_num <= 0 || port_num > 65535 {
        return Err((Error::OK, SelfE(InvalidPortNumber)));
    }
    let udp_check = unsafe { IsDlgButtonChecked(*G_HDLG.get().unwrap(), UDP_CHECKED) };
    let tcp_or_udp_string = if udp_check != 1 { "TCP" } else { "UDP" };
    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED)
            .map_err(|co_init_err| (co_init_err, APIE(CoInitializeFail)))?;
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
    unsafe {
        static_port_mapping_collection
            .Remove(port_num as i32, &BSTR::from(tcp_or_udp_string))
            .map_err(|remove_err| (remove_err, APIE(RemoveFail)))?
    };

    unsafe {
        CoUninitialize();
    }
    Ok(())
}
