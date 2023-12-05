use windows::{
    core::*,
    Win32::{
        NetworkManagement::WindowsFirewall::{IUPnPNAT, UPnPNAT},
        System::Com::{
            CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_INPROC_HANDLER,
            CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, CLSCTX_REMOTE_SERVER, COINIT_MULTITHREADED,
        },
    },
};

use crate::consts::{
    APICreatedError::*,
    ErrorKind::{self, *},
};

pub fn close_port(
    port_num: u16,
    protocol: &'static str,
) -> std::result::Result<(), (Error, ErrorKind)> {
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
            .Remove(port_num.into(), &BSTR::from(protocol))
            .map_err(|remove_err| (remove_err, APIE(RemoveFail)))?
    };

    unsafe {
        CoUninitialize();
    }
    Ok(())
}
