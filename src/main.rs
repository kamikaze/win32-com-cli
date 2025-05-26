use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use windows::Win32::System::Ole::DISPID_PROPERTYPUT;
use windows::Win32::System::Variant::{VARIANT, VariantToString};
use windows::{Win32::System::Com::*, core::*};

#[derive(Serialize, Deserialize)]
struct ComMethodCall {
    version: String,
    prog_id: String,
    method_name: String,
    properties: HashMap<String, String>,
}

fn to_pcwstr(s: &str) -> PCWSTR {
    let wide: Vec<u16> = s.encode_utf16().chain(std::iter::once(0)).collect();
    PCWSTR::from_raw(wide.as_ptr())
}

unsafe fn set_property(obj: &IDispatch, name: &str, value: &str) -> Result<()> {
    let wide_name = to_pcwstr(name);
    let mut dispatch_id = Default::default();

    unsafe {
        obj.GetIDsOfNames(&Default::default(), &wide_name, 1, 0, &mut dispatch_id)?;
    }

    let mut variant_value = VARIANT::from(BSTR::from(value));
    let mut dispid_put = DISPID_PROPERTYPUT;
    let params = DISPPARAMS {
        rgvarg: &mut variant_value,
        rgdispidNamedArgs: &mut dispid_put,
        cArgs: 1,
        cNamedArgs: 1,
    };

    unsafe {
        obj.Invoke(
            dispatch_id,
            &GUID::zeroed(), // Reserved, must be IID_NULL
            0,
            DISPATCH_PROPERTYPUT,
            &params,
            None, // Not needed for a property put
            None, // Not needed
            None, // Not needed
        )?;
    }

    Ok(())
}

/// Gets a property from a COM object as a String using IDispatch.
///
/// # Safety
///
/// This function is unsafe because it interacts with COM, which involves raw pointers
/// and adheres to a specific, unforgiving interface contract. The caller must ensure that:
/// 1. `obj` is a valid `IDispatch` pointer.
/// 2. The property identified by `name` exists and its value can be represented as a string.
unsafe fn get_property(obj: &IDispatch, name: &str) -> Result<String> {
    let wide_name = to_pcwstr(name);
    let mut dispatch_id = Default::default();

    unsafe {
        obj.GetIDsOfNames(&Default::default(), &wide_name, 1, 0, &mut dispatch_id)?;
    }

    let params = DISPPARAMS::default();
    let mut result = VARIANT::default();

    unsafe {
        obj.Invoke(
            dispatch_id,
            &GUID::zeroed(), // Reserved, must be IID_NULL
            0,               // Use system default locale
            DISPATCH_PROPERTYGET,
            &params,
            Some(&mut result),
            None,
            None,
        )?;
    }

    let bstr_val = BSTR::default();

    unsafe {
        VariantToString(&result, &mut bstr_val.to_vec())?;
    }

    Ok(bstr_val.to_string())
}

unsafe fn call_method(
    obj: &IDispatch,
    name: String,
    properties: HashMap<String, String>,
) -> Result<()> {
    let wide_name = to_pcwstr(name.as_str());
    let mut dispatch_id = Default::default();
    let mut variant = VARIANT::default();
    let params = DISPPARAMS {
        rgvarg: &mut variant,
        cArgs: 1,
        ..Default::default()
    };

    unsafe {
        obj.GetIDsOfNames(&Default::default(), &wide_name, 1, 0, &mut dispatch_id)?;
        obj.Invoke(
            dispatch_id,
            &Default::default(),
            0,
            DISPATCH_METHOD,
            &params,
            None,
            None,
            None,
        )?;
    }

    Ok(())
}

fn main() -> Result<()> {
    let data = r#"
        {
            "version": "1",
            "prog_id": "ECR2ATL.ECR2Transaction",
            "method": "Cancellation",
            "properties": {
                "ECRNameAndVersion": "CL E-kvits Ver. 2025.5.22",
                "ReqInvoiceNumber": "CL12345",
                "ReqDateTime": "2025-05-22 12:33:44",
            }
        }"#;
    let com_method_call: ComMethodCall =
        serde_json::from_str(data).expect("Failed to deserialize ComMethodCall JSON");

    unsafe {
        let _ = CoInitialize(None);
        let prog_id = to_pcwstr(com_method_call.prog_id.as_str());
        let clsid = CLSIDFromProgID(prog_id)?;
        let obj: IDispatch = CoCreateInstance(&clsid, None, CLSCTX_ALL)?;

        call_method(
            &obj,
            com_method_call.method_name,
            com_method_call.properties,
        )?;

        let error_code = get_property(&obj, "ErrorCode")?;

        println!("Error Code: {}", error_code);
        CoUninitialize();
    }

    Ok(())
}
