use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use windows::Win32::System::Variant::{VARIANT, VT_BSTR};
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
    obj.GetIDsOfNames(&Default::default(), &wide_name, 1, 0, &mut dispatch_id)?;

    let bstr = BSTR::from(value);
    let mut variant = VARIANT::default();
    variant.Anonymous.Anonymous.vt = VT_BSTR;
    variant.Anonymous.Anonymous.Anonymous.bstrVal = std::mem::ManuallyDrop::new(bstr);

    let params = DISPPARAMS {
        rgvarg: &mut variant,
        cArgs: 1,
        ..Default::default()
    };
    obj.Invoke(
        dispatch_id,
        &Default::default(),
        0,
        DISPATCH_PROPERTYPUT,
        &params,
        None,
        None,
        None,
    )?;
    Ok(())
}

unsafe fn get_property(obj: &IDispatch, name: &str) -> Result<i32> {
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
            &Default::default(),
            0,
            DISPATCH_PROPERTYGET,
            &params,
            Some(&mut result),
            None,
            None,
        )?;
    }

    Ok(result.Anonymous.Anonymous.Anonymous.lVal)
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

        call_method(&obj, com_method_call.method_name, com_method_call.properties)?;

        let error_code = get_property(&obj, "ErrorCode")?;

        println!("Error Code: {}", error_code);
        CoUninitialize();
    }

    Ok(())
}
