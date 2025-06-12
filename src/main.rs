use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::HashMap;
use std::io::{self, Read};
use windows::Win32::System::Ole::DISPID_PROPERTYPUT;
use windows::Win32::System::Variant::{VARIANT, VariantToString};
use windows::{Win32::System::Com::*, core::*};

#[derive(Serialize, Deserialize)]
struct ComMethodCall {
    version: String,
    prog_id: String,
    method: String,
    properties: HashMap<String, Value>,
}

fn to_pcwstr(s: &str) -> PCWSTR {
    let wide: Vec<u16> = s.encode_utf16().chain(std::iter::once(0)).collect();
    PCWSTR::from_raw(wide.as_ptr())
}

unsafe fn value_to_variant(value: &Value) -> VARIANT {
    match value {
        Value::String(s) => VARIANT::from(BSTR::from(s.as_str())),
        Value::Number(n) => {
            // Prioritize integer conversion if possible
            if n.is_i64() {
                // Assuming i32 is sufficient for integer properties.
                // If larger integers are expected, consider VT_I8 or custom handling.
                n.as_i64()
                    .map_or(VARIANT::default(), |i| VARIANT::from(i as i32))
            } else if n.is_f64() {
                // Handle floating-point numbers
                n.as_f64().map_or(VARIANT::default(), |f| VARIANT::from(f))
            } else {
                // Fallback for numbers that don't fit i64 or f64 (e.g., very large BigInts)
                eprintln!(
                    "Warning: Unsupported number type in JSON, defaulting to empty VARIANT. \
                    Value: {n}"
                );
                VARIANT::default()
            }
        }
        Value::Bool(b) => VARIANT::from(*b),
        Value::Null => {
            eprintln!("Warning: Unable to set NULL as a VARIANT");
            VARIANT::default()
        }
        Value::Array(_) => {
            eprintln!(
                "Warning: JSON Array type is not directly supported for simple VARIANT conversion \
                for property setting. Defaulting to empty VARIANT."
            );
            VARIANT::default()
        }
        Value::Object(_) => {
            eprintln!(
                "Warning: JSON Object type is not directly supported for simple VARIANT conversion \
                for property setting. Defaulting to empty VARIANT."
            );
            VARIANT::default()
        }
    }
}

unsafe fn set_property(obj: &IDispatch, name: &str, value: &Value) -> Result<()> {
    let wide_name = to_pcwstr(name);
    let mut dispatch_id = Default::default();

    // Get the DISPID for the property name
    unsafe {
        obj.GetIDsOfNames(&Default::default(), &wide_name, 1, 0, &mut dispatch_id)?;
    }

    let mut dispid_put = DISPID_PROPERTYPUT; // Special DISPID for property put operations

    // Convert serde_json::Value to VARIANT
    unsafe {
        let mut variant_value = value_to_variant(value);

        // Prepare DISPPARAMS for setting a property
        let params = DISPPARAMS {
            rgvarg: &mut variant_value,         // The value to set
            rgdispidNamedArgs: &mut dispid_put, // Indicates this is a property put
            cArgs: 1,                           // One argument (the value)
            cNamedArgs: 1,                      // One named argument (DISPID_PROPERTYPUT)
        };

        // Invoke the property put operation
        obj.Invoke(
            dispatch_id,          // DISPID of the property
            &GUID::zeroed(),      // Reserved, must be IID_NULL for Invoke
            0,                    // Locale ID (LOCALE_USER_DEFAULT)
            DISPATCH_PROPERTYPUT, // Flag indicating a property put
            &params,              // Parameters for the invocation
            None,                 // No return value expected for property put
            None,                 // No exception info needed
            None,                 // No argument error info needed
        )?;
    }

    Ok(())
}

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
    properties: HashMap<String, Value>,
) -> Result<()> {
    for (prop_name, prop_value) in properties {
        println!("Setting property: {prop_name} = {prop_value:?}");

        unsafe {
            set_property(obj, &prop_name, &prop_value)?;
        }
    }

    let wide_name = to_pcwstr(name.as_str());
    let mut dispatch_id = Default::default();

    // Get the DISPID for the method name
    unsafe {
        obj.GetIDsOfNames(&Default::default(), &wide_name, 1, 0, &mut dispatch_id)?;
    }

    let mut variant_result = VARIANT::default(); // For potential return value of the method
    let params = DISPPARAMS {
        rgvarg: &mut variant_result, // If the method returns a value, it would be stored here.
        cArgs: 0,                    // No arguments passed to the method itself
        ..Default::default()
    };

    println!("Calling method: {name}");
    // Invoke the method
    unsafe {
        obj.Invoke(
            dispatch_id,     // DISPID of the method
            &GUID::zeroed(), // Reserved, must be IID_NULL for Invoke
            0,               // Locale ID (LOCALE_USER_DEFAULT)
            DISPATCH_METHOD, // Flag indicating a method call
            &params,         // Parameters for the invocation
            None, // No return value needed to be captured here (already in pVarResult if provided)
            None, // No exception info needed
            None, // No argument error info needed
        )?;
    }

    Ok(())
}

fn get_data_from_stdio() -> String {
    let mut buffer = String::new();
    io::stdin()
        .read_to_string(&mut buffer)
        .expect("Failed to read from stdin");

    buffer
}
fn get_call_params_from_json_buffer(buffer: String) -> ComMethodCall {
    let com_method_call: ComMethodCall =
        serde_json::from_str(&buffer).expect("Failed to deserialize ComMethodCall JSON");

    com_method_call
}

fn execute(com_method_call: ComMethodCall) -> Result<()> {
    unsafe {
        let _ = CoInitialize(None);
        let prog_id = to_pcwstr(com_method_call.prog_id.as_str());
        let clsid = CLSIDFromProgID(prog_id)?;
        let obj: IDispatch = CoCreateInstance(&clsid, None, CLSCTX_ALL)?;

        call_method(&obj, com_method_call.method, com_method_call.properties)?;

        let error_code = get_property(&obj, "ErrorCode")?;

        println!("Error Code: {error_code}");
        CoUninitialize();
    }
    
    Ok(())
}

fn main() -> Result<()> {
    let buffer = get_data_from_stdio();
    let com_method_call = get_call_params_from_json_buffer(buffer);
    let result = execute(com_method_call);
    
    result
}
