/// com_dispatch.rs
/// Windows IDispatch 기반 COM 자동화 래퍼 (windows 0.52 호환)

#[cfg(windows)]
mod platform {
    use std::mem::ManuallyDrop;
    use windows::{
        core::{BSTR, ComInterface, GUID, Interface, PCWSTR},
        Win32::{
            Foundation::VARIANT_BOOL,
            System::{
                Com::{
                    CLSIDFromProgID, CoCreateInstance, CoInitializeEx, CoUninitialize,
                    IDispatch, CLSCTX_ALL, COINIT_APARTMENTTHREADED,
                    DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
                    DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
                },
                Ole::GetActiveObject,
                Variant::{VARIANT, VARENUM, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_EMPTY, VT_I4, VT_R8},
            },
        },
    };

    pub fn com_initialize() -> anyhow::Result<()> {
        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED)
                .map_err(|e| anyhow::anyhow!("COM 초기화 실패: {e}"))
        }
    }

    pub fn com_uninitialize() {
        unsafe { CoUninitialize() };
    }

    #[derive(Debug, Clone)]
    pub enum Variant {
        Empty,
        Bool(bool),
        I32(i32),
        F64(f64),
        String(String),
        Object(ComObject),
    }

    impl Variant {
        pub fn as_string(&self) -> Option<&str> {
            if let Variant::String(s) = self { Some(s) } else { None }
        }
        pub fn as_i32(&self) -> Option<i32> {
            match self {
                Variant::I32(n) => Some(*n),
                Variant::F64(f) => Some(*f as i32),
                Variant::Bool(b) => Some(if *b { 1 } else { 0 }),
                _ => None,
            }
        }
        pub fn as_bool(&self) -> Option<bool> {
            match self {
                Variant::Bool(b) => Some(*b),
                Variant::I32(n) => Some(*n != 0),
                _ => None,
            }
        }
        pub fn as_object(self) -> Option<ComObject> {
            if let Variant::Object(o) = self { Some(o) } else { None }
        }
        pub fn to_string_repr(&self) -> String {
            match self {
                Variant::Empty => String::new(),
                Variant::Bool(b) => b.to_string(),
                Variant::I32(n) => n.to_string(),
                Variant::F64(f) => f.to_string(),
                Variant::String(s) => s.clone(),
                Variant::Object(_) => "[Object]".to_string(),
            }
        }
    }

    #[derive(Debug, Clone)]
    pub struct ComObject(pub IDispatch);

    impl ComObject {
        pub fn from_prog_id(prog_id: &str) -> anyhow::Result<Self> {
            unsafe {
                let wide: Vec<u16> = prog_id.encode_utf16().chain(std::iter::once(0)).collect();
                let clsid = CLSIDFromProgID(PCWSTR(wide.as_ptr()))
                    .map_err(|e| anyhow::anyhow!("CLSIDFromProgID 실패 ({prog_id}): {e}"))?;
                // Python win32com 호환: CLSCTX_SERVER (INPROC|LOCAL|REMOTE)
                let clsctx = windows::Win32::System::Com::CLSCTX(
                    windows::Win32::System::Com::CLSCTX_INPROC_SERVER.0
                        | windows::Win32::System::Com::CLSCTX_LOCAL_SERVER.0
                        | windows::Win32::System::Com::CLSCTX_REMOTE_SERVER.0,
                );
                let dispatch: IDispatch = CoCreateInstance(&clsid, None, clsctx)
                    .map_err(|e| anyhow::anyhow!("CoCreateInstance 실패 ({prog_id}): {e}"))?;
                Ok(ComObject(dispatch))
            }
        }

        /// 이미 실행 중인 COM 객체에 연결 (ROT → FindWindow+Accessibility 순)
        pub fn from_active_object(prog_id: &str) -> anyhow::Result<Self> {
            // 1단계: ROT에서 찾기
            if let Ok(obj) = Self::from_rot(prog_id) {
                return Ok(obj);
            }
            // 2단계: HWP 창을 직접 찾아서 IDispatch 취득
            Self::from_hwp_window()
        }

        fn from_rot(prog_id: &str) -> anyhow::Result<Self> {
            unsafe {
                let wide: Vec<u16> = prog_id.encode_utf16().chain(std::iter::once(0)).collect();
                let clsid = CLSIDFromProgID(PCWSTR(wide.as_ptr()))
                    .map_err(|e| anyhow::anyhow!("CLSIDFromProgID 실패: {e}"))?;
                let mut unknown: Option<windows::core::IUnknown> = None;
                GetActiveObject(&clsid, None, &mut unknown)
                    .map_err(|e| anyhow::anyhow!("ROT 조회 실패: {e}"))?;
                let unknown = unknown
                    .ok_or_else(|| anyhow::anyhow!("ROT에 HWP 없음"))?;
                let dispatch: IDispatch = unknown.cast()
                    .map_err(|e| anyhow::anyhow!("IDispatch 변환 실패: {e}"))?;
                Ok(ComObject(dispatch))
            }
        }

        fn from_hwp_window() -> anyhow::Result<Self> {
            use windows::Win32::UI::WindowsAndMessaging::FindWindowW;
            use windows::Win32::UI::Accessibility::AccessibleObjectFromWindow;

            unsafe {
                // 여러 클래스명으로 HWP 창 검색
                let class_names = ["HwpFrame", "HWPFrame", "HwpMainFrame"];
                let mut hwnd = windows::Win32::Foundation::HWND(0);
                let mut found_class = "";
                for name in &class_names {
                    let wide: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
                    let h = FindWindowW(PCWSTR(wide.as_ptr()), PCWSTR::null());
                    if h.0 != 0 {
                        hwnd = h;
                        found_class = name;
                        eprintln!("[HWP] FindWindow 성공: class={name}, hwnd={:?}", h.0);
                        break;
                    }
                }
                if hwnd.0 == 0 {
                    anyhow::bail!("실행 중인 HWP 창을 찾을 수 없습니다. HWP를 먼저 실행하세요.");
                }

                // OBJID_NATIVEOM (-16) 으로 네이티브 자동화 객체 취득 시도
                const OBJID_NATIVEOM: i32 = -16;
                let mut result: *mut std::ffi::c_void = std::ptr::null_mut();
                let hr = AccessibleObjectFromWindow(
                    hwnd,
                    OBJID_NATIVEOM as u32,
                    &IDispatch::IID,
                    &mut result,
                );
                if let Err(e) = hr {
                    eprintln!("[HWP] OBJID_NATIVEOM 실패 (class={found_class}): {e}");
                    // fallback: OBJID_CLIENT (기본 접근성 객체)
                    result = std::ptr::null_mut();
                    let _ = AccessibleObjectFromWindow(
                        hwnd,
                        0xFFFFFFFC, // OBJID_CLIENT
                        &IDispatch::IID,
                        &mut result,
                    );
                }

                if result.is_null() {
                    anyhow::bail!(
                        "HWP 창(class={found_class})은 찾았으나 자동화 인터페이스를 얻을 수 없습니다."
                    );
                }
                let dispatch = IDispatch::from_raw(result);
                Ok(ComObject(dispatch))
            }
        }

        pub fn call(&self, name: &str, args: Vec<Variant>) -> anyhow::Result<Variant> {
            // Python win32com 호환: DISPATCH_METHOD | DISPATCH_PROPERTYGET
            self.invoke(name, args, DISPATCH_FLAGS(DISPATCH_METHOD.0 | DISPATCH_PROPERTYGET.0))
        }

        pub fn get(&self, name: &str) -> anyhow::Result<Variant> {
            self.invoke(name, vec![], DISPATCH_PROPERTYGET)
        }

        pub fn put(&self, name: &str, value: Variant) -> anyhow::Result<()> {
            unsafe {
                let dispid = self.get_dispid(name)?;
                let mut raw = variant_to_raw(&value);
                let mut put_id: i32 = -3; // DISPID_PROPERTYPUT

                let dp = DISPPARAMS {
                    rgvarg: &mut raw,
                    rgdispidNamedArgs: &mut put_id,
                    cArgs: 1,
                    cNamedArgs: 1,
                };
                let mut ei = EXCEPINFO::default();
                let mut ae = 0u32;

                self.0.Invoke(dispid, &GUID::zeroed(), 0, DISPATCH_PROPERTYPUT,
                    &dp, None, Some(&mut ei), Some(&mut ae))
                    .map_err(|e| anyhow::anyhow!("put '{name}' 실패: {e}"))?;

                drop_variant_bstr(&mut raw);
                Ok(())
            }
        }

        fn get_dispid(&self, name: &str) -> anyhow::Result<i32> {
            unsafe {
                let wide: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
                let mut id = 0i32;
                self.0.GetIDsOfNames(&GUID::zeroed(), &PCWSTR(wide.as_ptr()), 1, 0, &mut id)
                    .map_err(|e| anyhow::anyhow!("GetIDsOfNames '{name}' 실패: {e}"))?;
                Ok(id)
            }
        }

        fn invoke(&self, name: &str, mut args: Vec<Variant>, flags: DISPATCH_FLAGS) -> anyhow::Result<Variant> {
            unsafe {
                let dispid = self.get_dispid(name)?;
                args.reverse(); // COM은 역순
                let mut raws: Vec<VARIANT> = args.iter().map(|v| variant_to_raw(v)).collect();

                let dp = DISPPARAMS {
                    rgvarg: if raws.is_empty() { std::ptr::null_mut() } else { raws.as_mut_ptr() },
                    rgdispidNamedArgs: std::ptr::null_mut(),
                    cArgs: raws.len() as u32,
                    cNamedArgs: 0,
                };
                let mut result = VARIANT::default();
                let mut ei = EXCEPINFO::default();
                let mut ae = 0u32;

                self.0.Invoke(dispid, &GUID::zeroed(), 0, flags,
                    &dp, Some(&mut result), Some(&mut ei), Some(&mut ae))
                    .map_err(|e| {
                        let desc = if (*ei.bstrDescription).is_empty() {
                            e.to_string()
                        } else {
                            format!("{e}: {}", &*ei.bstrDescription)
                        };
                        anyhow::anyhow!("Invoke '{name}' 실패: {desc}")
                    })?;

                for r in &mut raws { drop_variant_bstr(r); }
                raw_to_variant(result)
            }
        }
    }

    unsafe fn variant_to_raw(v: &Variant) -> VARIANT {
        let mut var = VARIANT::default();
        let inner = &mut *var.Anonymous.Anonymous;
        match v {
            Variant::Empty => { inner.vt = VT_EMPTY; }
            Variant::Bool(b) => {
                inner.vt = VT_BOOL;
                inner.Anonymous.boolVal = VARIANT_BOOL(if *b { -1 } else { 0 });
            }
            Variant::I32(n) => {
                inner.vt = VT_I4;
                inner.Anonymous.lVal = *n;
            }
            Variant::F64(f) => {
                inner.vt = VT_R8;
                inner.Anonymous.dblVal = *f;
            }
            Variant::String(s) => {
                inner.vt = VT_BSTR;
                inner.Anonymous.bstrVal = ManuallyDrop::new(BSTR::from(s.as_str()));
            }
            Variant::Object(obj) => {
                inner.vt = VT_DISPATCH;
                inner.Anonymous.pdispVal = ManuallyDrop::new(Some(obj.0.clone()));
            }
        }
        var
    }

    unsafe fn raw_to_variant(var: VARIANT) -> anyhow::Result<Variant> {
        let inner = &*var.Anonymous.Anonymous;
        let vt = inner.vt;
        Ok(if vt == VT_EMPTY {
            Variant::Empty
        } else if vt == VT_BOOL {
            Variant::Bool(inner.Anonymous.boolVal.0 != 0)
        } else if vt == VT_I4 {
            Variant::I32(inner.Anonymous.lVal)
        } else if vt == VT_R8 {
            Variant::F64(inner.Anonymous.dblVal)
        } else if vt == VT_BSTR {
            Variant::String((*inner.Anonymous.bstrVal).to_string())
        } else if vt == VT_DISPATCH {
            match &*inner.Anonymous.pdispVal {
                Some(d) => Variant::Object(ComObject(d.clone())),
                None => Variant::Empty,
            }
        } else {
            match vt {
                VARENUM(2)  => Variant::I32(inner.Anonymous.iVal as i32),    // VT_I2
                VARENUM(18) => Variant::I32(inner.Anonymous.uiVal as i32),   // VT_UI2
                VARENUM(19) => Variant::I32(inner.Anonymous.ulVal as i32),   // VT_UI4
                VARENUM(4)  => Variant::F64(inner.Anonymous.fltVal as f64),  // VT_R4
                VARENUM(10) => { // VT_ERROR - return empty
                    Variant::Empty
                }
                _ => Variant::Empty,
            }
        })
    }

    unsafe fn drop_variant_bstr(var: &mut VARIANT) {
        let inner = &mut *var.Anonymous.Anonymous;
        if inner.vt == VT_BSTR {
            let _ = ManuallyDrop::take(&mut inner.Anonymous.bstrVal);
            inner.vt = VT_EMPTY;
        }
    }
}

// ── Non-Windows stub ─────────────────────────────────────────

#[cfg(not(windows))]
mod platform {
    #[derive(Debug, Clone)]
    pub enum Variant {
        Empty, Bool(bool), I32(i32), F64(f64), String(String), Object(ComObject),
    }
    impl Variant {
        pub fn as_string(&self) -> Option<&str> { None }
        pub fn as_i32(&self) -> Option<i32> { None }
        pub fn as_bool(&self) -> Option<bool> { None }
        pub fn as_object(self) -> Option<ComObject> { None }
        pub fn to_string_repr(&self) -> std::string::String { std::string::String::new() }
    }
    #[derive(Debug, Clone)]
    pub struct ComObject;
    impl ComObject {
        pub fn from_prog_id(_: &str) -> anyhow::Result<Self> {
            anyhow::bail!("COM automation은 Windows 전용입니다.")
        }
        pub fn from_active_object(_: &str) -> anyhow::Result<Self> {
            anyhow::bail!("COM automation은 Windows 전용입니다.")
        }
        pub fn call(&self, _: &str, _: Vec<Variant>) -> anyhow::Result<Variant> {
            anyhow::bail!("COM automation은 Windows 전용입니다.")
        }
        pub fn get(&self, _: &str) -> anyhow::Result<Variant> {
            anyhow::bail!("COM automation은 Windows 전용입니다.")
        }
        pub fn put(&self, _: &str, _: Variant) -> anyhow::Result<()> {
            anyhow::bail!("COM automation은 Windows 전용입니다.")
        }
    }
    pub fn com_initialize() -> anyhow::Result<()> { Ok(()) }
    pub fn com_uninitialize() {}
}

pub use platform::{com_initialize, com_uninitialize, ComObject, Variant};
