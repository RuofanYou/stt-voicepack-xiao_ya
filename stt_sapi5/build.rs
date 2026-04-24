fn main() {
    println!("cargo::rerun-if-changed=build.rs");
    if std::env::var_os("CARGO_CFG_WINDOWS").is_some() {
        println!("cargo::rerun-if-changed=\"Cargo.toml\"");
        let res = winresource::WindowsResource::new();
        res.compile().unwrap();

        // Delay-load sherpa-onnx and onnxruntime so DllMain can call
        // SetDllDirectoryW before the first real resolution. This lets
        // Windows find the runtime DLLs next to our own DLL without
        // polluting PATH or System32.
        println!("cargo::rustc-link-arg=/DELAYLOAD:sherpa-onnx-c-api.dll");
        println!("cargo::rustc-link-arg=/DELAYLOAD:onnxruntime.dll");
        println!("cargo::rustc-link-arg=delayimp.lib");
    }
}
