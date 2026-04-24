fn main() {
    println!("cargo::rerun-if-changed=build.rs");
    if std::env::var_os("CARGO_CFG_WINDOWS").is_some() {
        println!("cargo::rerun-if-changed=\"Cargo.toml\"");
        let res = winresource::WindowsResource::new();
        res.compile().unwrap();
    }
}
