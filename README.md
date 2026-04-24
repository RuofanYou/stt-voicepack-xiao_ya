# STT-VoicePack-xiao_ya

Windows SAPI5 voice DLL for Chinese **xiao_ya** voice, backed by sherpa-onnx.
Designed to be installed on Windows machines so WoW's STT addon can list and use
"STT 可爱少女 (xiao_ya)" as a TTS voice.

## How this works

- **`stt_sapi5/`** — Rust crate producing `stt_xiao_ya_sapi5.dll`, a SAPI5 COM
  server that implements `ISpTTSEngine`. Backend is
  [`sherpa-rs`](https://github.com/thewh1teagle/sherpa-rs) running a Piper
  `zh_CN-xiao_ya-medium` ONNX model.
- **`stt_sapi5/windows_tts_engine/`** — Lej77's SAPI5 COM scaffolding (copied in,
  MIT/Apache-2.0). Handles `DllGetClassObject`, `DllRegisterServer`, voice-token
  registry writing.
- **`.github/workflows/build.yml`** — GitHub Actions pipeline that, on push,
  compiles the DLL on a Windows runner and bundles the final distributable zip.
- **`installer/`** — `install.bat`, `uninstall.bat`, `README.txt` copied into
  the zip for end users.

## Final artifact

A single ~60-100 MB zip: **`STT-xiao_ya-VoicePack.zip`**, containing

```
STT-xiao_ya-VoicePack/
├── stt_xiao_ya_sapi5.dll      # our COM server
├── xiao_ya/                   # sherpa-onnx Piper zh_CN xiao_ya model bundle
├── install.bat                # run as admin to register
├── uninstall.bat              # run as admin to deregister
└── README.txt
```

User workflow: unzip → right-click `install.bat` → Run as administrator →
reboot WoW → select "STT 可爱少女 (xiao_ya)" in STT addon.

## Build locally (Windows only)

```powershell
rustup target add x86_64-pc-windows-msvc
cargo build --release --package stt_sapi5_xiao_ya --target x86_64-pc-windows-msvc
```

Cross-compiling from Mac is not supported because `sherpa-rs-sys` compiles
`sherpa-onnx` via CMake and needs the native MSVC toolchain. Use GitHub Actions.

## Licenses

- This code: MIT OR Apache-2.0
- `windows_tts_engine` (Lej77): MIT OR Apache-2.0
- `sherpa-onnx` / `sherpa-rs`: Apache-2.0 / MIT
- `xiao_ya` voice model: trained on BZNSYP (non-commercial research dataset)
