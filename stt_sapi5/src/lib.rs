//! STT xiao_ya Chinese SAPI5 voice DLL.
//!
//! COM server implementing ISpTTSEngine, backed by sherpa-onnx (via sherpa-rs)
//! loading a `xiao_ya` VITS model packaged alongside the DLL.

use std::{
    ffi::OsString,
    os::windows::ffi::OsStringExt,
    panic::AssertUnwindSafe,
    path::PathBuf,
    sync::{Mutex, OnceLock},
    time::Instant,
};

use sherpa_rs::tts::{CommonTtsConfig, VitsTts, VitsTtsConfig};
use sherpa_rs::OnnxConfig;

use windows::{
    core::GUID,
    Win32::{
        Foundation::MAX_PATH,
        Media::{
            Audio::{WAVEFORMATEX, WAVE_FORMAT_PCM},
            Speech::{ISpObjectToken, ISpTTSEngineSite, SPVES_ABORT, SPVES_CONTINUE},
        },
        System::Registry::HKEY_LOCAL_MACHINE,
    },
};
use windows_tts_engine::{
    com_server::{
        dll_export_com_server_fns, ComClassInfo, ComServerPath, ComThreadingModel, SafeTtsComServer,
    },
    logging::DllLogger,
    utils::{get_current_dll_path, safe_catch_unwind},
    voices::{ParentRegKey, VoiceAttributes, VoiceKeyData},
    SafeTtsEngine, SpeechFormat, TextFrag, TextFragIter,
};

/// Standard E_FAIL HRESULT (0x80004005).
fn e_fail() -> windows::core::Error {
    windows::core::Error::from_hresult(windows::core::HRESULT(-2147467259_i32))
}

// xiao_ya native sample rate (from .onnx.json: 22050 Hz).
const SAMPLE_RATE: u32 = 22050;
// Sub-directory (next to the DLL) where the model bundle lives.
const MODEL_SUBDIR: &str = "xiao_ya";

/// Our unique GUID for this engine. Generated with `uuidgen`.
pub const CLSID_STT_XIAO_YA_ENGINE: GUID = GUID::from_u128(0x5C7A9D4A_3B6F_4E18_9A2D_7B1E2F8C6A31);

pub struct OurTtsEngine {
    tts: OnceLock<Mutex<VitsTts>>,
}

impl OurTtsEngine {
    fn new() -> Self {
        Self { tts: OnceLock::new() }
    }

    /// Resolve the model directory relative to the DLL's own path.
    fn model_dir() -> Option<PathBuf> {
        let mut buf = [0u16; MAX_PATH as usize];
        let path_slice = get_current_dll_path(&mut buf)
            .map_err(|e| log::error!("get_current_dll_path failed: {e}"))
            .ok()?;
        let path_slice = path_slice.strip_suffix(&[0]).unwrap_or(path_slice);
        let mut dir: PathBuf = <OsString as OsStringExt>::from_wide(path_slice).into();
        dir.pop();
        dir.push(MODEL_SUBDIR);
        if !dir.is_dir() {
            log::error!("Model directory not found: {}", dir.display());
            return None;
        }
        Some(dir)
    }

    /// 从我们 DLL 同目录用绝对路径主动 LoadLibraryW 预加载 sherpa-onnx 依赖。
    /// 必须放在第一次 speak 时（不在 DllMain loader-lock 下），避免
    /// onnxruntime 的 DllMain 在 loader lock 状态下死锁/崩。
    fn preload_runtime_dlls(dll_dir: &std::path::Path) {
        use std::os::windows::ffi::OsStrExt;
        unsafe {
            for dll_name in ["onnxruntime.dll", "sherpa-onnx-c-api.dll"] {
                let path = dll_dir.join(dll_name);
                let exists = path.is_file();
                log::info!("preload: trying {} (exists={})", path.display(), exists);
                if !exists {
                    continue;
                }
                let mut wide: Vec<u16> = path.as_os_str().encode_wide().collect();
                wide.push(0);
                let h = LoadLibraryW(PCWSTR(wide.as_ptr()));
                match h {
                    Ok(_) => log::info!("preload: {} loaded", dll_name),
                    Err(e) => log::error!("preload: {} FAILED: {e}", dll_name),
                }
            }
        }
    }

    fn build_tts() -> Option<VitsTts> {
        log::info!("build_tts: enter");
        let model_dir = Self::model_dir()?;
        log::info!("build_tts: model_dir = {}", model_dir.display());

        // 主动预加载依赖 DLL（绝对路径，从我们 DLL 同目录加载）。
        // model_dir 是 <dll_dir>/xiao_ya，所以父目录就是 dll_dir。
        if let Some(dll_dir) = model_dir.parent() {
            Self::preload_runtime_dlls(dll_dir);
        }

        let j = |name: &str| model_dir.join(name).to_string_lossy().into_owned();

        // 检查每个关键文件是否存在并 log
        for n in [
            "zh_CN-xiao_ya-medium.onnx",
            "tokens.txt",
            "lexicon.txt",
            "phone.fst",
            "date.fst",
            "number.fst",
        ] {
            let p = model_dir.join(n);
            log::info!("build_tts: file {} exists={} size={:?}",
                n, p.is_file(),
                std::fs::metadata(&p).map(|m| m.len()).ok());
        }

        // Optional FSTs (may not all exist for every xiao_ya variant).
        let mut fsts = Vec::new();
        for n in ["phone.fst", "date.fst", "number.fst", "new_heteronym.fst"] {
            let p = model_dir.join(n);
            if p.is_file() {
                fsts.push(p.to_string_lossy().into_owned());
            }
        }
        let rule_fsts = fsts.join(",");
        log::info!("build_tts: rule_fsts = {}", rule_fsts);

        let config = VitsTtsConfig {
            model: j("zh_CN-xiao_ya-medium.onnx"),
            tokens: j("tokens.txt"),
            lexicon: j("lexicon.txt"),
            dict_dir: if model_dir.join("dict").is_dir() { j("dict") } else { String::new() },
            data_dir: String::new(),
            length_scale: 1.0,
            noise_scale: 0.667,
            noise_scale_w: 0.8,
            silence_scale: 0.2,
            onnx_config: OnnxConfig {
                num_threads: 2,
                provider: "cpu".into(),
                debug: false,
            },
            tts_config: CommonTtsConfig {
                max_num_sentences: 2,
                rule_fsts,
                rule_fars: String::new(),
                silence_scale: 0.2,
            },
        };

        log::info!("build_tts: config built, calling VitsTts::new ...");
        let started = Instant::now();
        let tts = VitsTts::new(config);
        log::info!("build_tts: VitsTts::new returned in {:?}", started.elapsed());
        Some(tts)
    }

    fn with_tts<R>(&self, f: impl FnOnce(&mut VitsTts) -> R) -> Option<R> {
        // Try to initialise lazily; if init fails, return None instead of panicking.
        if self.tts.get().is_none() {
            match safe_catch_unwind(AssertUnwindSafe(Self::build_tts)).flatten() {
                Some(tts) => {
                    let _ = self.tts.set(Mutex::new(tts));
                }
                None => {
                    log::error!("build_tts failed or panicked; refusing to speak");
                    return None;
                }
            }
        }
        let slot = self.tts.get()?;
        let mut guard = slot.lock().ok()?;
        Some(f(&mut guard))
    }
}

impl SafeTtsEngine for OurTtsEngine {
    fn set_object_token(&self, _token: &ISpObjectToken) -> windows::core::Result<()> {
        log::debug!("set_object_token");
        Ok(())
    }

    #[allow(non_snake_case)]
    fn speak(
        &self,
        _token: &ISpObjectToken,
        _speak_punctuation: bool,
        _wave_format: SpeechFormat,
        text_fragments: Option<TextFrag<'_>>,
        output_site: &ISpTTSEngineSite,
    ) -> windows::core::Result<()> {
        // Catch any residual panic so that WoW never aborts because of TTS.
        let result: Option<windows::core::Result<()>> =
            safe_catch_unwind(AssertUnwindSafe(|| -> windows::core::Result<()> {
                let text_utf16 = TextFragIter::new(text_fragments)
                    .flat_map(|frag| frag.utf16_text().iter().copied().chain([' ' as u16]))
                    .collect::<Vec<u16>>();
                let text = String::from_utf16_lossy(&text_utf16);
                let text = text.trim();
                log::debug!("Speak: {}", text);

                if text.is_empty() {
                    return Ok(());
                }

                let started = Instant::now();
                let synth = self.with_tts(|tts| tts.create(text, 0, 1.0)).ok_or_else(|| {
                    log::error!("TTS engine not initialised; returning E_FAIL");
                    e_fail()
                })?;
                let audio = synth.map_err(|e| {
                    log::error!("sherpa-rs synthesize failed: {e}");
                    e_fail()
                })?;
                log::debug!(
                    "Synthesized {} samples ({} Hz) in {:?}",
                    audio.samples.len(),
                    audio.sample_rate,
                    started.elapsed()
                );

                // f32 samples (-1.0..=1.0) -> 16-bit signed little-endian PCM.
                let mut pcm_bytes = Vec::with_capacity(audio.samples.len() * 2);
                for &s in &audio.samples {
                    let clamped = s.clamp(-1.0, 1.0);
                    let sample_i16 = (clamped * 32767.0).round() as i16;
                    pcm_bytes.extend_from_slice(&sample_i16.to_le_bytes());
                }

                // Stream to output_site in 4 KB chunks, honouring abort requests.
                let mut buffer = pcm_bytes.as_slice();
                while !buffer.is_empty() {
                    let written_bytes = unsafe {
                        output_site.Write(buffer.as_ptr().cast(), buffer.len().min(4096) as u32)
                    }?;
                    buffer = &buffer[written_bytes as usize..];

                    let actions = unsafe { output_site.GetActions() } as i32;
                    if actions == SPVES_CONTINUE.0 {
                        continue;
                    }
                    if SPVES_ABORT.0 & actions != 0 {
                        return Ok(());
                    }
                }
                Ok(())
            }));

        match result {
            Some(inner) => inner,
            None => {
                log::error!("speak() panicked; returning E_FAIL instead of unwinding into WoW");
                Err(e_fail())
            }
        }
    }

    #[allow(non_snake_case)]
    fn get_output_format(
        &self,
        _token: &ISpObjectToken,
        target_format: Option<SpeechFormat>,
    ) -> windows::core::Result<SpeechFormat> {
        let result: Option<windows::core::Result<SpeechFormat>> =
            safe_catch_unwind(AssertUnwindSafe(|| -> windows::core::Result<SpeechFormat> {
                log::debug!("get_output_format: {target_format:?}");
                if let Some(SpeechFormat::DebugText) = target_format {
                    return Ok(SpeechFormat::DebugText);
                }
                let nBlockAlign: u16 = 2;
                Ok(SpeechFormat::Wave(WAVEFORMATEX {
                    wFormatTag: WAVE_FORMAT_PCM as _,
                    nChannels: 1,
                    nBlockAlign,
                    wBitsPerSample: 16,
                    nSamplesPerSec: SAMPLE_RATE,
                    nAvgBytesPerSec: SAMPLE_RATE * (nBlockAlign as u32),
                    cbSize: 0,
                }))
            }));
        match result {
            Some(inner) => inner,
            None => {
                log::error!("get_output_format() panicked; returning E_FAIL");
                Err(e_fail())
            }
        }
    }
}

fn xiao_ya_voice_data() -> VoiceKeyData {
    VoiceKeyData {
        key_name: "STT_XIAO_YA_ZH_CN".to_owned(),
        long_name: "STT 可爱少女 (xiao_ya)".to_owned(),
        class_id: CLSID_STT_XIAO_YA_ENGINE,
        attributes: VoiceAttributes {
            name: "STT 可爱少女".to_owned(),
            gender: "Female".to_owned(),
            age: "Adult".to_owned(),
            language: "804".to_owned(), // zh-CN
            vendor: "STT Voice Pack".to_owned(),
        },
    }
}

struct TtsComServer;

impl SafeTtsComServer for TtsComServer {
    const CLSID_TTS_ENGINE: GUID = CLSID_STT_XIAO_YA_ENGINE;

    type TtsEngine = OurTtsEngine;

    fn create_engine() -> Self::TtsEngine {
        OurTtsEngine::new()
    }

    fn initialize() {
        static DLL_LOGGER: DllLogger = DllLogger::new();
        DLL_LOGGER.install()
    }

    fn register_server() {
        ComClassInfo {
            clsid: CLSID_STT_XIAO_YA_ENGINE,
            class_name: Some("stt_xiao_ya_sapi5".into()),
            threading_model: ComThreadingModel::Apartment,
            server_path: ComServerPath::CurrentModule,
        }
        .register()
        .expect("Failed to register COM Class");

        let voice = xiao_ya_voice_data();
        voice
            .write_to_registry(ParentRegKey::Path(
                HKEY_LOCAL_MACHINE,
                "SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\",
            ))
            .expect("Failed to register voice (SAPI5)");
        voice
            .write_to_registry(ParentRegKey::Path(
                HKEY_LOCAL_MACHINE,
                "SOFTWARE\\Microsoft\\Speech_OneCore\\Voices\\Tokens\\",
            ))
            .expect("Failed to register voice (OneCore)");
    }

    fn unregister_server() {
        let voice = xiao_ya_voice_data();
        let _ = voice.remove_from_registry(ParentRegKey::Path(
            HKEY_LOCAL_MACHINE,
            "SOFTWARE\\Microsoft\\Speech_OneCore\\Voices\\Tokens\\",
        ));
        let _ = voice.remove_from_registry(ParentRegKey::Path(
            HKEY_LOCAL_MACHINE,
            "SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\",
        ));
        let _ = ComClassInfo::unregister_class_id(CLSID_STT_XIAO_YA_ENGINE);
    }
}

// Export SAPI5 COM entry points from the DLL.
dll_export_com_server_fns!(TtsComServer);

// ---------------------------------------------------------------------------
// DllMain: 故意只返回 TRUE。**不要** 在这里 LoadLibrary 任何东西——
// onnxruntime 自己 DllMain 在 loader-lock 下做的初始化曾导致 WoW 崩
// (0xc0000409 / FAST_FAIL_FATAL_APP_EXIT)。
// 真正的依赖预加载推迟到第一次 build_tts() 时（不在 loader lock 下）。
// ---------------------------------------------------------------------------
use windows::core::PCWSTR;
use windows::Win32::Foundation::{BOOL, TRUE};
use windows::Win32::System::LibraryLoader::LoadLibraryW;

#[no_mangle]
#[allow(non_snake_case)]
pub unsafe extern "system" fn DllMain(
    _hinst_dll: *mut core::ffi::c_void,
    _fdw_reason: u32,
    _lpv_reserved: *mut core::ffi::c_void,
) -> BOOL {
    TRUE
}
