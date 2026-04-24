//! STT xiao_ya Chinese SAPI5 voice DLL.
//!
//! COM server implementing ISpTTSEngine, backed by sherpa-onnx (via sherpa-rs)
//! loading a `xiao_ya` VITS model packaged alongside the DLL.

use std::{
    ffi::OsString,
    os::windows::ffi::OsStringExt,
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
    utils::get_current_dll_path,
    voices::{ParentRegKey, VoiceAttributes, VoiceKeyData},
    SafeTtsEngine, SpeechFormat, TextFrag, TextFragIter,
};

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

    fn build_tts() -> Option<VitsTts> {
        let dir = Self::model_dir()?;
        let j = |name: &str| dir.join(name).to_string_lossy().into_owned();

        // Optional FSTs (may not all exist for every xiao_ya variant).
        let mut fsts = Vec::new();
        for n in ["phone.fst", "date.fst", "number.fst", "new_heteronym.fst"] {
            let p = dir.join(n);
            if p.is_file() {
                fsts.push(p.to_string_lossy().into_owned());
            }
        }
        let rule_fsts = fsts.join(",");

        let config = VitsTtsConfig {
            model: j("zh_CN-xiao_ya-medium.onnx"),
            tokens: j("tokens.txt"),
            lexicon: j("lexicon.txt"),
            dict_dir: if dir.join("dict").is_dir() { j("dict") } else { String::new() },
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

        let started = Instant::now();
        let tts = VitsTts::new(config);
        log::info!("VitsTts initialized in {:?}", started.elapsed());
        Some(tts)
    }

    fn with_tts<R>(&self, f: impl FnOnce(&mut VitsTts) -> R) -> Option<R> {
        let slot = self.tts.get_or_init(|| {
            let tts = Self::build_tts().expect("Failed to build xiao_ya VitsTts");
            Mutex::new(tts)
        });
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
        let audio = self
            .with_tts(|tts| tts.create(text, 0, 1.0))
            .expect("tts init")
            .map_err(|e| {
                log::error!("sherpa-rs synthesize failed: {e}");
                windows::core::Error::from_hresult(windows::core::HRESULT(-2147467259_i32))
            })?;
        log::debug!(
            "Synthesized {} samples ({} Hz) in {:?}",
            audio.samples.len(),
            audio.sample_rate,
            started.elapsed()
        );

        // Convert f32 samples (-1.0..=1.0) to 16-bit signed little-endian PCM bytes.
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
    }

    #[allow(non_snake_case)]
    fn get_output_format(
        &self,
        _token: &ISpObjectToken,
        target_format: Option<SpeechFormat>,
    ) -> windows::core::Result<SpeechFormat> {
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
