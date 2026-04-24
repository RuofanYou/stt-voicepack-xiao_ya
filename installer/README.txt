================================================================
  STT 可爱少女 (xiao_ya) 中文 Windows 语音包
  基于 sherpa-onnx + Piper / xiao_ya (BZNSYP) / MIT
================================================================

这是什么：
---------
一个能被 Windows 系统（以及 WoW 的 STT 插件）识别为标准语音 (SAPI5 voice)
的中文女声 TTS 语音包。声音是"小雅 xiao_ya"，清亮自然。

安装步骤（3 步）：
---------
1) 解压此 zip 到任意固定目录（建议 C:\STT-xiao_ya，不要放桌面，
   因为桌面会随用户名变化）。

2) 右键 install.bat --> 以管理员身份运行
   脚本会：
     - 把 stt_xiao_ya_sapi5.dll 注册到系统 COM
     - 在 SAPI5 和 OneCore 两处语音列表各添加一条 "STT 可爱少女 (xiao_ya)"

3) 重启 WoW，在 STT 插件里选择新出现的 "STT 可爱少女" 语音。

卸载：
---------
同目录下 uninstall.bat --> 以管理员身份运行

故障排查：
---------
  - install.log / uninstall.log 在同目录，失败时请把它发给作者
  - 找不到语音？打开 PowerShell 运行：
        Add-Type -AssemblyName System.Speech
        (New-Object System.Speech.Synthesis.SpeechSynthesizer).GetInstalledVoices() |
            ForEach-Object { $_.VoiceInfo.Name }
    看看输出里是否包含 "STT 可爱少女"。
  - 声音闷/没声音？检查系统音量和 WoW 的 TTS 音量
  - DLL 注册失败？确认你是以管理员运行的
  - 在 Windows Defender 拦截？本 DLL 未签名，可手动允许

文件结构：
---------
  STT-xiao_ya-VoicePack/
    ├─ stt_xiao_ya_sapi5.dll       ← 实际的 SAPI5 引擎 DLL (~25-40 MB)
    ├─ install.bat                 ← 安装
    ├─ uninstall.bat               ← 卸载
    ├─ README.txt                  ← 本文
    └─ xiao_ya/                    ← 模型包 (必须和 DLL 保持同级)
        ├─ zh_CN-xiao_ya-medium.onnx
        ├─ tokens.txt
        ├─ lexicon.txt
        ├─ phone.fst / date.fst / number.fst / new_heteronym.fst
        └─ (可选) dict/

重要：不要单独把 DLL 拷走，模型必须和 DLL 在同一父目录。

许可：
---------
  - 本 DLL 代码：MIT/Apache-2.0
  - 参考 Lej77/windows-text-to-speech：MIT/Apache-2.0
  - sherpa-onnx 运行时：Apache-2.0
  - xiao_ya 模型：标贝 BZNSYP 数据集，仅限非商业用途
