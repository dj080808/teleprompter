# 提词器 (Teleprompter)

基于 Whisper 转录格式的 KTV 风格提词器应用，支持分段录音、工程存档、预览和音频合并导出。

## 功能特点

- ✅ 解析 Whisper 转录文件（支持 `[时间-时间]` 格式）
- ✅ 自动/手动两种播放模式
- ✅ 大字体文案显示，类似 KTV 歌词效果
- ✅ 分段录音（44.1kHz 采样率）
- ✅ 支持重录任意句子
- ✅ 预览已录制音频、试听合并效果
- ✅ 工程存档与恢复（保存/加载 `project_meta.json`）
- ✅ 音频合并导出为 WAV 格式
- ✅ TTS 朗读当前句（需 `pyttsx3`）
- ✅ AI 演示播放（`ai_demos/` 目录下的 `ai_demo_*.wav`）
- ✅ 低分重录标记、静音裁剪
- ✅ 合并组、播放速度调节
- ✅ 拖拽加载（需 `tkinterdnd2`）
- ✅ 快捷键支持

## 安装依赖

**必需：**
```bash
pip install sounddevice numpy
```

**可选：**
```bash
pip install pyttsx3          # TTS 朗读
pip install tkinterdnd2      # 拖拽文件加载
```

## 使用方法

1. **运行程序：**
   ```bash
   python teleprompter.py
   ```
   > 若系统 PATH 中有 MSYS2，`python` 可能指向错误解释器，导致 `ModuleNotFoundError`。请使用完整路径：
   > `C:\Users\<用户名>\AppData\Local\Programs\Python\Python313\python.exe teleprompter.py`

2. **加载内容：**
   - 点击「📂 打开」或按 `Ctrl+O`
   - 支持加载 `.txt` 转录文件或 `project_meta.json` 工程
   - 若当前目录存在 `Mat1st_english.txt`，启动时自动加载

3. **播放模式：**
   - **手动模式**：使用方向键或导航按钮切换句子
   - **自动模式**：按时间轴自动播放，模拟 KTV 效果

4. **录音：**
   - 定位到要录制的句子
   - 点击「⏺ 开始录制」或按 `R` 键
   - 跟随提词器朗读
   - 再按 `R` 或点击「⏹ 停止录制」保存
   - 已录制的句子会显示 ✓ 标记

5. **预览：**
   - 选择已录制句子，点击「🔊 预览」或按 `P` 键

6. **导出：**
   - 录制完成后，点击「💾 导出」或按 `Enter`
   - 选择保存路径，程序合并所有片段为 WAV

7. **工程存档：**
   - 按 `Ctrl+S` 选择目录保存工程（字幕 + 音频 + 元数据）
   - 后续通过「📂 打开」选择 `project_meta.json` 恢复

## 快捷键

| 按键 | 功能 |
|------|------|
| `空格` | 播放/暂停 |
| `←` / `→` | 上一句 / 下一句 |
| `R` | 开始/停止录制 |
| `P` | 预览当前句子 |
| `Enter` | 导出音频 |
| `D` | 播放 AI 演示（`ai_demos/ai_demo_*.wav`） |
| `S` | 切换播放速度 |
| `Home` / `End` | 跳转到首句/末句 |
| `Delete` | 删除当前句录音 |
| `F1` | 帮助 |
| `F5` | 统计信息 |
| `Ctrl+S` | 保存工程 |
| `Ctrl+O` | 打开文件/工程 |

## 转录文件格式

支持的 `.txt` 格式示例：
```
[0.00-2.68]
Today, let's jump into a real mind-game match.

[2.68-4.96]
This is an eight-player free-for-all.
```

## 目录结构

- `recordings/` - 录音片段（`segment_000.wav` …）
- `ai_demos/` - AI 演示音频（`ai_demo_000.wav` …，可选）
- 工程目录：含 `project_meta.json`、字幕副本、`audio/` 子目录

## 项目文件

| 文件 | 说明 |
|------|------|
| `teleprompter.py` | 提词器主程序 |
| `parser.py` | 转录文件解析 |
| `key_remap.py` | 独立改键脚本（按 `,` 连按 `d` 100 次，需 `keyboard` 库） |
