# PPTX to Video Converter

Convert PowerPoint presentations to videos with AI-generated voiceovers from speaker notes. Supports click-triggered animations.

## Requirements

- **Windows 10/11**
- **Microsoft PowerPoint** (Office 365, 2019, 2021)
- **Python 3.10+** - [Download](https://www.python.org/downloads/)
- **FFmpeg** - [Download](https://ffmpeg.org/download.html) or run `winget install ffmpeg`

## Quick Setup

1. **Run the setup script:**
   ```
   setup.bat
   ```

2. **Download Kokoro model files** (for free local TTS):
   - Go to: https://github.com/thewh1teagle/kokoro-onnx/releases
   - Download `kokoro-v1.0.onnx` (~330 MB)
   - Download `voices-v1.0.bin` (~52 MB)
   - Place both files in this folder

3. **Test it:**
   ```
   python convert.py test_presentation.pptx output.mp4 --provider kokoro
   ```

## Usage

### Basic Usage
```cmd
python convert.py input.pptx output.mp4 --provider kokoro
```

### With Different Voice
```cmd
python convert.py input.pptx output.mp4 --provider kokoro --voice af_bella
```

### Available Kokoro Voices
- `af_heart` (default) - American female
- `af_bella` - American female
- `af_sarah` - American female  
- `am_adam` - American male
- `am_michael` - American male
- `bf_emma` - British female
- `bm_george` - British male

### Using Cloud TTS (requires API key)

**ElevenLabs** (best quality):
```cmd
set ELEVENLABS_API_KEY=your-key
python convert.py input.pptx output.mp4 --provider elevenlabs --voice Rachel
```

**OpenAI**:
```cmd
set OPENAI_API_KEY=your-key
python convert.py input.pptx output.mp4 --provider openai --voice alloy
```

### Batch Processing
```cmd
python batch_convert.py ./presentations ./videos --provider kokoro --workers 2
```

## Animation Support

Add `[CLICK]` markers in your speaker notes to sync voiceover with animations:

```
Welcome to the presentation. [CLICK] Here's our first point. [CLICK] And our second point.
```

**How it works:**
- Text before first `[CLICK]` plays on initial slide state
- Each `[CLICK]` advances to the next animation state
- The number of `[CLICK]` markers should match your animation clicks

**Example:** A slide with 2 `[CLICK]` markers and 2 click-triggered animations:
1. Shows slide with first elements → speaks "Welcome to the presentation"
2. Click 1 animation plays → speaks "Here's our first point"
3. Click 2 animation plays → speaks "And our second point"

## Configuration

Edit settings at the top of `convert.py`:

```python
PADDING_SECONDS = 0.5          # Silence before/after each clip
NOTELESS_SLIDE_DURATION = 3.0  # Duration for slides without notes
MINIMUM_SLIDE_DURATION = 2.0   # Minimum slide duration
OUTPUT_WIDTH = 1920            # Video width
OUTPUT_HEIGHT = 1080           # Video height
```

## Files Included

| File | Description |
|------|-------------|
| `convert.py` | Main converter script |
| `powerpoint_capture.py` | PowerPoint automation module |
| `batch_convert.py` | Batch processing script |
| `create_test.py` | Creates test presentation |
| `test_presentation.pptx` | Sample presentation with animations |
| `setup.bat` | Windows setup script |
| `requirements.txt` | Python dependencies |

## Troubleshooting

### "FFmpeg not found"
- Run `winget install ffmpeg` and restart your terminal
- Or download from ffmpeg.org and add to PATH

### "Kokoro model files not found"
- Download from https://github.com/thewh1teagle/kokoro-onnx/releases
- Place `kokoro-v1.0.onnx` and `voices-v1.0.bin` in the same folder as convert.py

### "PowerPoint COM not available"
- Make sure Microsoft PowerPoint is installed
- Run `pip install pywin32`

### Animations not captured correctly
- Ensure `[CLICK]` count matches your animation clicks
- Open the PPTX in PowerPoint to verify animations work as expected

## License

MIT License - Free for personal and commercial use.
