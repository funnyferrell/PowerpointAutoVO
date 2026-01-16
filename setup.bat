@echo off
echo ============================================
echo  PPTX to Video Converter - Setup
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.10+ from python.org
    pause
    exit /b 1
)

echo [1/4] Installing Python dependencies...
pip install python-pptx pywin32 requests soundfile kokoro-onnx

echo.
echo [2/4] Checking FFmpeg...
ffmpeg -version >nul 2>&1
if errorlevel 1 (
    echo.
    echo WARNING: FFmpeg not found in PATH.
    echo.
    echo Please install FFmpeg:
    echo   Option A: Run "winget install ffmpeg" then restart your terminal
    echo   Option B: Download from https://ffmpeg.org/download.html
    echo            Extract to C:\ffmpeg and add C:\ffmpeg\bin to your PATH
    echo.
    pause
) else (
    echo FFmpeg found.
)

echo.
echo [3/4] Checking for Kokoro model files...
if exist "kokoro-v1.0.onnx" (
    if exist "voices-v1.0.bin" (
        echo Kokoro model files found.
    ) else (
        echo WARNING: voices-v1.0.bin not found.
        goto :download_models
    )
) else (
    echo WARNING: kokoro-v1.0.onnx not found.
    goto :download_models
)
goto :skip_download

:download_models
echo.
echo To use Kokoro TTS (free, local), download these files:
echo   https://github.com/thewh1teagle/kokoro-onnx/releases
echo.
echo   - kokoro-v1.0.onnx (~330 MB)
echo   - voices-v1.0.bin (~52 MB)
echo.
echo Place them in this folder: %CD%
echo.
echo Alternatively, use cloud TTS (requires API key):
echo   --provider elevenlabs
echo   --provider openai
echo   --provider azure
echo.

:skip_download

echo.
echo [4/4] Setup complete!
echo.
echo ============================================
echo  Quick Start
echo ============================================
echo.
echo Test the installation:
echo   python convert.py test_presentation.pptx output.mp4 --provider kokoro
echo.
echo With ElevenLabs:
echo   set ELEVENLABS_API_KEY=your-key
echo   python convert.py input.pptx output.mp4 --provider elevenlabs
echo.
pause
