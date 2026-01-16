#!/usr/bin/env python3
"""
PPTX to Video Converter with TTS Voiceovers

Converts PowerPoint presentations to videos with AI-generated voiceovers
synced to speaker notes. Supports animation capture on Windows via PowerPoint
COM automation.

Requirements:
    Windows:
        - Microsoft PowerPoint installed
        - pip install pywin32 pillow python-pptx requests
        - FFmpeg in PATH
    
    TTS (choose one):
        - ElevenLabs: Set ELEVENLABS_API_KEY environment variable
        - OpenAI: Set OPENAI_API_KEY environment variable
        - Azure: Set AZURE_SPEECH_KEY and AZURE_SPEECH_REGION
        - Kokoro: pip install kokoro-onnx soundfile (local GPU)

Usage:
    python convert.py input.pptx output.mp4
    python convert.py input.pptx output.mp4 --provider elevenlabs --voice Rachel
    python convert.py input.pptx output.mp4 --provider kokoro --voice af_heart
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from pptx import Presentation


# =============================================================================
# CONFIGURATION - Edit these settings as needed
# =============================================================================

# --- Timing Settings ---
PADDING_SECONDS = 0.5          # Silence before/after each audio clip
NOTELESS_SLIDE_DURATION = 3.0  # Duration for slides without speaker notes (reduced)
MINIMUM_SLIDE_DURATION = 2.0   # Minimum duration even if audio is shorter

# --- Output Settings ---
OUTPUT_WIDTH = 1920            # Video width in pixels
OUTPUT_HEIGHT = 1080           # Video height in pixels
OUTPUT_FPS = 30                # Frames per second
OUTPUT_AUDIO_RATE = 44100      # Audio sample rate

# --- TTS Settings ---
TTS_PROVIDER = "elevenlabs"    # Options: "elevenlabs", "openai", "azure", "kokoro", "test"

# ElevenLabs
ELEVENLABS_VOICE_ID = "Rachel"
ELEVENLABS_MODEL = "eleven_multilingual_v2"

# OpenAI
OPENAI_VOICE = "alloy"         # alloy, echo, fable, onyx, nova, shimmer
OPENAI_MODEL = "tts-1"         # tts-1 or tts-1-hd

# Azure
AZURE_VOICE = "en-US-JennyNeural"

# Kokoro (local)
KOKORO_VOICE = "af_heart"      # af_bella, af_sarah, am_adam, am_michael, bf_emma, bm_george

# --- Animation Settings ---
CLICK_MARKER = "[CLICK]"       # Marker in speaker notes for animation advance


# =============================================================================
# TTS PROVIDERS
# =============================================================================

def tts_elevenlabs(text: str, output_path: str, voice: str = None) -> bool:
    """Generate TTS using ElevenLabs API."""
    import requests
    
    api_key = os.environ.get("ELEVENLABS_API_KEY")
    if not api_key:
        print("Error: ELEVENLABS_API_KEY environment variable not set")
        return False
    
    voice_id = voice or ELEVENLABS_VOICE_ID
    
    # If voice is a name, we need to look up the ID
    if not voice_id.startswith("21m") and len(voice_id) < 20:
        # Assume it's a name, try to use it directly (ElevenLabs accepts names)
        pass
    
    url = f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}"
    
    headers = {
        "Accept": "audio/mpeg",
        "Content-Type": "application/json",
        "xi-api-key": api_key
    }
    
    data = {
        "text": text,
        "model_id": ELEVENLABS_MODEL,
        "voice_settings": {
            "stability": 0.5,
            "similarity_boost": 0.75
        }
    }
    
    response = requests.post(url, json=data, headers=headers)
    
    if response.status_code == 200:
        with open(output_path, "wb") as f:
            f.write(response.content)
        return True
    else:
        print(f"ElevenLabs API error: {response.status_code} - {response.text}")
        return False


def tts_openai(text: str, output_path: str, voice: str = None) -> bool:
    """Generate TTS using OpenAI API."""
    import requests
    
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("Error: OPENAI_API_KEY environment variable not set")
        return False
    
    url = "https://api.openai.com/v1/audio/speech"
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    data = {
        "model": OPENAI_MODEL,
        "input": text,
        "voice": voice or OPENAI_VOICE,
        "response_format": "mp3"
    }
    
    response = requests.post(url, json=data, headers=headers)
    
    if response.status_code == 200:
        with open(output_path, "wb") as f:
            f.write(response.content)
        return True
    else:
        print(f"OpenAI API error: {response.status_code} - {response.text}")
        return False


def tts_azure(text: str, output_path: str, voice: str = None) -> bool:
    """Generate TTS using Azure Cognitive Services."""
    import requests
    
    api_key = os.environ.get("AZURE_SPEECH_KEY")
    region = os.environ.get("AZURE_SPEECH_REGION", "eastus")
    
    if not api_key:
        print("Error: AZURE_SPEECH_KEY environment variable not set")
        return False
    
    voice_name = voice or AZURE_VOICE
    
    url = f"https://{region}.tts.speech.microsoft.com/cognitiveservices/v1"
    
    headers = {
        "Ocp-Apim-Subscription-Key": api_key,
        "Content-Type": "application/ssml+xml",
        "X-Microsoft-OutputFormat": "audio-16khz-128kbitrate-mono-mp3"
    }
    
    ssml = f"""
    <speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='en-US'>
        <voice name='{voice_name}'>{text}</voice>
    </speak>
    """
    
    response = requests.post(url, headers=headers, data=ssml.encode('utf-8'))
    
    if response.status_code == 200:
        with open(output_path, "wb") as f:
            f.write(response.content)
        return True
    else:
        print(f"Azure API error: {response.status_code} - {response.text}")
        return False


def tts_kokoro(text: str, output_path: str, voice: str = None) -> bool:
    """Generate TTS using Kokoro (local).
    
    First time setup:
        pip install kokoro-onnx soundfile
        
    Then download model files to your working directory:
        - kokoro-v1.0.onnx from https://github.com/thewh1teagle/kokoro-onnx/releases
        - voices-v1.0.bin from https://github.com/thewh1teagle/kokoro-onnx/releases
    """
    try:
        from kokoro_onnx import Kokoro
        import soundfile as sf
    except ImportError:
        print("Error: Kokoro not installed. Run: pip install kokoro-onnx soundfile")
        return False
    
    voice_name = voice or KOKORO_VOICE
    
    # Look for model files in current directory or common locations
    model_paths = [
        "kokoro-v1.0.onnx",
        "models/kokoro-v1.0.onnx",
        os.path.join(os.path.dirname(__file__), "kokoro-v1.0.onnx"),
    ]
    voices_paths = [
        "voices-v1.0.bin",
        "models/voices-v1.0.bin", 
        os.path.join(os.path.dirname(__file__), "voices-v1.0.bin"),
    ]
    
    model_path = None
    voices_path = None
    
    for p in model_paths:
        if os.path.exists(p):
            model_path = p
            break
    
    for p in voices_paths:
        if os.path.exists(p):
            voices_path = p
            break
    
    if not model_path or not voices_path:
        print("Error: Kokoro model files not found.")
        print("Download from: https://github.com/thewh1teagle/kokoro-onnx/releases")
        print("  - kokoro-v1.0.onnx")
        print("  - voices-v1.0.bin")
        print("Place them in the current directory or a 'models' subdirectory.")
        return False
    
    try:
        kokoro = Kokoro(model_path, voices_path)
        samples, sample_rate = kokoro.create(text, voice=voice_name, speed=1.0)
        
        # Save as WAV first, then convert to MP3
        wav_path = output_path.rsplit('.', 1)[0] + '.wav'
        sf.write(wav_path, samples, sample_rate)
        
        # Check WAV was created
        if os.path.exists(wav_path):
            wav_size = os.path.getsize(wav_path)
            print(f"      [kokoro] Generated WAV: {wav_size:,} bytes, {len(samples)/sample_rate:.1f}s")
        else:
            print(f"      [kokoro] ERROR: WAV not created!")
            return False
        
        # Convert to MP3
        result = subprocess.run([
            'ffmpeg', '-y', '-i', wav_path,
            '-acodec', 'libmp3lame', '-ab', '192k',
            output_path
        ], capture_output=True, text=True)
        
        if result.returncode != 0:
            print(f"      [kokoro] FFmpeg error: {result.stderr}")
            return False
        
        # Check MP3 was created
        if os.path.exists(output_path):
            mp3_size = os.path.getsize(output_path)
            print(f"      [kokoro] Converted MP3: {mp3_size:,} bytes")
        else:
            print(f"      [kokoro] ERROR: MP3 not created!")
            return False
        
        os.remove(wav_path)
        return True
        
    except Exception as e:
        print(f"Kokoro error: {e}")
        return False


def tts_test_mode(text: str, output_path: str, voice: str = None) -> bool:
    """Generate silent audio for testing (no API needed).
    
    NOTE: This generates SILENT audio of appropriate duration.
    Use a real TTS provider for actual voiceovers.
    """
    # Create 1 second of silence per 10 words
    words = len(text.split())
    duration = max(1, words / 2.5)  # ~150 words per minute
    
    try:
        subprocess.run([
            'ffmpeg', '-y',
            '-f', 'lavfi', '-i', f'anullsrc=r={OUTPUT_AUDIO_RATE}:cl=mono',
            '-t', str(duration),
            '-acodec', 'libmp3lame',
            output_path
        ], check=True, capture_output=True, text=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f"FFmpeg error generating test audio: {e.stderr}")
        return False


def generate_tts(text: str, output_path: str, voice: str = None) -> bool:
    """Generate TTS audio using the configured provider."""
    providers = {
        "elevenlabs": tts_elevenlabs,
        "openai": tts_openai,
        "azure": tts_azure,
        "kokoro": tts_kokoro,
        "test": tts_test_mode,
    }
    
    provider_func = providers.get(TTS_PROVIDER)
    if not provider_func:
        raise ValueError(f"Unknown TTS provider: {TTS_PROVIDER}")
    
    return provider_func(text, output_path, voice)


# =============================================================================
# SPEAKER NOTES PARSING
# =============================================================================

def has_click_markers(text: str) -> bool:
    """Check if text contains [CLICK] markers."""
    if not text:
        return False
    return CLICK_MARKER.lower() in text.lower()


def parse_notes_with_clicks(notes_text: str) -> list[str]:
    """Split speaker notes by [CLICK] markers."""
    if not notes_text:
        return [""]
    segments = re.split(r'\s*\[CLICK\]\s*', notes_text, flags=re.IGNORECASE)
    return [s.strip() for s in segments]


def extract_speaker_notes(pptx_path: str) -> list[str]:
    """Extract speaker notes from each slide."""
    prs = Presentation(pptx_path)
    notes = []
    
    for slide in prs.slides:
        note_text = ""
        if slide.has_notes_slide:
            notes_slide = slide.notes_slide
            if notes_slide.notes_text_frame:
                note_text = notes_slide.notes_text_frame.text.strip()
        notes.append(note_text)
    
    return notes


# =============================================================================
# SLIDE EXPORT (PowerPoint COM)
# =============================================================================

def export_slides_as_images_windows(pptx_path: str, output_dir: str) -> list[str]:
    """Export all slides as images using PowerPoint COM."""
    from powerpoint_capture import export_all_slides_as_images, WINDOWS_AVAILABLE
    
    if not WINDOWS_AVAILABLE:
        raise RuntimeError("PowerPoint COM not available")
    
    return export_all_slides_as_images(pptx_path, output_dir, OUTPUT_WIDTH, OUTPUT_HEIGHT)


def export_slides_as_images_fallback(pptx_path: str, output_dir: str) -> list[str]:
    """Fallback: Export slides using LibreOffice (cross-platform)."""
    os.makedirs(output_dir, exist_ok=True)
    
    # Convert to PDF first
    subprocess.run([
        'soffice', '--headless', '--convert-to', 'pdf',
        '--outdir', output_dir, pptx_path
    ], check=True, capture_output=True)
    
    pdf_name = os.path.splitext(os.path.basename(pptx_path))[0] + '.pdf'
    pdf_path = os.path.join(output_dir, pdf_name)
    
    # Convert PDF pages to images
    subprocess.run([
        'pdftoppm', '-jpeg', '-r', '150', pdf_path,
        os.path.join(output_dir, 'slide')
    ], check=True, capture_output=True)
    
    # Collect generated images
    images = sorted(Path(output_dir).glob('slide-*.jpg'))
    
    # Rename to consistent format
    result = []
    for i, img in enumerate(images):
        new_name = os.path.join(output_dir, f'slide_{i:04d}.jpg')
        shutil.move(str(img), new_name)
        result.append(new_name)
    
    # Cleanup
    os.remove(pdf_path)
    
    return result


def export_slides_as_images(pptx_path: str, output_dir: str) -> list[str]:
    """Export slides as images, using best available method."""
    try:
        from powerpoint_capture import WINDOWS_AVAILABLE
        if WINDOWS_AVAILABLE:
            return export_slides_as_images_windows(pptx_path, output_dir)
    except ImportError:
        pass
    
    # Fallback to LibreOffice
    return export_slides_as_images_fallback(pptx_path, output_dir)


# =============================================================================
# ANIMATION CAPTURE
# =============================================================================

def capture_animation_states(pptx_path: str, slide_index: int, num_clicks: int,
                              output_dir: str) -> list[str]:
    """Capture animation states for a slide."""
    try:
        from powerpoint_capture import capture_animation_states_windows, WINDOWS_AVAILABLE
        
        if WINDOWS_AVAILABLE:
            return capture_animation_states_windows(
                pptx_path, slide_index, num_clicks, output_dir,
                OUTPUT_WIDTH, OUTPUT_HEIGHT
            )
    except ImportError:
        pass
    
    # No Windows - return empty to trigger fallback
    print("    Warning: PowerPoint COM not available, animations won't be captured")
    return []


# =============================================================================
# AUDIO PROCESSING
# =============================================================================

def get_audio_duration(audio_path: str) -> float:
    """Get duration of an audio file in seconds."""
    if not os.path.exists(audio_path):
        print(f"      [duration] WARNING: Audio file not found: {audio_path}")
        return 0.0
        
    result = subprocess.run([
        'ffprobe', '-v', 'error', '-show_entries', 'format=duration',
        '-of', 'default=noprint_wrappers=1:nokey=1', audio_path
    ], capture_output=True, text=True)
    
    if result.returncode != 0 or not result.stdout.strip():
        print(f"      [duration] WARNING: Could not get duration: {result.stderr}")
        return 0.0
    
    try:
        duration = float(result.stdout.strip())
        print(f"      [duration] Audio duration: {duration:.1f}s")
        return duration
    except ValueError:
        print(f"      [duration] WARNING: Invalid duration: {result.stdout}")
        return 0.0


def add_padding_to_audio(input_path: str, output_path: str, padding: float):
    """Add silence padding before and after audio."""
    # Use adelay for start padding and apad for end padding
    # Convert padding to milliseconds for adelay
    delay_ms = int(padding * 1000)
    
    result = subprocess.run([
        'ffmpeg', '-y',
        '-i', input_path,
        '-af', f'adelay={delay_ms}|{delay_ms},apad=pad_dur={padding}',
        '-acodec', 'aac', '-b:a', '192k',
        output_path
    ], capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"      [padding] FFmpeg error: {result.stderr}")
        raise RuntimeError(f"Audio padding failed: {result.stderr}")
    
    if os.path.exists(output_path):
        size = os.path.getsize(output_path)
        print(f"      [padding] Padded audio: {size:,} bytes")


# =============================================================================
# VIDEO CREATION
# =============================================================================

def create_slide_video(image_path: str, audio_path: str, duration: float,
                       output_path: str):
    """Create a video clip from a single image and audio."""
    # Verify inputs exist
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"Image not found: {image_path}")
    if not os.path.exists(audio_path):
        raise FileNotFoundError(f"Audio not found: {audio_path}")
    
    audio_size = os.path.getsize(audio_path)
    print(f"      [video] Creating from image + audio ({audio_size:,} bytes), duration={duration:.1f}s")
    
    try:
        result = subprocess.run([
            'ffmpeg', '-y',
            '-loop', '1',
            '-i', image_path,
            '-i', audio_path,
            '-c:v', 'libx264',
            '-tune', 'stillimage',
            '-c:a', 'aac',
            '-b:a', '192k',
            '-pix_fmt', 'yuv420p',
            '-shortest',
            '-t', str(duration),
            '-r', str(OUTPUT_FPS),
            output_path
        ], capture_output=True, text=True)
        
        if result.returncode != 0:
            print(f"      [video] FFmpeg error: {result.stderr}")
            raise RuntimeError(f"Video creation failed")
            
        if os.path.exists(output_path):
            size = os.path.getsize(output_path)
            print(f"      [video] Created: {size:,} bytes")
            
    except subprocess.CalledProcessError as e:
        print(f"FFmpeg error creating slide video: {e.stderr}")
        raise


def create_silent_slide_video(image_path: str, duration: float, output_path: str):
    """Create a video clip with no audio."""
    try:
        subprocess.run([
            'ffmpeg', '-y',
            '-loop', '1',
            '-i', image_path,
            '-f', 'lavfi', '-i', f'anullsrc=r={OUTPUT_AUDIO_RATE}:cl=stereo',
            '-c:v', 'libx264',
            '-tune', 'stillimage',
            '-c:a', 'aac',
            '-pix_fmt', 'yuv420p',
            '-t', str(duration),
            '-r', str(OUTPUT_FPS),
            output_path
        ], check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as e:
        print(f"FFmpeg error creating silent video: {e.stderr}")
        raise


def concatenate_videos(video_paths: list[str], output_path: str):
    """Concatenate multiple video clips into one."""
    if len(video_paths) == 1:
        shutil.copy(video_paths[0], output_path)
        return
    
    # Debug: Check each video for audio
    print("  Checking video clips...")
    for vp in video_paths:
        # Check if video has audio stream
        result = subprocess.run([
            'ffprobe', '-v', 'error', '-select_streams', 'a',
            '-show_entries', 'stream=codec_name',
            '-of', 'default=noprint_wrappers=1:nokey=1', vp
        ], capture_output=True, text=True)
        audio_codec = result.stdout.strip() if result.stdout else "NO AUDIO"
        size = os.path.getsize(vp) if os.path.exists(vp) else 0
        print(f"    {os.path.basename(vp)}: {size:,} bytes, audio: {audio_codec}")
    
    # Use file-based concatenation (more reliable on Windows)
    # Create a temp file listing all videos
    list_file = os.path.join(os.path.dirname(video_paths[0]), "concat_list.txt")
    
    with open(list_file, "w", encoding="utf-8") as f:
        for vp in video_paths:
            # FFmpeg concat demuxer needs forward slashes and escaped quotes
            safe_path = vp.replace("\\", "/")
            f.write(f"file '{safe_path}'\n")
    
    cmd = [
        "ffmpeg", "-y",
        "-f", "concat",
        "-safe", "0",
        "-i", list_file,
        "-c:v", "libx264",
        "-c:a", "aac",
        "-movflags", "+faststart",
        output_path
    ]
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as e:
        print(f"FFmpeg concatenation failed!")
        print(f"stderr: {e.stderr}")
        raise
    finally:
        # Cleanup list file
        if os.path.exists(list_file):
            os.remove(list_file)


# =============================================================================
# ANIMATED SLIDE PROCESSING
# =============================================================================

def process_animated_slide(pptx_path: str, slide_index: int, note_text: str,
                           work_dir: str, voice: str = None,
                           static_image: str = None) -> list[str]:
    """
    Process a slide with [CLICK] animation markers.
    
    Returns list of video clip paths, one per animation state.
    """
    # Parse notes into segments
    text_segments = parse_notes_with_clicks(note_text)
    num_clicks = len(text_segments) - 1
    
    # Try to capture animation states
    anim_dir = os.path.join(work_dir, f"anim_slide_{slide_index}")
    os.makedirs(anim_dir, exist_ok=True)
    
    image_paths = capture_animation_states(pptx_path, slide_index, num_clicks, anim_dir)
    
    # Fallback to static image if capture failed
    if not image_paths:
        if static_image:
            print(f"    Using static image for all {len(text_segments)} segments")
            image_paths = [static_image] * len(text_segments)
        else:
            return None
    
    # Ensure we have the right number of images
    while len(image_paths) < len(text_segments):
        image_paths.append(image_paths[-1])
    image_paths = image_paths[:len(text_segments)]
    
    # Generate video clip for each animation state
    video_clips = []
    
    for state_idx, (image_path, text) in enumerate(zip(image_paths, text_segments)):
        video_path = os.path.join(work_dir, f"slide_{slide_index:04d}_state_{state_idx:04d}.mp4")
        
        if text:
            # Generate TTS
            audio_raw = os.path.join(work_dir, f"audio_{slide_index}_{state_idx}_raw.mp3")
            audio_padded = os.path.join(work_dir, f"audio_{slide_index}_{state_idx}.m4a")
            
            if generate_tts(text, audio_raw, voice):
                add_padding_to_audio(audio_raw, audio_padded, PADDING_SECONDS)
                duration = max(get_audio_duration(audio_padded), MINIMUM_SLIDE_DURATION)
                create_slide_video(image_path, audio_padded, duration, video_path)
            else:
                create_silent_slide_video(image_path, NOTELESS_SLIDE_DURATION, video_path)
        else:
            # Empty segment (consecutive clicks) - brief pause
            create_silent_slide_video(image_path, 1.0, video_path)
        
        video_clips.append(video_path)
    
    return video_clips


# =============================================================================
# MAIN PIPELINE
# =============================================================================

def convert_pptx_to_video(pptx_path: str, output_path: str, voice: str = None,
                          verbose: bool = True):
    """Main conversion pipeline."""
    if verbose:
        print(f"Converting: {pptx_path}")
        print(f"TTS Provider: {TTS_PROVIDER}")
        if TTS_PROVIDER == "test":
            print(f"  NOTE: Test mode produces SILENT audio (for pipeline testing)")
        print(f"Output: {output_path}")
        print()
    
    # Create temp directory
    work_dir = tempfile.mkdtemp(prefix="pptx2video_")
    
    try:
        # Step 1: Extract speaker notes
        if verbose:
            print("Extracting speaker notes...")
        notes = extract_speaker_notes(pptx_path)
        if verbose:
            print(f"  Found {len(notes)} slides, {sum(1 for n in notes if n)} have notes")
        
        # Step 2: Export slides as images
        if verbose:
            print("Exporting slides as images...")
        images = export_slides_as_images(pptx_path, work_dir)
        if verbose:
            print(f"  Generated {len(images)} images")
        
        if len(images) != len(notes):
            print(f"Warning: Image count ({len(images)}) != slide count ({len(notes)})")
        
        # Check for animated slides
        animated_slides = [i for i, n in enumerate(notes) if has_click_markers(n)]
        if animated_slides and verbose:
            print(f"  Found {len(animated_slides)} slides with [CLICK] markers")
        
        # Step 3: Generate video for each slide
        slide_videos = []
        
        for i, (image_path, note_text) in enumerate(zip(images, notes)):
            slide_num = i + 1
            if verbose:
                print(f"Processing slide {slide_num}/{len(images)}...", end=" ")
            
            # Check if this slide has animation markers
            if has_click_markers(note_text):
                if verbose:
                    print("(animated)", end=" ")
                
                anim_clips = process_animated_slide(
                    pptx_path, i, note_text, work_dir, voice,
                    static_image=image_path
                )
                
                if anim_clips:
                    slide_videos.extend(anim_clips)
                    if verbose:
                        print(f"({len(anim_clips)} states)")
                    continue
                else:
                    if verbose:
                        print("(fallback)", end=" ")
                    note_text = ' '.join(parse_notes_with_clicks(note_text))
            
            # Regular slide processing
            video_path = os.path.join(work_dir, f"slide_{i:04d}.mp4")
            
            if note_text:
                audio_raw = os.path.join(work_dir, f"audio_{i:04d}_raw.mp3")
                audio_padded = os.path.join(work_dir, f"audio_{i:04d}.m4a")
                
                if not generate_tts(note_text, audio_raw, voice):
                    print(f"TTS failed for slide {slide_num}")
                    create_silent_slide_video(image_path, NOTELESS_SLIDE_DURATION, video_path)
                else:
                    add_padding_to_audio(audio_raw, audio_padded, PADDING_SECONDS)
                    duration = max(get_audio_duration(audio_padded), MINIMUM_SLIDE_DURATION)
                    create_slide_video(image_path, audio_padded, duration, video_path)
                    
                    if verbose:
                        print(f"({duration:.1f}s)")
            else:
                create_silent_slide_video(image_path, NOTELESS_SLIDE_DURATION, video_path)
                if verbose:
                    print(f"(no notes, {NOTELESS_SLIDE_DURATION}s)")
            
            slide_videos.append(video_path)
        
        # Step 4: Concatenate
        if verbose:
            print("Concatenating final video...")
        concatenate_videos(slide_videos, output_path)
        
        if verbose:
            result = subprocess.run([
                "ffprobe", "-v", "error", "-show_entries", "format=duration",
                "-of", "default=noprint_wrappers=1:nokey=1", output_path
            ], capture_output=True, text=True)
            total_duration = float(result.stdout.strip())
            print(f"\nDone! Total duration: {total_duration:.1f}s")
            print(f"Output saved to: {output_path}")
    
    finally:
        shutil.rmtree(work_dir)


def main():
    parser = argparse.ArgumentParser(
        description="Convert PowerPoint presentations to videos with voiceovers"
    )
    parser.add_argument("input", help="Input PPTX file")
    parser.add_argument("output", help="Output MP4 file")
    parser.add_argument("--provider", choices=["elevenlabs", "openai", "azure", "kokoro", "test"],
                        help="TTS provider to use")
    parser.add_argument("--voice", help="Voice name/ID for TTS")
    parser.add_argument("--quiet", action="store_true", help="Suppress progress output")
    
    args = parser.parse_args()
    
    if args.provider:
        global TTS_PROVIDER
        TTS_PROVIDER = args.provider
    
    if not os.path.exists(args.input):
        print(f"Error: Input file not found: {args.input}")
        sys.exit(1)
    
    convert_pptx_to_video(args.input, args.output, args.voice, not args.quiet)


if __name__ == "__main__":
    main()
