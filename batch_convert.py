#!/usr/bin/env python3
"""
Batch PPTX to Video Converter

Converts multiple PowerPoint presentations to videos with parallel processing.

Usage:
    python batch_convert.py ./presentations ./videos
    python batch_convert.py ./presentations ./videos --workers 4 --provider kokoro
"""

import argparse
import os
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path


def convert_single(args):
    """Convert a single PPTX file. Used for parallel processing."""
    input_path, output_path, provider, voice = args
    
    # Import here to avoid issues with multiprocessing
    from convert import convert_pptx_to_video, TTS_PROVIDER
    
    # Set provider if specified
    if provider:
        import convert
        convert.TTS_PROVIDER = provider
    
    try:
        convert_pptx_to_video(input_path, output_path, voice, verbose=False)
        return input_path, True, None
    except Exception as e:
        return input_path, False, str(e)


def batch_convert(input_dir: str, output_dir: str, workers: int = 2,
                  provider: str = None, voice: str = None, recursive: bool = True):
    """
    Convert all PPTX files in a directory.
    
    Args:
        input_dir: Directory containing PPTX files
        output_dir: Directory for output videos
        workers: Number of parallel workers
        provider: TTS provider to use
        voice: Voice name/ID
        recursive: Search subdirectories
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    
    # Find all PPTX files
    if recursive:
        pptx_files = list(input_path.rglob("*.pptx"))
    else:
        pptx_files = list(input_path.glob("*.pptx"))
    
    # Filter out temp files
    pptx_files = [f for f in pptx_files if not f.name.startswith("~$")]
    
    if not pptx_files:
        print(f"No PPTX files found in {input_dir}")
        return
    
    print(f"Found {len(pptx_files)} PPTX files")
    print(f"Using {workers} parallel workers")
    print(f"TTS Provider: {provider or 'default'}")
    print()
    
    # Prepare conversion tasks
    tasks = []
    for pptx_file in pptx_files:
        # Maintain directory structure
        relative = pptx_file.relative_to(input_path)
        output_file = output_path / relative.with_suffix(".mp4")
        
        # Create output directory
        output_file.parent.mkdir(parents=True, exist_ok=True)
        
        tasks.append((str(pptx_file), str(output_file), provider, voice))
    
    # Process in parallel
    completed = 0
    failed = 0
    start_time = time.time()
    
    with ProcessPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(convert_single, task): task for task in tasks}
        
        for future in as_completed(futures):
            input_file, success, error = future.result()
            completed += 1
            
            filename = os.path.basename(input_file)
            
            if success:
                print(f"[{completed}/{len(tasks)}] ✓ {filename}")
            else:
                failed += 1
                print(f"[{completed}/{len(tasks)}] ✗ {filename}: {error}")
    
    # Summary
    elapsed = time.time() - start_time
    print()
    print(f"Completed: {completed - failed}/{len(tasks)} successful")
    print(f"Time: {elapsed:.1f}s ({elapsed/len(tasks):.1f}s per file)")
    
    if failed > 0:
        print(f"Failed: {failed} files")


def main():
    parser = argparse.ArgumentParser(
        description="Batch convert PowerPoint presentations to videos"
    )
    parser.add_argument("input_dir", help="Input directory with PPTX files")
    parser.add_argument("output_dir", help="Output directory for videos")
    parser.add_argument("--workers", type=int, default=2,
                        help="Number of parallel workers (default: 2)")
    parser.add_argument("--provider", choices=["elevenlabs", "openai", "azure", "kokoro", "test"],
                        help="TTS provider")
    parser.add_argument("--voice", help="Voice name/ID")
    parser.add_argument("--no-recursive", action="store_true",
                        help="Don't search subdirectories")
    
    args = parser.parse_args()
    
    if not os.path.isdir(args.input_dir):
        print(f"Error: Input directory not found: {args.input_dir}")
        sys.exit(1)
    
    batch_convert(
        args.input_dir,
        args.output_dir,
        args.workers,
        args.provider,
        args.voice,
        not args.no_recursive
    )


if __name__ == "__main__":
    main()
