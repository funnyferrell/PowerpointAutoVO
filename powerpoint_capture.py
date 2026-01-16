"""
PowerPoint Animation Capture using Windows COM Automation.

This module uses pywin32 to control Microsoft PowerPoint and capture
each animation state by manipulating shape visibility and using native export.

Requirements:
    - Windows OS
    - Microsoft PowerPoint installed
    - pip install pywin32
"""

import os
import time
import tempfile
from pathlib import Path

try:
    import win32com.client
    import pythoncom
    WINDOWS_AVAILABLE = True
except ImportError:
    WINDOWS_AVAILABLE = False


# PowerPoint constants
msoTrue = -1
msoFalse = 0


def check_windows_dependencies():
    """Check if Windows dependencies are available."""
    if not WINDOWS_AVAILABLE:
        return False, "pywin32 required. Install with: pip install pywin32"
    return True, "Windows dependencies available"


def get_animated_shape_ids(slide) -> list[list[int]]:
    """
    Get shape IDs that have click-triggered animations.
    
    Returns list of shape ID groups, one per click trigger.
    """
    try:
        timeline = slide.TimeLine
        main_sequence = timeline.MainSequence
        
        if main_sequence.Count == 0:
            return []
        
        click_groups = []
        current_group = []
        
        for i in range(1, main_sequence.Count + 1):
            effect = main_sequence.Item(i)
            shape_id = effect.Shape.Id
            trigger_type = effect.Timing.TriggerType
            
            # TriggerType: 1 = OnClick, 2 = WithPrevious, 3 = AfterPrevious
            if trigger_type == 1:  # OnClick - new click group
                if current_group:
                    click_groups.append(current_group)
                current_group = [shape_id]
            else:  # WithPrevious or AfterPrevious - same click group
                current_group.append(shape_id)
        
        if current_group:
            click_groups.append(current_group)
        
        return click_groups
        
    except Exception as e:
        print(f"    Warning: Could not parse animations: {e}")
        return []


def capture_animation_states_windows(pptx_path: str, slide_index: int,
                                      num_clicks: int, output_dir: str,
                                      width: int = 1920, height: int = 1080) -> list[str]:
    """
    Capture animation states by manipulating shape visibility and exporting.
    
    This approach:
    1. Identifies which shapes are animated
    2. Hides all animated shapes initially
    3. For each click state, reveals the appropriate shapes
    4. Exports the slide as an image
    5. Restores all shapes at the end
    """
    if not WINDOWS_AVAILABLE:
        raise RuntimeError("Windows COM automation not available")
    
    pptx_path = os.path.abspath(pptx_path)
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"    [powerpoint] Capturing slide {slide_index + 1} with {num_clicks} clicks...")
    
    # Initialize COM
    pythoncom.CoInitialize()
    
    ppt = None
    presentation = None
    captured_images = []
    original_visibility = {}  # Store original visibility to restore later
    
    try:
        # Start PowerPoint (hidden)
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        # Don't set Visible - keep it hidden for cleaner operation
        
        # Open presentation
        presentation = ppt.Presentations.Open(pptx_path, WithWindow=msoFalse)
        
        # Get the slide (1-indexed in COM)
        slide = presentation.Slides(slide_index + 1)
        
        # Get animation sequence
        click_groups = get_animated_shape_ids(slide)
        
        if not click_groups:
            print(f"    [powerpoint] No click animations found, using static export")
            # Just export the slide as-is
            img_path = os.path.join(output_dir, f"slide_{slide_index:04d}_state_0000.png")
            slide.Export(img_path, "PNG", width, height)
            return [img_path] if os.path.exists(img_path) else []
        
        print(f"    [powerpoint] Found {len(click_groups)} animation groups")
        
        # Get all animated shape IDs
        all_animated_ids = set()
        for group in click_groups:
            all_animated_ids.update(group)
        
        # Build a map of shape ID -> shape object
        shape_map = {}
        for shape in slide.Shapes:
            if shape.Id in all_animated_ids:
                shape_map[shape.Id] = shape
                # Store original visibility
                original_visibility[shape.Id] = shape.Visible
        
        # Hide all animated shapes initially
        for shape_id, shape in shape_map.items():
            shape.Visible = msoFalse
        
        # Capture initial state (all animated shapes hidden)
        img_path = os.path.join(output_dir, f"slide_{slide_index:04d}_state_0000.png")
        slide.Export(img_path, "PNG", width, height)
        captured_images.append(img_path)
        print(f"    [powerpoint] State 0: exported")
        
        # Reveal shapes group by group and capture
        states_to_capture = min(num_clicks, len(click_groups))
        
        for state_idx in range(states_to_capture):
            # Reveal this click group's shapes
            for shape_id in click_groups[state_idx]:
                if shape_id in shape_map:
                    shape_map[shape_id].Visible = msoTrue
            
            # Export
            img_path = os.path.join(output_dir, f"slide_{slide_index:04d}_state_{state_idx + 1:04d}.png")
            slide.Export(img_path, "PNG", width, height)
            captured_images.append(img_path)
            print(f"    [powerpoint] State {state_idx + 1}: exported")
        
        # Verify images
        for i, img_path in enumerate(captured_images):
            if os.path.exists(img_path):
                size = os.path.getsize(img_path)
                print(f"    [powerpoint] State {i}: {size:,} bytes")
        
        return captured_images
        
    except Exception as e:
        print(f"    [powerpoint] Error: {e}")
        import traceback
        traceback.print_exc()
        raise
        
    finally:
        # Restore original visibility
        try:
            if presentation and original_visibility:
                slide = presentation.Slides(slide_index + 1)
                for shape in slide.Shapes:
                    if shape.Id in original_visibility:
                        shape.Visible = original_visibility[shape.Id]
        except:
            pass
        
        # Close without saving (we modified visibility)
        try:
            if presentation:
                presentation.Close()
            if ppt:
                ppt.Quit()
        except:
            pass
        
        pythoncom.CoUninitialize()


def export_slide_as_image(pptx_path: str, slide_index: int, output_path: str,
                          width: int = 1920, height: int = 1080) -> bool:
    """
    Export a single slide as an image using PowerPoint.
    """
    if not WINDOWS_AVAILABLE:
        raise RuntimeError("Windows COM automation not available")
    
    pptx_path = os.path.abspath(pptx_path)
    output_path = os.path.abspath(output_path)
    
    pythoncom.CoInitialize()
    
    ppt = None
    presentation = None
    
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        presentation = ppt.Presentations.Open(pptx_path, WithWindow=msoFalse)
        
        slide = presentation.Slides(slide_index + 1)
        slide.Export(output_path, "PNG", width, height)
        
        return os.path.exists(output_path)
        
    except Exception as e:
        print(f"Export error: {e}")
        return False
        
    finally:
        try:
            if presentation:
                presentation.Close()
            if ppt:
                ppt.Quit()
        except:
            pass
        
        pythoncom.CoUninitialize()


def export_all_slides_as_images(pptx_path: str, output_dir: str,
                                 width: int = 1920, height: int = 1080) -> list[str]:
    """
    Export all slides as images using PowerPoint native export.
    """
    if not WINDOWS_AVAILABLE:
        raise RuntimeError("Windows COM automation not available")
    
    pptx_path = os.path.abspath(pptx_path)
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    
    pythoncom.CoInitialize()
    
    ppt = None
    presentation = None
    
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        presentation = ppt.Presentations.Open(pptx_path, WithWindow=msoFalse)
        
        image_paths = []
        
        for i in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides(i)
            img_path = os.path.join(output_dir, f"slide_{i-1:04d}.png")
            slide.Export(img_path, "PNG", width, height)
            image_paths.append(img_path)
        
        return image_paths
        
    finally:
        try:
            if presentation:
                presentation.Close()
            if ppt:
                ppt.Quit()
        except:
            pass
        
        pythoncom.CoUninitialize()


# =============================================================================
# CLI for testing
# =============================================================================

if __name__ == "__main__":
    import sys
    
    ok, msg = check_windows_dependencies()
    if not ok:
        print(f"Error: {msg}")
        sys.exit(1)
    
    if len(sys.argv) < 2:
        print("Usage: python powerpoint_capture.py <pptx_file> [slide_index] [num_clicks]")
        print()
        print("Example:")
        print("  python powerpoint_capture.py presentation.pptx 1 3")
        sys.exit(1)
    
    pptx_file = sys.argv[1]
    slide_index = int(sys.argv[2]) if len(sys.argv) > 2 else 0
    num_clicks = int(sys.argv[3]) if len(sys.argv) > 3 else 3
    
    output_dir = os.path.join(tempfile.gettempdir(), "ppt_capture_test")
    
    print(f"Testing PowerPoint animation capture")
    print(f"  PPTX: {pptx_file}")
    print(f"  Slide: {slide_index + 1}")
    print(f"  Clicks: {num_clicks}")
    print(f"  Output: {output_dir}")
    print()
    
    try:
        images = capture_animation_states_windows(
            pptx_file, slide_index, num_clicks, output_dir
        )
        
        print()
        print(f"Success! Captured {len(images)} images:")
        for img in images:
            print(f"  {img}")
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
