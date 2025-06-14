import os
import sys
from datetime import datetime
import exifread
import pywintypes
import win32file
import win32con

def get_exif_datetime_original(file_path):
    """Extract DateTimeOriginal from EXIF data of an image file."""
    with open(file_path, 'rb') as f:
        tags = exifread.process_file(f, details=False)
        if 'EXIF DateTimeOriginal' in tags:
            return str(tags['EXIF DateTimeOriginal'])
    return None

def parse_exif_datetime(dt_str):
    """Convert EXIF datetime string to datetime object."""
    try:
        return datetime.strptime(dt_str, '%Y:%m:%d %H:%M:%S')
    except (ValueError, TypeError):
        return None

def set_file_dates(file_path, dt):
    """Set both creation and modification dates of a file on Windows."""
    if not isinstance(dt, datetime):
        raise ValueError("dt must be a datetime object")
    
    # Convert datetime to Windows time format
    wintime = pywintypes.Time(dt)
    
    # Get file handle
    handle = win32file.CreateFile(
        file_path,
        win32con.GENERIC_WRITE,
        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
        None,
        win32con.OPEN_EXISTING,
        0,
        None
    )
    
    # Set file times (creation, last access, last write)
    win32file.SetFileTime(handle, wintime, wintime, wintime)
    handle.close()

def process_image_file(file_path):
    """Process a single image file to update its dates from EXIF."""
    print(f"Processing: {file_path}")
    
    # Get EXIF DateTimeOriginal
    dt_str = get_exif_datetime_original(file_path)
    if not dt_str:
        print(f"  Warning: No DateTimeOriginal EXIF data found in {file_path}")
        return False
    
    # Parse datetime
    dt = parse_exif_datetime(dt_str)
    if not dt:
        print(f"  Warning: Could not parse DateTimeOriginal: {dt_str}")
        return False
    
    # Set file dates
    try:
        set_file_dates(file_path, dt)
        print(f"  Success: Set file dates to {dt}")
        return True
    except Exception as e:
        print(f"  Error setting file dates: {str(e)}")
        return False

def main():
    if len(sys.argv) < 2:
        print("Usage: python set_photo_dates.py <file_or_directory>")
        print("  Processes all JPEG files in the specified directory or file")
        return
    
    target = sys.argv[1]
    
    if os.path.isfile(target):
        # Process single file
        if target.lower().endswith(('.jpg', '.jpeg', '.tiff', '.tif', '.nef', '.cr2', '.arw')):
            process_image_file(target)
        else:
            print("Error: Unsupported file format")
    elif os.path.isdir(target):
        # Process all image files in directory
        for root, _, files in os.walk(target):
            for file in files:
                if file.lower().endswith(('.jpg', '.jpeg', '.tiff', '.tif', '.nef', '.cr2', '.arw')):
                    process_image_file(os.path.join(root, file))
    else:
        print("Error: Target path does not exist")

if __name__ == "__main__":
    main()
