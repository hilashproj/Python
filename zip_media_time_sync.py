#!/usr/bin/env python3
"""
zip_media_time_sync.py

Extract a ZIP archive into a new folder, then update media file timestamps from
Google Photos supplemental metadata JSON files found alongside the media.

Notes:
- On Linux/Unix, true file "creation time" (birthtime) is generally not settable.
  This tool updates modification time (mtime) and access time (atime). On
  Windows, creation time could be updated with platform-specific APIs, which is
  not implemented here.

Usage example:
  python zip_media_time_sync.py /path/to/Takeout.zip --output-dir /path/to/output --verbose

"""

from __future__ import annotations

import argparse
import json
import logging
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import os
from datetime import datetime

import piexif
from PIL import Image
from PIL.ExifTags import TAGS

try:
    import win32file
    import win32con
    import pywintypes
    WINDOWS_AVAILABLE = True

except ImportError as e:
    WINDOWS_AVAILABLE = False
    print(f"An ImportError occurred: {e}")
    print(f"Name of the missing module/item: {e.name}")
    # Note: path might be None if the issue isn't file-related.
#   print(f"Path checked: {e.path}")
    print("Warning: pywin32 not installed. Install it with: python -m pip install pywin32")



# ----------------------------- Logging Utilities -----------------------------


def configure_logging(verbosity: int) -> None:
    """
    Handling the logging within the process
    :param verbosity:
    :return:
    """
    level = logging.WARNING
    if verbosity >= 2:
        level = logging.DEBUG
    elif verbosity == 1:
        level = logging.INFO
    logging.basicConfig(
        level=level,
        format="%(levelname)s: %(message)s",
    )


# ------------------------------- Data Classes --------------------------------


@dataclass
class UpdateResult:
    """
    The Json data content that is relevant for
    """
    media_path: Path
    json_path: Path
    timestamp: int
    updated: bool
    reason: str = ""


# --------------------------------- Helpers -----------------------------------


def ensure_new_extraction_dir(zip_path: Path, output_dir: Optional[Path]) -> Path:
    """
    The method creates a new directory at the output directory as it was
    sent as a parameter to the command line and create a new folder:

    {Zip file name (without the extension)_extracted_2}_extracted{_number that doesnt exist}
    Ex: If the output_dir is: F:\\Project_output\\Output
    and the zip file is: GooglePhotosTakeout.zip
    It will create the folder: F:\\Project_output\\Output\\GooglePhotosTakeout_extracted
    and if it exists it will create: F:\\Project_output\\Output\\GooglePhotosTakeout_extracted_1 ...

    :param zip_path:
    :param output_dir:
    :return:
    """
    zip_stem = zip_path.stem
    parent = output_dir if output_dir else zip_path.parent
    # Ensure a deterministic new folder name. Avoid clobbering existing contents.
    candidate = parent / f"{zip_stem}_extracted"
    suffix = 1
    while candidate.exists():
        candidate = parent / f"{zip_stem}_extracted_{suffix}"
        suffix += 1
    candidate.mkdir(parents=True, exist_ok=False)
    return candidate


def _is_within_directory(directory: Path, target: Path) -> bool:
    try:
        directory_abs = directory.resolve(strict=False)
        target_abs = target.resolve(strict=False)
        common = os.path.commonpath([str(directory_abs), str(target_abs)])
        return common == str(directory_abs)
    except Exception:
        return False


def safe_extract_zip(zip_path: Path, destination: Path) -> None:
    """
    This method extracts the zip file and also creates the relevant folders and sub folders
    :param zip_path:
    :param destination:
    :return:
    """
    import zipfile

    logging.info("Extracting ZIP to %s", destination)
    with zipfile.ZipFile(zip_path, 'r') as zf:
        for member in zf.infolist():
            member_name = member.filename
            # Zip entries may contain backslashes on Windows-created zips; normalize.
            member_name = member_name.replace("\\", "/")

            # Skip absolute paths
            if member_name.startswith("/"):
                logging.warning("Skipping absolute path member: %s", member_name)
                continue

            dest_path = destination / member_name
            if not _is_within_directory(destination, dest_path):
                logging.warning("Skipping path traversal member: %s", member_name)
                continue

            if member.is_dir():
                dest_path.mkdir(parents=True, exist_ok=True)
                continue

            dest_path.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(member, 'r') as src, open(dest_path, 'wb') as dst:
                shutil.copyfileobj(src, dst)


def iter_metadata_json_files(root: Path) -> Iterable[Path]:
    # Using rglob('*.json') across the extracted tree
    return root.rglob('*.json')


def candidate_media_name_from_json(json_filename: str) -> Optional[str]:
    """
     Accept patterns like:
    <name>.<ext>.json
    <name>.<ext>.metadata.json
    <name>.<ext>.supplemental-metadata.json
    <name>.<ext>.suppl.json
    <name>.<ext>.s.json
Fix the     But since there could be more, we will extract the name.ext from the json file name
    :param json_filename:
    :return:
    """

    # Splits the string into 3 parts starting from the right, then takes the first part
    split_text = json_filename.rsplit('.', 3)

    if split_text[-1] == 'json' and split_text[1] != 'json':
        base = split_text[0] + '.' + split_text[1]
        return base
    else:
        return None


def find_matching_media(json_path: Path, case_insensitive: bool) -> Optional[Path]:
    """

    :param json_path:
    :param case_insensitive:
    :return:
    """
    candidate_name = candidate_media_name_from_json(json_path.name)
    if not candidate_name:
        return None
    # the with_name() method is used to return a new path object where the final component
    # (the filename or directory name) is replaced by a new name.
    direct_candidate = json_path.with_name(candidate_name)
    if direct_candidate.exists():
        return direct_candidate

    if not case_insensitive:
        return None

    # Case-insensitive fallback within the same directory
    try:
        lower_candidate = candidate_name.lower()
        matches = [p for p in json_path.parent.iterdir() if p.is_file() and p.name.lower() == lower_candidate]
        if len(matches) == 1:
            return matches[0]
    except FileNotFoundError:
        return None
    return None


def parse_timestamp_from_json(json_path: Path) -> tuple[int | None, int | None] | None:
    """
    Parse the timestamp from a JSON file. This method returns the photo_taken_time and creation_time
    :param json_path:
    :return: photo_taken_time, creation_time
    :rtype int, int
    """
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as exc:
        logging.warning("Failed to read JSON %s: %s", json_path, exc)
        return None


    # Google Photos metadata commonly has objects with 'timestamp' fields
    def extract(obj: Dict) -> Optional[int]:
        if not isinstance(obj, dict):
            return None
        ts = obj.get('timestamp')
        if ts is None:
            return None
        try:
            # Timestamps are usually seconds as strings
            return int(float(ts))
        except Exception:
            return None

    val = data.get('photoTakenTime')
    photo_taken_time = extract(val)
    val = data.get('creationTime')
    creation_time = extract(val)
    if photo_taken_time is not None and creation_time is not None:
        return photo_taken_time, creation_time


    # for key in ('photoTakenTime', 'creationTime', 'imageCreationTime', 'modificationTime'):
    #     val = data.get(key)
    #     ts = extract(val)
    #     if ts is not None:
    #         return ts

    logging.debug("No supported timestamp in %s", json_path)
    return None


def apply_timestamp_to_file(media_path: Path, photo_taken_time: int, creation_time: int, dry_run: bool) -> Tuple[bool, str]:
    """
    This method updates the timestamp of the file.
    It set the Access Time (atime) and Modification Time (mtime) of the jpeg files.
    For movies (*.mov) it will update the "Media Created."
    :param media_path:
    :param photo_taken_time:
    :param creation_time:
    :param dry_run:
    :return:
    """

    #get_exif_data_of_file(file_path,timestamp)

    try:
        if dry_run:
            return False, "dry-run"
        #Update access time and modification time


        #Update the date taken time
        update_date_taken(media_path, photo_taken_time)
        # Update the creation Date of the jpg file
        update_creation_date(str(media_path), creation_time)
        return True, ""
    except Exception as exc:
        return False, str(exc)


def scan_and_update(root: Path, *, dry_run: bool, case_insensitive: bool) -> List[UpdateResult]:
    """
    This method finds all the json files, then finds all the relevant file files that relates
    to the json file inorder to change its 'Photo Taken Time' and 'Creation Time'
    :param root:
    :param dry_run:
    :param case_insensitive:
    :return:
    """
    results: List[UpdateResult] = []
    for json_path in iter_metadata_json_files(root):
        media_path = find_matching_media(json_path, case_insensitive)
        if not media_path:
            logging.debug("No matching media for %s", json_path)
            continue
        photo_taken_time, creation_time = parse_timestamp_from_json(json_path)
        if photo_taken_time is None and creation_time is None:
            results.append(UpdateResult(media_path=media_path, json_path=json_path, timestamp=-1, updated=False, reason="no-timestamp"))
            continue
        #Update the dateTaken
        updated, err = apply_timestamp_to_file(media_path, photo_taken_time, creation_time, dry_run)
        if updated:
            logging.info("Updated %s from %s to %s", media_path, json_path, photo_taken_time)
            results.append(UpdateResult(media_path=media_path, json_path=json_path, timestamp=photo_taken_time, updated=True))
            # Once the media was updated you can delete the json file
            delete_json_file(json_path)
        else:
            if err == "dry-run":
                logging.info("Would update %s from %s to %s", media_path, json_path, photo_taken_time)
                results.append(UpdateResult(media_path=media_path, json_path=json_path, timestamp=photo_taken_time, updated=False, reason="dry-run"))
            else:
                logging.warning("Failed to update %s: %s", media_path, err)
                results.append(UpdateResult(media_path=media_path, json_path=json_path, timestamp=photo_taken_time, updated=False, reason=err))

    return results


def delete_json_file(json_path: Path) -> None:
    json_path.unlink(missing_ok=True)
    return None


def update_creation_date(file_path, new_date :int):
    """
    Update the creation date of a file.

    Args:
        file_path (str): Path to the file
        new_date (datetime): New creation date/time
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    if not WINDOWS_AVAILABLE:
        raise ImportError("pywin32 is required for Windows. Install with: pip install pywin32")

    # Convert datetime to Windows FILETIME format
    filetime = pywintypes.Time(new_date)
    try:
        # Open the file handle
        handle = win32file.CreateFile(
            file_path,
            win32con.GENERIC_WRITE,
            win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
            None,
            win32con.OPEN_EXISTING,
            0,
            None
        )

        # Get current file times
        creation_time, access_time, write_time = win32file.GetFileTime(handle)

        # Update only the creation time, keep others unchanged
        win32file.SetFileTime(handle, filetime, access_time, write_time,False)
        print(f"Successfully updated creation date of '{file_path}' to {new_date}")
    finally:
        win32file.CloseHandle(handle)

    return None


# ----------------------------------- CLI -------------------------------------


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    """
    This method takes the parameters that were sent from the command line and add them to
    :param argv:
    :return:
    """
    parser = argparse.ArgumentParser(description="Extract ZIP and sync media timestamps from Google Photos JSON metadata.")
    parser.add_argument("zip_path", type=Path, help="Path to the ZIP archive (e.g., Google Takeout)")
    parser.add_argument("--output-dir", type=Path, default=None, help="Directory where a new extraction folder will be created.")
    parser.add_argument("--dry-run", action="store_true", help="Do not modify files, just report intended changes.")
    parser.add_argument("-v", "--verbose", action="count", default=0, help="Increase verbosity (use -v or -vv).")
    parser.add_argument("--case-insensitive", action="store_true", help="Allow case-insensitive filename matching if exact match is missing.")
    return parser.parse_args(argv)

def get_exif_data_of_file(file_path: Path, timestamp: int):
    """
    This method is for assistance only. Currently, it doesn't being called, but it checks the EXIF
    of the files.
    The taken photo is saved in the Exif metadata of the media files. so in order to update it
    we need to update the Exif data.
    :param file_path:
    :param timestamp:
    :return:
    """
    img = Image.open(file_path)
    exif_data = img.getexif()

    for tag_id, value in exif_data.items():
        tag_name = TAGS.get(tag_id, tag_id)
        print(f"{tag_name} ({tag_id}): {value}")

    return None


def update_date_taken(file_path: Path, timestamp: int) -> None:
    """
    Updates the date taken field of a JPEG or movie file.

    Args:
        file_path: Path to the JPEG or movie file
        timestamp: Unix timestamp (seconds since epoch) to set as the date taken
    """
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    # Convert timestamp to datetime
    dt = datetime.fromtimestamp(timestamp)

    # Get file extension to determine file type
    ext = file_path.suffix.lower()

    if ext in ['.jpg', '.jpeg']:
        _update_jpeg_date(file_path, dt)
    elif ext in ['.mp4', '.mov', '.m4v', '.3gp']:
        _update_movie_date(file_path, dt)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Supported formats: JPEG, MP4, MOV, M4V, 3GP")

    os.utime(file_path, (timestamp, timestamp))


def _update_jpeg_date(file_path: Path, dt: datetime) -> None:
    """
    Updates EXIF date taken field for JPEG files.
    :param file_path:
    :param dt:
    :return:
    """

    # Format as EXIF date string (YYYY:MM:DD HH:MM:SS)
    exif_date = dt.strftime("%Y:%m:%d %H:%M:%S")

    # Open the image
    with (Image.open(file_path) as img):
        # Get existing EXIF data
        #exif_dict = img.getexif()
        # Load existing EXIF data or create new dict
        try:
            exif_dict = piexif.load(img.info.get('exif', b''))
        except:
            exif_dict = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}}

        # Update date/time fields
        # 306 = DateTime
        # 36867 = DateTimeOriginal (date/time when original image was taken)
        # 36868 = DateTimeDigitized

        #exif_dict[306] = exif_date
        #exif_dict[36867] = exif_date
        #exif_dict[36868] = exif_date
        exif_dict["0th"][piexif.ImageIFD.DateTime] = exif_date
        exif_dict["Exif"][piexif.ExifIFD.DateTimeOriginal] = exif_date
        exif_dict["Exif"][piexif.ExifIFD.DateTimeDigitized] = exif_date


        if piexif.ExifIFD.SceneType in exif_dict['Exif'] and isinstance(exif_dict['Exif'][piexif.ExifIFD.SceneType], int):
            exif_dict['Exif'][piexif.ExifIFD.SceneType] = str(exif_dict['Exif'][piexif.ExifIFD.SceneType]).encode('utf-8')

        # Convert to bytes and save
        exif_bytes = piexif.dump(exif_dict)
        # Save the image with updated EXIF data
        img.save(file_path, exif=exif_bytes, quality=95)


def _update_movie_date(file_path: Path, dt: datetime) -> None:
    """
    Updates creation date field for movie files.
    :param file_path:
    :param dt:
    :return:
    """

    try:
        from mutagen.mp4 import MP4, MP4Tags
    except ImportError:
        raise ImportError(
            "mutagen library is required for movie file support. "
            "Install it with: pip install mutagen"
        )

    # Format as ISO 8601 date string (YYYY-MM-DDTHH:MM:SS)
    iso_date = dt.strftime("%Y-%m-%dT%H:%M:%S")

    # Open the MP4 file
    mp4_file = MP4(str(file_path))

    # Update creation date and modification date
    # '\xa9day' is the creation date tag
    # 'day' is also used for creation date
    mp4_file['\xa9day'] = [iso_date]

    # Also update the modification time
    mp4_file['\xa9mod'] = [iso_date]

    # Save the changes
    mp4_file.save()
    return None


def main(argv: Optional[List[str]] = None) -> int:
# Handles the parameters that were send with the command line
    args = parse_args(argv)
    configure_logging(args.verbose)

# Make sure the zip path that was sent exists and is a file (Not a folder)
    zip_path: Path = args.zip_path
    if not zip_path.exists():
        logging.error("ZIP not found: %s", zip_path)
        return 2
    if not zip_path.is_file():
        logging.error("Not a file: %s", zip_path)
        return 2

    try:
        # Creates the output folder based on the zip file name
        extraction_root = ensure_new_extraction_dir(zip_path, args.output_dir)
    except Exception as exc:
        logging.error("Failed to create extraction directory: %s", exc)
        return 2

    try:
        # Extracts the zip file into the new output folder
        safe_extract_zip(zip_path, extraction_root)
    except Exception as exc:
        logging.error("Extraction failed: %s", exc)
        return 2

    results = scan_and_update(extraction_root, dry_run=args.dry_run, case_insensitive=args.case_insensitive)

    updated = sum(1 for r in results if r.updated)
    would_update = sum(1 for r in results if (not r.updated and r.reason == 'dry-run'))
    missing_ts = sum(1 for r in results if r.reason == 'no-timestamp')
    failures = sum(1 for r in results if (not r.updated and r.reason not in ('dry-run', 'no-timestamp')))
    total_pairs = len(results)

    print(f"Extraction folder: {extraction_root}")
    print(f"Metadata/media pairs found: {total_pairs}")
    if args.dry_run:
        print(f"Would update: {would_update}")
    print(f"Updated: {updated}")
    print(f"Missing timestamp: {missing_ts}")
    print(f"Failures: {failures}")

    return 0 if failures == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
