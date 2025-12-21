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
import os
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import os
from datetime import datetime

try:
    import win32file
    import win32con
    import pywintypes
    WINDOWS_AVAILABLE = True
except ImportError:
    WINDOWS_AVAILABLE = False
    print("Warning: pywin32 not installed. Install it with: python -m pip install pywin32")



# ----------------------------- Logging Utilities -----------------------------


def configure_logging(verbosity: int) -> None:
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
    media_path: Path
    json_path: Path
    timestamp: int
    updated: bool
    reason: str = ""


# --------------------------------- Helpers -----------------------------------


def ensure_new_extraction_dir(zip_path: Path, output_dir: Optional[Path]) -> Path:
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
    # Accept patterns like:
    #   name.ext.json
    #   name.ext.metadata.json
    #   name.ext.supplemental-metadata.json
    if not json_filename.endswith('.json'):
        return None
    base = json_filename[:-5]  # strip .json
    for suffix in ('.supplemental-metadata', '.metadata'):
        if base.endswith(suffix):
            base = base[: -len(suffix)]
            break
    # base should now be the original media file name: name.ext
    if '.' not in base:
        # Likely not a Google Photos metadata naming pattern
        return None
    return base


def find_matching_media(json_path: Path, case_insensitive: bool) -> Optional[Path]:
    candidate_name = candidate_media_name_from_json(json_path.name)
    if not candidate_name:
        return None
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


def apply_timestamp_to_file(file_path: Path, timestamp: int, dry_run: bool) -> Tuple[bool, str]:
    try:
        if dry_run:
            return False, "dry-run"
        os.utime(file_path, (timestamp, timestamp))
        return True, ""
    except Exception as exc:
        return False, str(exc)


def scan_and_update(root: Path, *, dry_run: bool, case_insensitive: bool) -> List[UpdateResult]:
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
        updated, err = apply_timestamp_to_file(media_path, photo_taken_time, dry_run)
        if updated:
            logging.info("Updated %s from %s to %s", media_path, json_path, photo_taken_time)
            results.append(UpdateResult(media_path=media_path, json_path=json_path, timestamp=photo_taken_time, updated=True))
        else:
            if err == "dry-run":
                logging.info("Would update %s from %s to %s", media_path, json_path, photo_taken_time)
                results.append(UpdateResult(media_path=media_path, json_path=json_path, timestamp=photo_taken_time, updated=False, reason="dry-run"))
            else:
                logging.warning("Failed to update %s: %s", media_path, err)
                results.append(UpdateResult(media_path=media_path, json_path=json_path, timestamp=photo_taken_time, updated=False, reason=err))

        # Update the creation Date of the jpg file
        update_creation_date(str(media_path), creation_time)

    return results


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
    #timestamp = new_date.timestamp()
    filetime = pywintypes.Time(new_date)

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
    try:
        # Get current file times
        creation_time, access_time, write_time = win32file.GetFileTime(handle)

        # Update only the creation time, keep others unchanged
        win32file.SetFileTime(handle, filetime, access_time, write_time)
        print(f"Successfully updated creation date of '{file_path}' to {new_date}")
    finally:
        win32file.CloseHandle(handle)

    return None


# ----------------------------------- CLI -------------------------------------


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract ZIP and sync media timestamps from Google Photos JSON metadata.")
    parser.add_argument("zip_path", type=Path, help="Path to the ZIP archive (e.g., Google Takeout)")
    parser.add_argument("--output-dir", type=Path, default=None, help="Directory where a new extraction folder will be created.")
    parser.add_argument("--dry-run", action="store_true", help="Do not modify files, just report intended changes.")
    parser.add_argument("-v", "--verbose", action="count", default=0, help="Increase verbosity (use -v or -vv).")
    parser.add_argument("--case-insensitive", action="store_true", help="Allow case-insensitive filename matching if exact match is missing.")
    return parser.parse_args(argv)

def dry_run():
    # Example: Update creation date of a file
    file_to_update = "F:\\Project_output\\Output\\GooglePhotosTakeout_extracted_2\\Takeout\\Google Photos\\Photos from 2019\\20190921_124620_HDR.jpg"

    # Set new creation date (e.g., January 1, 2020, 12:00 PM)
    new_creation_date = datetime(2020, 1, 1, 12, 0, 0)
    try:
        update_creation_date(file_to_update, new_creation_date)
    except Exception as e:
        print(f"Error: {e}")
    return 0

def main(argv: Optional[List[str]] = None) -> int:

#   dry_run()
#   return 0

    args = parse_args(argv)
    configure_logging(args.verbose)

    zip_path: Path = args.zip_path
    if not zip_path.exists():
        logging.error("ZIP not found: %s", zip_path)
        return 2
    if not zip_path.is_file():
        logging.error("Not a file: %s", zip_path)
        return 2

    try:
        extraction_root = ensure_new_extraction_dir(zip_path, args.output_dir)
    except Exception as exc:
        logging.error("Failed to create extraction directory: %s", exc)
        return 2

    try:
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
