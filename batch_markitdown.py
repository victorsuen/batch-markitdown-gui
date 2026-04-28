from __future__ import annotations

import argparse
import sys
from pathlib import Path

from markitdown import MarkItDown


SUPPORTED_EXTENSIONS = {
    ".pdf",
    ".pptx",
    ".docx",
    ".xlsx",
    ".xls",
    ".csv",
    ".json",
    ".xml",
    ".html",
    ".htm",
    ".txt",
    ".md",
    ".epub",
    ".zip",
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Batch convert files to Markdown using MarkItDown."
    )
    parser.add_argument(
        "input_dir",
        help="Source folder to scan for supported files.",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=str(Path.home() / "Desktop"),
        help="Folder to store Markdown output (default: Desktop).",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        default=True,
        help="Scan subfolders recursively (default: enabled).",
    )
    parser.add_argument(
        "--no-recursive",
        action="store_false",
        dest="recursive",
        help="Only scan top-level files in input folder.",
    )
    parser.add_argument(
        "--use-plugins",
        action="store_true",
        help="Enable MarkItDown third-party plugins.",
    )
    return parser


def discover_files(input_dir: Path, recursive: bool) -> list[Path]:
    pattern = "**/*" if recursive else "*"
    return sorted(
        file_path
        for file_path in input_dir.glob(pattern)
        if file_path.is_file() and file_path.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def make_output_path(src: Path, input_root: Path, output_root: Path) -> Path:
    relative = src.relative_to(input_root)
    return (output_root / relative).with_suffix(".md")


def convert_batch(
    source_files: list[Path],
    input_root: Path,
    output_root: Path,
    use_plugins: bool = False,
) -> tuple[int, int]:
    converter = MarkItDown(enable_plugins=use_plugins)
    success = 0
    failed = 0

    for src in source_files:
        dst = make_output_path(src, input_root, output_root)
        dst.parent.mkdir(parents=True, exist_ok=True)

        try:
            result = converter.convert(str(src))
            dst.write_text(result.text_content, encoding="utf-8")
            success += 1
            print(f"[OK] {src} -> {dst}")
        except Exception as exc:  # noqa: BLE001
            failed += 1
            print(f"[FAIL] {src}: {exc}", file=sys.stderr)

    return success, failed


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    input_dir = Path(args.input_dir).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()

    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Input folder does not exist or is not a folder: {input_dir}", file=sys.stderr)
        return 1

    files = discover_files(input_dir, recursive=args.recursive)
    if not files:
        print("No supported files found in the input folder.")
        print(
            "Supported extensions:",
            ", ".join(sorted(SUPPORTED_EXTENSIONS)),
        )
        return 0

    print(f"Source folder: {input_dir}")
    print(f"Output folder: {output_dir}")
    print(f"Found {len(files)} supported file(s). Start converting...\n")

    success, failed = convert_batch(
        source_files=files,
        input_root=input_dir,
        output_root=output_dir,
        use_plugins=args.use_plugins,
    )

    print("\nDone.")
    print(f"Successful: {success}")
    print(f"Failed: {failed}")
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
