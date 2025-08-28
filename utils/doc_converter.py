from __future__ import annotations

import subprocess
from pathlib import Path


def convert_doc_to_docx(input_doc_path: str, output_directory: str | None = None) -> str:
    """
    Convert a legacy .doc file to .docx using LibreOffice (soffice) in headless mode.

    Args:
        input_doc_path: Path to the .doc file to convert.
        output_directory: Optional directory where the .docx should be written.
            Defaults to the parent of the input file.

    Returns:
        The absolute path to the generated .docx file.
    """

    input_path = Path(input_doc_path).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"File not found: {input_path}")
    if input_path.suffix.lower() != ".doc":
        raise ValueError(f"Input must be a .doc file, got: {input_path.suffix}")

    outdir = Path(output_directory).expanduser().resolve() if output_directory else input_path.parent
    outdir.mkdir(parents=True, exist_ok=True)

    # Run LibreOffice conversion
    subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to",
            'docx:"MS Word 2007 XML"',
            str(input_path),
            "--outdir",
            str(outdir),
        ],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

    output_path = outdir / f"{input_path.stem}.docx"
    if not output_path.exists():
        raise RuntimeError("LibreOffice reported success but .docx not found.")

    return str(output_path)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Convert .doc to .docx using LibreOffice")
    parser.add_argument("input_doc", help="Path to the .doc file")
    parser.add_argument("--outdir", help="Output directory for the .docx", default=None)
    args = parser.parse_args()

    result_path = convert_doc_to_docx(args.input_doc, args.outdir)
    print(result_path)


