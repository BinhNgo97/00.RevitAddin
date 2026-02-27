from __future__ import annotations

from pathlib import Path

from pypdf import PdfReader


def extract_preview(pdf_path: Path, *, max_chars: int = 1200) -> None:
    reader = PdfReader(str(pdf_path))
    print(f"file: {pdf_path.name}")
    print(f"pages: {len(reader.pages)}")

    for i in range(min(3, len(reader.pages))):
        text = reader.pages[i].extract_text() or ""
        text = "\n".join(line.rstrip() for line in text.splitlines()).strip()
        print(f"\n--- page {i+1} ({len(text)} chars) ---")
        print(text[:max_chars])


def main() -> None:
    pdf_path = Path("AI - Xây dựng AI Agent.pdf")
    if not pdf_path.exists():
        raise SystemExit(f"PDF not found: {pdf_path}")

    extract_preview(pdf_path)


if __name__ == "__main__":
    main()
