import chardet

# Map common detected encodings to names accepted by Polars CSV functions.
_POLARS_ENCODING_MAP = {
    "ascii": "utf8",  # ASCII is a subset of UTF-8
    "utf-8": "utf8",
    "utf-8-sig": "utf8",
    "utf8-sig": "utf8",
    "utf8": "utf8",
    "iso-8859-1": "latin1",
    "latin1": "latin1",
    "windows-1252": "windows-1252",
    "cp1252": "windows-1252",
    "gbk": "gbk",
    "gb2312": "gbk",
    "big5": "big5",
    "shift-jis": "shift-jis",
    "shift_jis": "shift-jis",
    "euc-jp": "euc-jp",
    "euc-kr": "euc-kr",
}


def detect_encoding(file_path: str) -> str:
    """Detect file encoding for CSV files.

    Normalizes the detected encoding to a name compatible with Polars CSV
    functions. UTF-8-SIG is mapped to 'utf8' because Polars handles the BOM
    automatically and does not accept the '-sig' suffix.

    Args:
        file_path: Path to the CSV file.

    Returns:
        The detected and normalized encoding string. Defaults to 'utf-8'
        if detection fails.
    """
    with open(file_path, "rb") as f:
        raw = f.read(50000)
        result = chardet.detect(raw)

    encoding = result.get("encoding") or "utf-8"

    # Normalize to lowercase with hyphens for consistent lookup.
    norm = encoding.lower().replace("_", "-")
    return _POLARS_ENCODING_MAP.get(norm, encoding)
