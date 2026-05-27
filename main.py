import json
import sys
from pathlib import Path


def self_test(output_path: str | None = None) -> int:
    ssl_info = None
    ssl_error = None
    try:
        import ssl

        ssl_info = getattr(ssl, "OPENSSL_VERSION", "unknown")
    except Exception as exc:
        ssl_error = f"{type(exc).__name__}: {exc}"

    result = {
        "python": sys.version,
        "executable": sys.executable,
        "ssl": ssl_info,
        "ssl_error": ssl_error,
        "requests": None,
        "ok": False,
    }
    try:
        import requests

        result["requests"] = {
            "version": requests.__version__,
            "file": getattr(requests, "__file__", "bundled"),
        }
        result["ok"] = ssl_error is None
    except Exception as exc:
        result["requests_error"] = f"{type(exc).__name__}: {exc}"

    text = json.dumps(result, ensure_ascii=False, indent=2)
    if output_path:
        Path(output_path).write_text(text, encoding="utf-8")
    else:
        print(text)
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        index = sys.argv.index("--self-test")
        target = sys.argv[index + 1] if len(sys.argv) > index + 1 else None
        raise SystemExit(self_test(target))
    from paper2excel.gui import main

    main()
