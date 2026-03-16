import subprocess
import sys
from pathlib import Path


def test_app_py_compiles() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    app_path = repo_root / "app.py"
    result = subprocess.run(
        [sys.executable, "-m", "py_compile", str(app_path)],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stderr
