import subprocess
import sys
import os


def run(cmd: list[str]) -> int:
    print("$", " ".join(cmd))
    return subprocess.call(cmd)


def main() -> int:
    # Ensure working directory is repo root
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    os.chdir(repo_root)

    # Install dev deps if missing (best-effort, optional)
    dev_deps_code = run([sys.executable, "-m", "pip", "install", "-r", "requirements-dev.txt"])  # noqa: E999
    if dev_deps_code != 0:
        print("Warning: Failed to install development dependencies (requirements-dev.txt). Subsequent steps may fail.")
    # Run tests with coverage (XML + HTML)
    code = run([
        sys.executable, "-m", "pytest", "test_excel_to_sql_converter.py", "-v",
        "--cov=excel_to_sql_converter", "--cov-report=xml", "--cov-report=html"
    ])
    if code != 0:
        print("Tests failed; coverage may be incomplete.")

    # Generate local badge (coverage.svg) in repo root if coverage-badge is installed
    run([sys.executable, "-m", "coverage_badge", "-o", "coverage.svg", "-f"])
    if badge_code != 0:
        print("Warning: Coverage badge generation failed. Is 'coverage-badge' installed?")

    print("\nCoverage report generated:")
    print(" - HTML: ./htmlcov/index.html")
    print(" - XML:  ./coverage.xml")
    print(" - Badge: ./coverage.svg (embed in README if desired)")
    return code


if __name__ == "__main__":
    raise SystemExit(main())
