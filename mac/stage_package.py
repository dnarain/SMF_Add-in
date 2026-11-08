"""Stage the SMF element files and package the native Mac port."""
from pathlib import Path
import argparse
import hashlib
import json
import shutil
import zipfile

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "mac/out"


def stage_elements():
    OUT.mkdir(parents=True, exist_ok=True)
    for source in sorted((ROOT / "docs/Elements").glob("*.txt")):
        shutil.copy2(source, OUT / source.name)


def package():
    stage_elements()
    files = []
    for folder in ("code", "licenses", "mac"):
        for source in sorted((ROOT / folder).rglob("*")):
            relative = source.relative_to(ROOT)
            if not source.is_file() or "__pycache__" in relative.parts:
                continue
            if source.suffix in (".zip", ".pyc") or source.name.startswith((".", "~$")):
                continue
            if relative.parts[:3] == ("mac", "diagnostics", "out"):
                continue
            if relative.parts[:2] == ("mac", "out") and source.name != "SMF-Mac.xlam" and source.suffix != ".txt":
                continue
            files.append(source)
    for source in sorted((ROOT / "docs/samples").glob("*")):
        if source.is_file() and not source.name.startswith("~$"):
            files.append(source)
    manifest = {
        "upstream": "https://github.com/ClimberMel/SMF_Add-in",
        "upstream_commit": "a0daad05699b1a7e213cf72ee641f30fc615af46",
        "native_addin_included": (OUT / "SMF-Mac.xlam").is_file(),
        "files": {str(p.relative_to(ROOT)): hashlib.sha256(p.read_bytes()).hexdigest()
                  for p in files if p.name != "package-manifest.json"},
    }
    manifest_file = ROOT / "mac/package-manifest.json"
    manifest_file.write_text(json.dumps(manifest, indent=2) + "\n")
    if manifest_file not in files:
        files.append(manifest_file)
    destination = ROOT / "SMF-Mac-port.zip"
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for source in files:
            archive.write(source, "SMF-Mac/" + str(source.relative_to(ROOT)))
    print(json.dumps({"archive": str(destination), "files": len(files),
                      "native_addin_included": manifest["native_addin_included"]}))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--elements-only", action="store_true")
    args = parser.parse_args()
    if args.elements_only:
        stage_elements()
        print("Staged element files beside the target add-in")
    else:
        package()
