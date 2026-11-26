"""Fetch only publicly hosted, byte-pinned QA inputs into the runner's private temp directory."""
import argparse
import hashlib
import io
import json
import sys
import urllib.request
import zipfile
from pathlib import Path
from urllib.parse import urlparse


def verified_download(entry):
    if urlparse(entry["url"]).scheme != "https":
        raise ValueError("not-https")
    request = urllib.request.Request(entry["url"], headers={"User-Agent": "Mozilla/5.0 WordStudioQA/1.0"})
    with urllib.request.urlopen(request, timeout=45) as response:
        if urlparse(response.geturl()).scheme != "https":
            raise ValueError("not-https")
        data = response.read(entry["bytes"] + 1)
    if len(data) != entry["bytes"] or hashlib.sha256(data).hexdigest() != entry["sha256"]:
        raise ValueError("different-bytes")
    return data


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("group", choices=["word"] + [f"{label}-{n}" for label in ("hwp", "public") for n in range(1, 4)])
    parser.add_argument("directory", type=Path)
    args = parser.parse_args()
    sources = json.loads(Path(__file__).with_name("public-sources.json").read_text())
    label, _, part = args.group.partition("-")
    entries = sources[label]
    args.directory.mkdir(parents=True)
    if part:
        entries = entries[int(part)-1::3]
    archive = None
    if label == "word":
        try:
            archive = zipfile.ZipFile(io.BytesIO(verified_download(sources["wordArchive"])))
        except Exception as error:
            print(f"QA_INPUT_ARCHIVE unavailable {type(error).__name__}")
    inventory = []
    for entry in entries:
        row = {key: entry[key] for key in ("id", "ext")}
        try:
            if label == "word":
                if archive is None:
                    raise ValueError("archive-unavailable")
                data = archive.read(entry["member"])
                if hashlib.sha256(data).hexdigest() != entry["sha256"]:
                    raise ValueError("different-member")
            else:
                data = verified_download(entry)
            path = args.directory / f'{entry["id"]}.{entry["ext"]}'
            path.write_bytes(data)
            row["path"] = str(path.resolve())
            row["verified"] = True
            print(f'QA_INPUT {entry["id"]} hash-verified')
        except Exception as error:
            row["verified"] = False
            row["failure"] = type(error).__name__
            print(f'QA_INPUT {entry["id"]} unavailable {type(error).__name__}')
        inventory.append(row)
    (args.directory / "inventory.json").write_text(json.dumps(inventory, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
