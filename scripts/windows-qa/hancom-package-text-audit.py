"""Compare two HWPX packages internally and emit only counts and element positions."""
from collections import Counter, defaultdict
from difflib import SequenceMatcher
import json
from pathlib import Path
import re
import sys
from xml.etree import ElementTree as XML
from zipfile import ZipFile


def read_package(file_path):
    counts = Counter()
    sections = defaultdict(list)
    section_name = re.compile(r"Contents/section([0-9]+)\.xml")
    with ZipFile(file_path) as package:
        for name in package.namelist():
            match = section_name.fullmatch(name)
            if not match:
                continue
            section = int(match[1])
            root = XML.fromstring(package.read(name))
            paragraph_count = 0

            def walk(element, role="body", paragraph=None):
                nonlocal paragraph_count
                name = element.tag.rsplit("}", 1)[-1]
                counts[name] += 1
                if name == "p":
                    paragraph_count += 1
                    paragraph = [section, paragraph_count, "", 0, 0]
                    sections[role].append(paragraph)
                if name == "t" and paragraph is not None:
                    paragraph[2] += "".join(
                        chunk for chunk in [element.text, *(child.tail for child in element)] if chunk
                    )
                    paragraph[3] += sum(child.tag.endswith("}tab") for child in element)
                    paragraph[4] += sum(child.tag.endswith("}lineBreak") for child in element)
                next_role = {
                    "tc": "cell", "subList": role, "drawText": "drawing", "caption": "caption",
                    "header": "header", "footer": "footer", "footNote": "footnote", "endNote": "endnote",
                }.get(name, role)
                for child in element:
                    walk(child, next_role, paragraph)

            walk(root)
    return counts, sections


def normalize(text):
    return "".join(text.split())


def audit_role(before, after):
    a = [normalize(record[2]) for record in before]
    b = [normalize(record[2]) for record in after]
    counter_a, counter_b = Counter("".join(a)), Counter("".join(b))
    paragraphs_a, paragraphs_b = Counter(value for value in a if value), Counter(value for value in b if value)
    unpaired = paragraphs_a - paragraphs_b
    first_missing = []
    for index, value in enumerate(a):
        if not unpaired[value]:
            continue
        first_missing.append([*before[index][:2], len(value)])
        unpaired[value] -= 1
        if len(first_missing) == 12:
            break
    by_role = {
        "logicalNonspace": [counter_a.total(), counter_b.total()],
        "missingCharacters": (counter_a - counter_b).total(),
        "additionalCharacters": (counter_b - counter_a).total(),
        "nonemptyParagraphs": [paragraphs_a.total(), paragraphs_b.total()],
        "unpairedParagraphs": [(paragraphs_a - paragraphs_b).total(), (paragraphs_b - paragraphs_a).total()],
        "orderedNonspaceEqual": "".join(a) == "".join(b),
        "orderedLiteralEqual": "".join(row[2] for row in before) == "".join(row[2] for row in after),
        "tabs": [sum(row[3] for row in before), sum(row[3] for row in after)],
        "lineBreaks": [sum(row[4] for row in before), sum(row[4] for row in after)],
        "firstUnpairedSourceAddressAndLength": first_missing,
    }
    if not by_role["orderedNonspaceEqual"]:
        source_rows = [(index, value) for index, value in enumerate(a) if value]
        saved_rows = [(index, value) for index, value in enumerate(b) if value]
        matcher = SequenceMatcher(a=[row[1] for row in source_rows],
                                  b=[row[1] for row in saved_rows], autojunk=False)
        edits = Counter()
        for tag, start, end, new_start, new_end in matcher.get_opcodes():
            if tag != "equal":
                edits[tag] += 1
        by_role["orderedParagraphEditSpans"] = edits
    return by_role


def report(source_path, saved_path):
    source_counts, source_roles = read_package(source_path)
    saved_counts, saved_roles = read_package(saved_path)
    return {
        "nodes": {key: [source_counts[key], saved_counts[key]]
                  for key in ("p", "tbl", "tc", "pic", "lineBreak", "tab", "header", "footer")},
        "roles": {role: audit_role(source_roles[role], saved_roles[role])
                  for role in sorted(source_roles.keys() | saved_roles.keys())},
    }


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("usage: hancom-package-text-audit.py source.hwpx saved.hwpx")
    print(json.dumps(report(Path(sys.argv[1]), Path(sys.argv[2])), ensure_ascii=True, separators=(",", ":")))
