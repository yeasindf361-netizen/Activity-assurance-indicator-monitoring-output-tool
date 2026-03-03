#!/usr/bin/env python3
"""Verify equivalence between two directories of output files.

Usage: verify_equivalence.py --baseline_dir <dir> --candidate_dir <dir>

Compares .txt/.log/.xlsx files as specified by the task.
"""
import argparse
import hashlib
import os
import sys
import zipfile
from pathlib import Path

CHUNK_SIZE = 8192


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as f:
        while True:
            chunk = f.read(CHUNK_SIZE)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def compare_xlsx(bpath: Path, cpath: Path, rel: str, mismatches, matches):
    # First try full-binary compare
    bh = sha256_file(bpath)
    ch = sha256_file(cpath)
    if bh == ch:
        matches.append(rel)
        return
    # If different, compare zipped members except docProps/core.xml
    try:
        with zipfile.ZipFile(bpath, 'r') as bz, zipfile.ZipFile(cpath, 'r') as cz:
            bnames = {n for n in bz.namelist() if not n.endswith('/')}
            cnames = {n for n in cz.namelist() if not n.endswith('/')}
    except Exception as e:
        mismatches.append((rel, 'xlsx-open-error', str(e)))
        return
    if bnames != cnames:
        mismatches.append((rel, 'xlsx-member-list-differ', sorted(list(bnames.symmetric_difference(cnames)))))
        return
    for member in sorted(bnames):
        if member == 'docProps/core.xml':
            # allowed to differ (timestamp-only diffs permitted)
            continue
        try:
            bdata = bz.read(member)
            cdata = cz.read(member)
        except Exception as e:
            mismatches.append((rel, 'xlsx-member-read-error', member))
            return
        if sha256_bytes(bdata) != sha256_bytes(cdata):
            mismatches.append((rel, 'xlsx-member-differ', member))
            return
    matches.append(rel)


def main():
    p = argparse.ArgumentParser(description='Verify equivalence between two directories')
    p.add_argument('--baseline_dir', required=True)
    p.add_argument('--candidate_dir', required=True)
    args = p.parse_args()

    base = Path(args.baseline_dir)
    cand = Path(args.candidate_dir)

    if not base.is_dir():
        print(f'Baseline dir not found: {base}', file=sys.stderr)
        sys.exit(2)
    if not cand.is_dir():
        print(f'Candidate dir not found: {cand}', file=sys.stderr)
        sys.exit(2)

    baseline_files = []
    for root, dirs, files in os.walk(base):
        for fn in files:
            f = Path(root) / fn
            rel = str(f.relative_to(base)).replace('/', '\\\\')
            baseline_files.append((f, rel))

    candidate_set = set()
    for root, dirs, files in os.walk(cand):
        for fn in files:
            f = Path(root) / fn
            rel = str(f.relative_to(cand)).replace('/', '\\\\')
            candidate_set.add(rel)

    mismatches = []
    matches = []

    baseline_rel_set = set(r for (_, r) in baseline_files)

    # First ensure sets match
    only_in_base = baseline_rel_set - candidate_set
    only_in_cand = candidate_set - baseline_rel_set
    if only_in_base:
        for r in sorted(only_in_base):
            mismatches.append((r, 'missing-in-candidate'))
    if only_in_cand:
        for r in sorted(only_in_cand):
            mismatches.append((r, 'extra-in-candidate'))

    # Compare each baseline file with candidate counterpart
    for bpath, rel in baseline_files:
        cpath = cand / Path(rel)
        if not cpath.exists():
            continue
        ext = bpath.suffix.lower()
        try:
            if ext == '.xlsx':
                compare_xlsx(bpath, cpath, rel, mismatches, matches)
            else:
                # For all other files (.txt, .log, etc.) do exact byte-hash compare
                bh = sha256_file(bpath)
                ch = sha256_file(cpath)
                if bh == ch:
                    matches.append(rel)
                else:
                    mismatches.append((rel, 'binary-differ'))
        except Exception as e:
            mismatches.append((rel, 'error', str(e)))

    # Print concise report
    print('\nEquivalence check report:')
    print(f'  Total baseline files: {len(baseline_files)}')
    print(f'  Matched files: {len(matches)}')
    print(f'  Mismatched files: {len(mismatches)}')
    if mismatches:
        print('\nMismatches (first 50 shown):')
        for item in mismatches[:50]:
            if isinstance(item, tuple):
                print('  -', item[0], ':', item[1], end='')
                if len(item) > 2:
                    print(':', item[2])
                else:
                    print()
            else:
                print('  -', item)

    sys.exit(0 if not mismatches else 1)

if __name__ == '__main__':
    main()
