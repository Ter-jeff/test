#!/usr/bin/env python3

# pyright: strict, reportUnknownMemberType=false, reportMissingTypeStubs=false

"""Publish a GitHub commit status with the branch-coverage percentage.

Runs from the Jenkinsfile's Coverage stage post block. A missing coverage report
or a failed status post must never fail the pipeline -- this is a PR-visible
nicety, not a gate -- so every failure path here prints a message and exits 0.
"""

import argparse
import json
import os
import re
import subprocess
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional

GITHUB_API = 'https://api.github.com'
REPO = 'Ter-jeff/test'
COVERAGE_CONTEXT = 'coverage/cobertura'
PREVIOUS_COVERAGE_LOOKBACK = 15


def resolve_git_sha() -> str:
    result = subprocess.run(['git', 'rev-parse', 'HEAD'], capture_output=True, text=True)
    return result.stdout.strip()


def fetch_previous_coverage(token: str, current_sha: str) -> Optional[float]:
    """Find the coverage % from the most recent prior commit that already has a status.

    The Statuses API only returns statuses for one commit per call, so this walks back
    through recent commit history (bounded by the shallow clone's depth) and asks for
    each one in turn, stopping at the first commit that already has a coverage/cobertura
    status. Returns None (never raises) if git history or every request comes up short --
    a missing baseline just means the description prints without a delta.
    """
    log = subprocess.run(
        ['git', 'log', '--format=%H', f'{current_sha}~1', '-n', str(PREVIOUS_COVERAGE_LOOKBACK)],
        capture_output=True, text=True,
    )
    if log.returncode != 0:
        return None

    for sha in log.stdout.split():
        req = urllib.request.Request(
            f'{GITHUB_API}/repos/{REPO}/commits/{sha}/status',
            headers={'Authorization': f'token {token}', 'Accept': 'application/vnd.github+json'},
        )
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read())
        except urllib.error.URLError:
            continue
        for status in data.get('statuses', []):
            if status.get('context') != COVERAGE_CONTEXT:
                continue
            match = re.match(r'([\d.]+)%', status.get('description', ''))
            if match:
                return float(match.group(1))
    return None


def find_branch_coverage(search_root: Path) -> Optional[float]:
    """Read branch coverage from ReportGenerator's aggregated Cobertura.xml.

    This is deliberately the *filtered* output (the Jenkinsfile's reportgenerator
    call passes -assemblyfilters:+*Automation*), not coverlet's raw per-run
    coverage.cobertura.xml under TestResults -- so this status and the Jenkins
    HTML coverage report always agree on scope without duplicating the filter here.
    """
    matches = sorted(search_root.rglob('Cobertura.xml'))
    if not matches:
        return None
    root = ET.parse(matches[0]).getroot()
    branch_rate = root.get('branch-rate')
    return float(branch_rate) * 100 if branch_rate is not None else None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('search_root', type=Path)
    args = parser.parse_args()
    search_root: Path = args.search_root

    token = os.environ.get('GITHUB_STATUS_TOKEN')
    if not token:
        print('GITHUB_STATUS_TOKEN not set; skipping coverage status publish.')
        return 0

    coverage = find_branch_coverage(search_root)
    if coverage is None:
        print(f'No usable coverage.cobertura.xml under {search_root}; skipping.')
        return 0

    sha = resolve_git_sha()
    previous_coverage = fetch_previous_coverage(token, sha)
    delta_text = f' ({coverage - previous_coverage:+.1f}%)' if previous_coverage is not None else ''

    build_url = os.environ.get('BUILD_URL', '')
    payload: dict[str, object] = {
        'state': 'success',
        'context': COVERAGE_CONTEXT,
        'description': f'{coverage:.1f}% branch coverage{delta_text}',
    }
    if build_url:
        payload['target_url'] = build_url + 'Unit_Test_Coverage/'

    req = urllib.request.Request(
        f'{GITHUB_API}/repos/{REPO}/statuses/{sha}',
        data=json.dumps(payload).encode(),
        headers={
            'Authorization': f'token {token}',
            'Accept': 'application/vnd.github+json',
            'Content-Type': 'application/json',
        },
        method='POST',
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            print(f'Published coverage status: HTTP {resp.status}')
    except urllib.error.URLError as exc:
        print(f'WARNING: failed to publish coverage status: {exc}')

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
