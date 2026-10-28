"""Audit Git's tracked and unignored source allowlist; never scans private decks."""
from pathlib import Path
import re
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]


def source_files():
    output = subprocess.check_output(['git', 'ls-files', '-z', '--cached', '--others', '--exclude-standard'], cwd=ROOT)
    return sorted(set(p for p in output.decode('utf-8').split('\0') if p))


def audit():
    issues = []
    allowed = {'.py', '.md', '.txt', '.yml', '.yaml', '.toml', ''}
    patterns = [r'(?i)[A-Z]:[\\/]Users[\\/][^\s\\/]+',
                r'-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----',
                r'gh[pousr]_[A-Za-z0-9]{30,}', r'github_pat_[A-Za-z0-9_]{40,}']
    files = source_files()
    for name in files:
        path = ROOT / name
        if path.is_symlink() or not path.is_file():
            issues.append(name + ': not a regular file'); continue
        if path.suffix.lower() not in allowed or path.stat().st_size > 2_000_000:
            issues.append(name + ': unexpected type/size'); continue
        try:
            content = path.read_text(encoding='utf-8')
        except UnicodeError:
            issues.append(name + ': not UTF-8 text'); continue
        for pattern in patterns:
            if re.search(pattern, content):
                issues.append(name + ': potential credential or personal absolute path')
    if issues:
        print('\n'.join(issues)); return 1
    print(f'PASS: {len(files)} source files; no blocked types or matching secret/path patterns. Manual history/content review still required.')
    return 0


if __name__ == '__main__':
    sys.exit(audit())
