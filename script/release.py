#!/usr/bin/env python3
"""
guellm release script
- tauri.conf.json / Cargo.toml 버전 bump
- pnpm tauri build
- git 커밋 & 태그 & push
- gh release create 로 GitHub Release 생성 및 산출물 업로드

사용법:
    python script/release.py                  # 패치 버전 bump (0.1.0 → 0.1.1)
    python script/release.py --minor          # 마이너 버전 bump (0.1.0 → 0.2.0)
    python script/release.py --major          # 메이저 버전 bump (0.1.0 → 1.0.0)
    python script/release.py --version 1.2.3  # 버전 직접 지정
    python script/release.py --no-build       # 빌드 스킵
    python script/release.py --no-push        # git push / release 스킵 (로컬 테스트용)

사전 조건:
    gh auth login  # 최초 1회
"""

import argparse
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
TAURI_CONF = ROOT / "src-tauri" / "tauri.conf.json"
CARGO_TOML = ROOT / "src-tauri" / "Cargo.toml"

ARTIFACT_PATTERNS = [
    "src-tauri/target/release/bundle/msi/*.msi",
    "src-tauri/target/release/bundle/nsis/*.exe",
    "src-tauri/target/release/bundle/msi/*.msi.zip",
    "src-tauri/target/release/bundle/nsis/*.exe.zip",
]


# ── 유틸 ──────────────────────────────────────────────────────────────────────

def run(cmd: list[str], **kwargs) -> subprocess.CompletedProcess:
    print(f"\n$ {' '.join(str(c) for c in cmd)}")
    result = subprocess.run(cmd, **kwargs)
    if result.returncode != 0:
        print(f"[ERROR] 명령 실패 (exit {result.returncode})")
        sys.exit(result.returncode)
    return result


def check_gh():
    result = subprocess.run(["gh", "auth", "status"], capture_output=True)
    if result.returncode != 0:
        print("[ERROR] gh CLI 로그인이 필요합니다.")
        print("  gh auth login")
        sys.exit(1)


# ── 버전 관리 ─────────────────────────────────────────────────────────────────

def read_version() -> str:
    return json.loads(TAURI_CONF.read_text(encoding="utf-8"))["version"]


def bump_version(current: str, mode: str) -> str:
    major, minor, patch = map(int, current.split("."))
    if mode == "major":
        return f"{major + 1}.0.0"
    elif mode == "minor":
        return f"{major}.{minor + 1}.0"
    else:
        return f"{major}.{minor}.{patch + 1}"


def write_version(version: str):
    # tauri.conf.json
    conf = json.loads(TAURI_CONF.read_text(encoding="utf-8"))
    conf["version"] = version
    TAURI_CONF.write_text(json.dumps(conf, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"tauri.conf.json  {version}")

    # Cargo.toml — [package] 섹션의 첫 번째 version = "..." 줄만 변경
    lines = CARGO_TOML.read_text(encoding="utf-8").splitlines()
    for i, line in enumerate(lines):
        if line.startswith('version = "'):
            lines[i] = f'version = "{version}"'
            break
    CARGO_TOML.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Cargo.toml       {version}")


# ── 빌드 ─────────────────────────────────────────────────────────────────────

def build():
    run(["pnpm", "tauri", "build"], cwd=ROOT)


def collect_artifacts() -> list[Path]:
    found = []
    for pattern in ARTIFACT_PATTERNS:
        found.extend(ROOT.glob(pattern))
    return found


# ── Git ───────────────────────────────────────────────────────────────────────

def git_commit_and_tag(version: str) -> str:
    tag = f"v{version}"
    run(["git", "add", str(TAURI_CONF), str(CARGO_TOML)], cwd=ROOT)
    run(["git", "commit", "-m", f"chore: release {tag}"], cwd=ROOT)
    run(["git", "tag", tag], cwd=ROOT)
    run(["git", "push", "origin", "HEAD"], cwd=ROOT)
    run(["git", "push", "origin", tag], cwd=ROOT)
    return tag


# ── GitHub Release (gh CLI) ───────────────────────────────────────────────────

def gh_release(tag: str, version: str, artifacts: list[Path]):
    if not artifacts:
        print("\n[WARNING] 업로드할 산출물이 없습니다. 빌드를 확인하세요.")

    cmd = [
        "gh", "release", "create", tag,
        "--title", f"guellm {tag}",
        "--notes", f"## guellm {tag}\n\n자동 릴리즈",
    ] + [str(a) for a in artifacts]

    run(cmd, cwd=ROOT)


# ── 메인 ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="guellm 빌드 & GitHub 릴리즈")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--major", action="store_true", help="메이저 버전 bump")
    group.add_argument("--minor", action="store_true", help="마이너 버전 bump")
    group.add_argument("--version", metavar="X.Y.Z", help="버전 직접 지정")
    parser.add_argument("--no-build", action="store_true", help="빌드 스킵 (기존 산출물 사용)")
    parser.add_argument("--no-push", action="store_true", help="git push 및 GitHub Release 스킵")
    args = parser.parse_args()

    if not args.no_push:
        check_gh()

    # 버전 결정
    current = read_version()
    if args.version:
        new_version = args.version
    elif args.major:
        new_version = bump_version(current, "major")
    elif args.minor:
        new_version = bump_version(current, "minor")
    else:
        new_version = bump_version(current, "patch")

    print(f"\n버전: {current} → {new_version}")
    if input("계속하시겠습니까? [y/N] ").strip().lower() != "y":
        print("취소됨.")
        sys.exit(0)

    write_version(new_version)

    if not args.no_build:
        build()

    artifacts = collect_artifacts()
    print(f"\n산출물 {len(artifacts)}개:")
    for a in artifacts:
        print(f"  {a.relative_to(ROOT)}")

    if args.no_push:
        print("\n--no-push: git push 및 GitHub Release 스킵")
        return

    tag = git_commit_and_tag(new_version)
    gh_release(tag, new_version, artifacts)

    print(f"\n릴리즈 완료: https://github.com/banatic/guellm/releases/tag/{tag}")


if __name__ == "__main__":
    main()
