#!/usr/bin/env python3
import os
import sys
import time
import tempfile
import requests

GITHUB_API_BASE = "https://api.github.com"
GITEE_API_BASE = "https://gitee.com/api/v5"
GITEE_MAX_ASSET_BYTES = 100 * 1024 * 1024  # 100 MB community/standard single asset limit


def require_env(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def fetch_github_releases(repo_full_name: str, github_token: str) -> list[dict]:
    headers = {
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    if github_token:
        headers["Authorization"] = f"Bearer {github_token}"

    releases = []
    page = 1
    while True:
        resp = requests.get(
            f"{GITHUB_API_BASE}/repos/{repo_full_name}/releases",
            headers=headers,
            params={"per_page": 100, "page": page},
            timeout=30,
        )
        resp.raise_for_status()
        batch = resp.json()
        if not batch:
            break
        releases.extend(batch)
        if len(batch) < 100:
            break
        page += 1
    return releases


def fetch_gitee_releases(owner: str, repo: str, token: str) -> dict[str, dict]:
    releases_by_tag: dict[str, dict] = {}
    page = 1
    while True:
        resp = requests.get(
            f"{GITEE_API_BASE}/repos/{owner}/{repo}/releases",
            params={"access_token": token, "per_page": 100, "page": page},
            timeout=30,
        )
        resp.raise_for_status()
        batch = resp.json()
        if not batch:
            break
        for item in batch:
            tag = item.get("tag_name")
            if tag:
                releases_by_tag[tag] = item
        if len(batch) < 100:
            break
        page += 1
    return releases_by_tag


def create_gitee_release(
    owner: str,
    repo: str,
    token: str,
    tag_name: str,
    name: str,
    body: str,
    target_commitish: str,
) -> dict:
    payload = {
        "access_token": token,
        "tag_name": tag_name,
        "name": name or tag_name,
        "body": body.strip() if body and body.strip() else "-",
        "target_commitish": target_commitish or "main",
    }
    resp = requests.post(
        f"{GITEE_API_BASE}/repos/{owner}/{repo}/releases",
        data=payload,
        timeout=30,
    )
    if not resp.ok:
        raise RuntimeError(f"Failed to create Gitee release for {tag_name}: HTTP {resp.status_code} {resp.text}")
    return resp.json()


def download_github_asset(url: str, github_token: str, dest_path: str) -> None:
    headers = {"Accept": "application/octet-stream"}
    if github_token:
        headers["Authorization"] = f"Bearer {github_token}"

    with requests.get(url, headers=headers, stream=True, timeout=120) as resp:
        resp.raise_for_status()
        with open(dest_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=64 * 1024):
                if chunk:
                    f.write(chunk)


def upload_gitee_asset(
    owner: str,
    repo: str,
    token: str,
    release_id: int | str,
    file_path: str,
    filename: str,
    max_retries: int = 3,
) -> None:
    url = f"{GITEE_API_BASE}/repos/{owner}/{repo}/releases/{release_id}/attach_files"
    last_err: Exception | None = None
    for attempt in range(1, max_retries + 1):
        try:
            with open(file_path, "rb") as f:
                resp = requests.post(
                    url,
                    data={"access_token": token},
                    files={"file": (filename, f, "application/octet-stream")},
                    timeout=180,
                )
            if not resp.ok:
                raise RuntimeError(f"HTTP {resp.status_code} {resp.text}")
            return
        except Exception as exc:
            last_err = exc
            if attempt < max_retries:
                wait_sec = 2 ** attempt
                print(f"Upload attempt {attempt} failed for {filename}: {exc}; retrying in {wait_sec}s...")
                time.sleep(wait_sec)
    raise RuntimeError(f"Failed to upload {filename} after {max_retries} attempts: {last_err}")


def main() -> int:
    gitee_token = os.environ.get("GITEE_TOKEN", "").strip()
    if not gitee_token:
        print("GITEE_TOKEN not configured; skipping release sync.")
        return 0

    repo_full_name = require_env("GITHUB_REPOSITORY")
    gh_default_owner, gh_default_repo = repo_full_name.split("/", 1)
    gitee_owner = os.environ.get("GITEE_OWNER", "").strip() or gh_default_owner
    gitee_repo = os.environ.get("GITEE_REPO", "").strip() or gh_default_repo
    github_token = os.environ.get("GITHUB_TOKEN", "").strip()
    target_tag = os.environ.get("TARGET_TAG", "").strip()

    gh_releases = fetch_github_releases(repo_full_name, github_token)
    published_releases = [r for r in gh_releases if not r.get("draft", False) and r.get("tag_name")]
    if target_tag:
        published_releases = [r for r in published_releases if r["tag_name"] == target_tag]
        if not published_releases:
            print(f"::warning::No published GitHub release found for tag '{target_tag}'.")
            return 0

    # Process oldest -> newest so Gitee release timeline order matches GitHub
    published_releases.sort(key=lambda r: r.get("created_at") or r.get("published_at") or "")
    gitee_releases = fetch_gitee_releases(gitee_owner, gitee_repo, gitee_token)

    print(
        f"Syncing {len(published_releases)} GitHub release(s) to "
        f"https://gitee.com/{gitee_owner}/{gitee_repo} (existing on Gitee: {len(gitee_releases)})"
    )

    for gh_rel in published_releases:
        tag_name = gh_rel["tag_name"]
        gitee_rel = gitee_releases.get(tag_name)

        if gitee_rel is None:
            print(f"Creating Gitee release for tag {tag_name}...")
            gitee_rel = create_gitee_release(
                owner=gitee_owner,
                repo=gitee_repo,
                token=gitee_token,
                tag_name=tag_name,
                name=gh_rel.get("name") or tag_name,
                body=gh_rel.get("body") or "-",
                target_commitish=gh_rel.get("target_commitish") or "main",
            )
            gitee_releases[tag_name] = gitee_rel
        else:
            print(f"Gitee release for tag {tag_name} already exists (id={gitee_rel['id']}); checking assets...")

        existing_asset_names = {
            asset.get("name")
            for asset in (gitee_rel.get("assets") or [])
            if asset.get("name")
        }

        gh_assets = gh_rel.get("assets") or []
        with tempfile.TemporaryDirectory() as tmpdir:
            for asset in gh_assets:
                asset_name = asset.get("name")
                download_url = asset.get("browser_download_url")
                asset_size = asset.get("size", 0)
                if not asset_name or not download_url:
                    continue
                if asset_name in existing_asset_names:
                    print(f"  [skip] {asset_name} already exists on Gitee")
                    continue
                if asset_size > GITEE_MAX_ASSET_BYTES:
                    print(
                        f"::warning::Skipping {asset_name} ({asset_size} bytes): "
                        f"exceeds Gitee 100 MB single-asset limit."
                    )
                    continue

                local_path = os.path.join(tmpdir, asset_name)
                print(f"  [sync] Downloading {asset_name} ({asset_size} bytes)...")
                download_github_asset(download_url, github_token, local_path)
                print(f"  [sync] Uploading {asset_name} to Gitee release {tag_name}...")
                upload_gitee_asset(
                    owner=gitee_owner,
                    repo=gitee_repo,
                    token=gitee_token,
                    release_id=gitee_rel["id"],
                    file_path=local_path,
                    filename=asset_name,
                )
                existing_asset_names.add(asset_name)
                print(f"  [done] {asset_name}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
