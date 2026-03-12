from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any


LATEST_RELEASE_API_URL = "https://api.github.com/repos/Leafuke/Physics-Lab-Uncertainty-Calculator/releases/latest"
RELEASES_API_URL = "https://api.github.com/repos/Leafuke/Physics-Lab-Uncertainty-Calculator/releases?per_page=1"


@dataclass(frozen=True)
class ReleaseAsset:
    name: str
    download_url: str
    content_type: str = ""
    size: int = 0


@dataclass(frozen=True)
class ReleaseInfo:
    version: str
    title: str
    body: str
    html_url: str
    download_url: str
    published_at: str
    assets: list[ReleaseAsset] = field(default_factory=list)
    is_newer: bool = False


def parse_release_payload(payload: Any, current_version: str) -> ReleaseInfo:
    payload = _normalize_release_payload(payload)
    version = str(payload.get("tag_name") or payload.get("name") or "").strip()
    if not version:
        api_message = str(payload.get("message") or "").strip()
        if api_message:
            raise ValueError(f"GitHub API 返回：{api_message}")
        raise ValueError("未获取到有效的 GitHub Release 版本号。")

    title = str(payload.get("name") or version).strip()
    body = str(payload.get("body") or "").strip()
    html_url = str(payload.get("html_url") or "").strip()
    published_at = str(payload.get("published_at") or "").strip()

    assets = [
        ReleaseAsset(
            name=str(item.get("name") or ""),
            download_url=str(item.get("browser_download_url") or ""),
            content_type=str(item.get("content_type") or ""),
            size=int(item.get("size") or 0),
        )
        for item in payload.get("assets", [])
        if item.get("browser_download_url")
    ]

    return ReleaseInfo(
        version=version,
        title=title,
        body=body,
        html_url=html_url,
        download_url=_pick_download_url(assets, html_url),
        published_at=published_at,
        assets=assets,
        is_newer=is_newer_version(version, current_version),
    )


def is_newer_version(latest_version: str, current_version: str) -> bool:
    latest_key = _version_key(latest_version)
    current_key = _version_key(current_version)
    size = max(len(latest_key), len(current_key))
    return latest_key + (0,) * (size - len(latest_key)) > current_key + (0,) * (size - len(current_key))


def _pick_download_url(assets: list[ReleaseAsset], fallback_url: str) -> str:
    if not assets:
        return fallback_url

    priority_suffixes = (".exe", ".msi", ".zip", ".7z")
    for suffix in priority_suffixes:
        for asset in assets:
            if asset.name.lower().endswith(suffix):
                return asset.download_url

    return assets[0].download_url


def _normalize_release_payload(payload: Any) -> dict[str, Any]:
    if isinstance(payload, list):
        for item in payload:
            if isinstance(item, dict):
                return item
        raise ValueError("当前仓库还没有已发布的 GitHub Release。")

    if isinstance(payload, dict):
        return payload

    raise ValueError("GitHub Release 返回数据格式不正确。")


def _version_key(version_text: str) -> tuple[int, ...]:
    normalized = version_text.strip().lstrip("vV")
    numbers = re.findall(r"\d+", normalized)
    if not numbers:
        return (0,)
    return tuple(int(number) for number in numbers)