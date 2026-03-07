from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QSettings, QStandardPaths

from .models import ProjectData

APP_ORG = "Leafuke"
APP_NAME = "UncertaintyCalculator"
MAX_RECENT_FILES = 10


def app_data_dir() -> Path:
    location = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppDataLocation)
    path = Path(location)
    path.mkdir(parents=True, exist_ok=True)
    return path


def autosave_file_path() -> Path:
    return app_data_dir() / "autosave.uncx"


def settings() -> QSettings:
    return QSettings(APP_ORG, APP_NAME)


def load_project_file(path: str) -> ProjectData:
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    project_data = payload.get("project", payload)
    project = ProjectData.from_dict(project_data)
    project.project_path = str(Path(path))
    return project


def save_project_file(project: ProjectData, path: str) -> ProjectData:
    project.ensure_defaults()
    now = datetime.now().isoformat(timespec="seconds")
    if not project.created_at:
        project.created_at = now
    project.updated_at = now

    payload = {
        "format_version": 1,
        "project": project.to_dict(),
    }

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    project.project_path = str(target)
    return project


def write_autosave(project: ProjectData) -> None:
    project.ensure_defaults()
    target = autosave_file_path()
    payload = {
        "format_version": 1,
        "project": project.to_dict(),
        "project_path": project.project_path,
    }
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_autosave() -> ProjectData | None:
    target = autosave_file_path()
    if not target.exists():
        return None

    payload = json.loads(target.read_text(encoding="utf-8"))
    project = ProjectData.from_dict(payload.get("project", payload))
    project.project_path = payload.get("project_path")
    return project


def recent_files() -> list[str]:
    raw_value = settings().value("recentFiles", [])
    if isinstance(raw_value, str):
        candidates = [raw_value]
    else:
        candidates = [str(item) for item in raw_value or []]

    existing_paths = [path for path in candidates if Path(path).exists()]
    settings().setValue("recentFiles", existing_paths)
    return existing_paths


def push_recent_file(path: str) -> list[str]:
    normalized = str(Path(path))
    items = [item for item in recent_files() if item != normalized]
    items.insert(0, normalized)
    items = items[:MAX_RECENT_FILES]
    settings().setValue("recentFiles", items)
    return items


def clear_recent_files() -> None:
    settings().setValue("recentFiles", [])


def last_project_path() -> str | None:
    value = settings().value("lastProjectPath", "")
    return str(value) if value else None


def set_last_project_path(path: str | None) -> None:
    settings().setValue("lastProjectPath", path or "")