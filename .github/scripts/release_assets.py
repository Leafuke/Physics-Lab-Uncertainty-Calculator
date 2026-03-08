from __future__ import annotations

import argparse
import ast
import shutil
import textwrap
import zipfile
from pathlib import Path

try:
    import py7zr
except ImportError:  # pragma: no cover - installed by workflow
    py7zr = None


METADATA_FILE = Path("uncertainty_app/__init__.py")
REQUIRED_METADATA = {"APP_DISPLAY_NAME", "APP_VERSION"}


def read_metadata(root: Path) -> dict[str, str]:
    source_path = root / METADATA_FILE
    tree = ast.parse(source_path.read_text(encoding="utf-8"), filename=str(source_path))
    metadata: dict[str, str] = {}

    for node in tree.body:
        if not isinstance(node, ast.Assign) or len(node.targets) != 1:
            continue
        target = node.targets[0]
        if not isinstance(target, ast.Name) or target.id not in REQUIRED_METADATA:
            continue
        value = node.value
        if isinstance(value, ast.Constant) and isinstance(value.value, str):
            metadata[target.id] = value.value

    missing = REQUIRED_METADATA - metadata.keys()
    if missing:
        missing_names = ", ".join(sorted(missing))
        raise RuntimeError(f"Missing metadata in {source_path}: {missing_names}")

    version = metadata["APP_VERSION"]
    return {
        "app_name": metadata["APP_DISPLAY_NAME"],
        "version": version,
        "tag": f"v{version}",
    }


def write_github_output(output_path: Path, values: dict[str, str]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("a", encoding="utf-8") as handle:
        for key, value in values.items():
            handle.write(f"{key}={value}\n")


def resolve_bundle_dir(dist_dir: Path) -> tuple[Path, Path]:
    if dist_dir.exists():
        return dist_dir, dist_dir / "_internal"

    app_bundle = dist_dir.with_suffix(".app")
    if app_bundle.exists():
        return app_bundle, app_bundle / "Contents" / "Resources" / "_internal"

    raise FileNotFoundError(f"Bundled application directory not found: {dist_dir}")


def build_release_notes(app_name: str) -> str:
    return (
        textwrap.dedent(
            f"""\
            请按自己的系统下载对应版本：

            - Windows：下载 Windows 压缩包（.zip），解压后双击 {app_name}.exe 运行。
            - Linux：下载 Linux 压缩包（.7z），解压后运行同目录下的 {app_name}。
            - macOS：下载 macOS 压缩包（.7z），解压后运行同目录下的 {app_name}。

            使用说明：
            1. Windows 可直接用资源管理器解压 .zip。
            2. Linux / macOS 可使用 7-Zip、Keka、The Unarchiver 或 p7zip 解压 .7z。
            3. 解压后请保持主程序与 _internal 文件夹位于同一目录，再启动应用。
            4. macOS 首次运行若被系统拦截，请在“系统设置 > 隐私与安全性”中允许运行；Apple Silicon 设备若提示安装 Rosetta，按系统提示完成即可。
            """
        ).strip()
        + "\n"
    )


def create_zip_archive(stage_dir: Path, archive_path: Path) -> None:
    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(stage_dir.rglob("*")):
            archive.write(path, path.relative_to(stage_dir.parent))


def create_7z_archive(stage_dir: Path, archive_path: Path) -> None:
    if py7zr is None:
        raise RuntimeError("py7zr is required to create .7z archives")
    with py7zr.SevenZipFile(archive_path, "w") as archive:
        archive.writeall(stage_dir, arcname=stage_dir.name)


def command_version(args: argparse.Namespace) -> int:
    metadata = read_metadata(Path(args.root).resolve())
    if args.github_output:
        write_github_output(Path(args.github_output), metadata)
        return 0

    for key, value in metadata.items():
        print(f"{key}={value}")
    return 0


def command_notes(args: argparse.Namespace) -> int:
    root = Path(args.root).resolve()
    metadata = read_metadata(root)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(build_release_notes(metadata["app_name"]), encoding="utf-8")
    return 0


def command_package(args: argparse.Namespace) -> int:
    root = Path(args.root).resolve()
    metadata = read_metadata(root)

    raw_dist_dir = Path(args.dist_dir)
    dist_dir = raw_dist_dir if raw_dist_dir.is_absolute() else (root / raw_dist_dir)

    raw_icon_path = Path(args.icon_path)
    icon_path = raw_icon_path if raw_icon_path.is_absolute() else (root / raw_icon_path)

    raw_output_dir = Path(args.output_dir)
    output_dir = raw_output_dir if raw_output_dir.is_absolute() else (root / raw_output_dir)

    bundle_dir, internal_dir = resolve_bundle_dir(dist_dir)
    internal_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(icon_path, internal_dir / icon_path.name)

    stage_name = f"{metadata['app_name']}-{metadata['tag']}-{args.platform.lower()}-{args.arch.lower()}"
    stage_dir = output_dir / stage_name
    if stage_dir.exists():
        shutil.rmtree(stage_dir)

    output_dir.mkdir(parents=True, exist_ok=True)
    shutil.copytree(bundle_dir, stage_dir, symlinks=True)

    archive_extension = ".zip" if args.platform.lower() == "windows" else ".7z"
    archive_path = output_dir / f"{stage_name}{archive_extension}"
    if archive_path.exists():
        archive_path.unlink()

    if archive_extension == ".zip":
        create_zip_archive(stage_dir, archive_path)
    else:
        create_7z_archive(stage_dir, archive_path)

    try:
        asset_path = archive_path.relative_to(root).as_posix()
    except ValueError:
        asset_path = str(archive_path)

    outputs = {
        "app_name": metadata["app_name"],
        "version": metadata["version"],
        "tag": metadata["tag"],
        "asset_name": archive_path.name,
        "asset_path": asset_path,
    }
    if args.github_output:
        write_github_output(Path(args.github_output), outputs)
        return 0

    for key, value in outputs.items():
        print(f"{key}={value}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Release helper commands for GitHub Actions")
    subparsers = parser.add_subparsers(dest="command", required=True)

    version_parser = subparsers.add_parser("version", help="Read application version metadata")
    version_parser.add_argument("--root", default=".")
    version_parser.add_argument("--github-output")
    version_parser.set_defaults(func=command_version)

    notes_parser = subparsers.add_parser("notes", help="Write release notes")
    notes_parser.add_argument("--root", default=".")
    notes_parser.add_argument("--output", required=True)
    notes_parser.set_defaults(func=command_notes)

    package_parser = subparsers.add_parser("package", help="Stage and archive a bundled build")
    package_parser.add_argument("--root", default=".")
    package_parser.add_argument("--platform", required=True)
    package_parser.add_argument("--arch", required=True)
    package_parser.add_argument("--dist-dir", required=True)
    package_parser.add_argument("--icon-path", required=True)
    package_parser.add_argument("--output-dir", required=True)
    package_parser.add_argument("--github-output")
    package_parser.set_defaults(func=command_package)

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())