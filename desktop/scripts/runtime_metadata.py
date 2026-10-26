"""Keep installed package metadata/licences and a version manifest in offline builds."""
from importlib.metadata import distribution
import json
from pathlib import Path
import shutil
import sys

from packaging.requirements import Requirement
from packaging.utils import canonicalize_name

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from desktop.runtime_libraries import environment  # noqa: E402


def main():
    backend = Path(sys.argv[1]).resolve(strict=True)
    destination = (backend / "_internal").resolve(strict=True)
    runtime = environment()
    pending = [name for name in runtime["libraries"] if name != "decimal"]
    if sys.platform == "win32":
        pending.append("pywin32")
        runtime["excelPreview"] = {"bridge": "pywin32", "version": distribution("pywin32").version,
                                   "requires": "本机安装并激活的 Microsoft Excel"}
    copied = set()
    while pending:
        name = pending.pop()
        key = canonicalize_name(name)
        if key in copied:
            continue
        package = distribution(name)
        metadata = next(path for path in package.files or [] if path.name == "METADATA" and path.parent.name.endswith(".dist-info"))
        source = Path(package.locate_file(metadata)).parent
        target = (destination / source.name).resolve()
        target.relative_to(destination)
        shutil.copytree(source, target, dirs_exist_ok=True)
        copied.add(key)
        for dependency in package.requires or []:
            required = Requirement(dependency)
            if required.marker is None or required.marker.evaluate({"extra": ""}):
                pending.append(required.name)
    (backend / "离线依赖清单.json").write_text(json.dumps(runtime, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已保留 {len(copied)} 个依赖包的元数据、许可证与离线版本清单。")


if __name__ == "__main__":
    main()
