"""Offline Python workbench. Runs user-authored code in its own local process.

This is a convenience runtime, not a permissions sandbox. Core stays pandas-free.
Copyright (c) 2026 AbeLin · MIT License
"""
import base64
from contextlib import redirect_stderr, redirect_stdout
from datetime import datetime
from decimal import Decimal
import io
import json
import os
from pathlib import Path
import re
import sys
import time
import traceback
from uuid import uuid4

from core.xlio import hidden_info
from desktop.runtime_libraries import environment


def clean_name(value):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(value)).strip().strip(".")[:100] or "未命名"
    if re.fullmatch(r"(?i)(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?", name):
        name = "_" + name
    return name


def unique_path(directory, filename):
    path = directory / filename
    index = 2
    while path.exists():
        path = directory / f"{Path(filename).stem}_{index}{Path(filename).suffix}"
        index += 1
    return path


class LogStream(io.TextIOBase):
    """Bound console output; keep both the UI and the on-disk log manageable."""
    def __init__(self, report, log_file):
        self.report, self.log_file = report, log_file
        self.pending = ""
        self.remaining = 200_000

    def write(self, text):
        length = len(text)
        if self.remaining <= 0:
            return length
        clipped = text[:self.remaining]
        self.remaining -= len(clipped)
        self.pending += clipped
        while "\n" in self.pending or len(self.pending) >= 2000:
            cut = self.pending.find("\n")
            cut = min(cut, 2000) if cut >= 0 else 2000
            line, self.pending = self.pending[:cut], self.pending[cut + (self.pending[cut:cut + 1] == "\n"):]
            self.send(line)
        if self.remaining == 0:
            self.flush()
            self.send("日志达到 200,000 字符上限，后续输出已省略。")
        return length

    def send(self, line):
        self.log_file.write(line + "\n")
        self.log_file.flush()
        self.report({"type": "log", "message": line})

    def flush(self):
        if self.pending:
            line, self.pending = self.pending, ""
            self.send(line)


def run_code(request, config, report):
    """Load the full selected Sheets, execute, and stream text/tables/PNG figures."""
    code = request.get("code")
    if not isinstance(code, str) or not code.strip() or len(code) > 200_000:
        raise ValueError("请输入 Python 代码（最多 200,000 字符）")
    output = request.get("outputPath")
    if not isinstance(output, str) or not output.strip():
        raise ValueError("请选择结果保存位置")
    root = Path(output).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    output_dir = root / f"Python_{datetime.now():%Y%m%d_%H%M%S}_{uuid4().hex[:6]}"
    output_dir.mkdir()
    report({"type": "started", "output": str(output_dir)})
    script_path = output_dir / "script.py"
    script_path.write_text(code, encoding="utf-8")
    started = time.monotonic()
    previous_cwd = Path.cwd()
    original_stdin = sys.stdin
    error = None
    row_count = 0
    plt = None
    with (output_dir / "运行日志.txt").open("w", encoding="utf-8") as log_file:
        log = LogStream(report, log_file)
        try:
            os.chdir(output_dir)
            # No GUI backend, external interpreter, notebook server, or network needed.
            os.environ["MPLCONFIGDIR"] = str(output_dir / ".matplotlib")
            with redirect_stdout(log), redirect_stderr(log):
                import pandas as pd
                import numpy as np
                import numpy_financial as npf
                import matplotlib
                matplotlib.use("Agg", force=True)
                import matplotlib.pyplot as plt
                from matplotlib import font_manager
                installed = {font.name for font in font_manager.fontManager.ttflist}
                candidates = ["Microsoft YaHei", "SimHei", "Noto Sans CJK SC", "WenQuanYi Micro Hei"]
                plt.rcParams["font.sans-serif"] = [n for n in candidates if n in installed] + ["DejaVu Sans"]
                plt.rcParams["axes.unicode_minus"] = False
                env = environment()
                (output_dir / "运行环境.json").write_text(json.dumps(env, ensure_ascii=False, indent=2), encoding="utf-8")

                from desktop.bridge import inspect_book
                tables, inputs = {}, []
                source = Path(config["input_path"])
                for filename in config["selected_input_files"]:
                    path = Path(filename)
                    info = {s["name"]: s for s in inspect_book(path, config)["sheets"]}
                    hidden = hidden_info(filename)
                    for name in config["sheet_selection"][os.path.normcase(str(path))]:
                        header = info[name]["headerRow"]
                        frame = pd.read_excel(path, sheet_name=name, header=header - 1,
                                              engine="xlrd" if path.suffix.lower() == ".xls" else "openpyxl")
                        excluded = hidden.get(name, ("visible", set()))[1]
                        frame = frame.iloc[:, [i for i in range(len(frame.columns)) if i + 1 not in excluded]]
                        relative = path.relative_to(source).as_posix() if source.is_dir() else path.name
                        key = f"{relative} / {name}"
                        tables[key] = frame
                        inputs.append({"key": key, "path": str(path), "sheet": name, "headerRow": header,
                                       "rows": len(frame), "columns": [str(c) for c in frame.columns]})
                        row_count += len(frame)
                        print(f"已载入 {key}：{len(frame):,} 行 × {len(frame.columns)} 列")
                (output_dir / "输入说明.json").write_text(json.dumps(inputs, ensure_ascii=False, indent=2), encoding="utf-8")

                output_count = 0
                inline_bytes = 0

                def display(value, title="分析结果"):
                    nonlocal output_count
                    if not isinstance(value, (pd.DataFrame, pd.Series, dict)):
                        print(str(value)[:20_000])
                        return
                    if output_count >= 40:
                        print("界面最多展示 40 项结果，请使用 save_excel 保存更多数据。")
                        return
                    if isinstance(value, dict):
                        value = pd.Series(value)
                    if isinstance(value, pd.Series):
                        value = value.to_frame(name=value.name if value.name is not None else "值")
                    columns = [str(c)[:200] for c in value.columns[:40]]
                    def cell_text(cell):
                        if cell is None or cell is pd.NA or cell is pd.NaT:
                            return ""
                        return str(cell)[:1000]
                    data = {"kind": "table", "title": str(title)[:200], "columns": columns,
                            "rows": [[cell_text(c) for c in row] for row in value.iloc[:100, :40].itertuples(index=False, name=None)],
                            "index": [str(i)[:200] for i in value.index[:100]],
                            "totalRows": len(value), "totalColumns": len(value.columns)}
                    report({"type": "output", "data": data})
                    output_count += 1

                def save_excel(data, filename="结果.xlsx", index=False):
                    name = clean_name(filename)
                    if not name.lower().endswith(".xlsx"):
                        name += ".xlsx"
                    path = unique_path(output_dir, name)
                    if isinstance(data, dict):
                        with pd.ExcelWriter(path, engine="openpyxl") as writer:
                            used = set()
                            for title, frame in data.items():
                                label = re.sub(r"[\\/*?:\[\]]", "_", str(title))[:31] or "Sheet"
                                base, suffix = label, 2
                                while label.lower() in used:
                                    label = base[:31 - len(str(suffix)) - 1] + f"_{suffix}"
                                    suffix += 1
                                used.add(label.lower())
                                frame.to_excel(writer, sheet_name=label, index=index)
                    else:
                        data.to_excel(path, index=index, engine="openpyxl")
                    print(f"已保存：{path.name}")
                    return path

                def merge_tables(source_tables=None, source_column="来源"):
                    mapping = tables if source_tables is None else source_tables
                    if not mapping:
                        raise ValueError("没有可合并的表格")
                    frames = []
                    for key, frame in mapping.items():
                        copied = frame.copy()
                        if source_column:
                            if source_column in copied.columns:
                                raise ValueError(f"已有「{source_column}」列，请换一个 source_column 或传 None")
                            copied[source_column] = key
                        frames.append(copied)
                    return pd.concat(frames, ignore_index=True, sort=False)

                def split_table(data, column, prefix="拆分"):
                    if column not in data.columns:
                        raise ValueError(f"没有「{column}」列，可用字段：{list(data.columns)}")
                    paths = []
                    for key, frame in data.groupby(column, dropna=False, sort=False, observed=True):
                        label = "空值" if pd.isna(key) else str(key)
                        paths.append(save_excel(frame, f"{prefix}_{clean_name(label)}.xlsx"))
                    return paths

                def save_figure(figure=None, filename="图表.png"):
                    nonlocal output_count, inline_bytes
                    figure = figure if figure is not None else plt.gcf()
                    name = clean_name(filename)
                    if not name.lower().endswith(".png"):
                        name += ".png"
                    path = unique_path(output_dir, name)
                    figure.savefig(path, dpi=140, bbox_inches="tight")
                    print(f"已保存：{path.name}")
                    file_size = path.stat().st_size
                    if file_size <= 4_000_000 and inline_bytes + file_size <= 20_000_000 and output_count < 40:
                        data = "data:image/png;base64," + base64.b64encode(path.read_bytes()).decode("ascii")
                        report({"type": "output", "data": {"kind": "image", "title": path.name, "src": data}})
                        inline_bytes += file_size
                        output_count += 1
                    else:
                        print("图表已保存到结果目录，超过界面预览容量。")
                    plt.close(figure)
                    return path

                def show(*args, **kwargs):
                    for number in plt.get_fignums():
                        save_figure(plt.figure(number))

                plt.show = show
                # input() cannot wait for the IPC pipe. Scripts should use variables instead.
                sys.stdin = io.StringIO("")
                namespace = {"__name__": "__main__", "__file__": str(script_path),
                             "pd": pd, "np": np, "plt": plt, "npf": npf, "Decimal": Decimal, "df": next(iter(tables.values())),
                             "tables": tables, "inputs": inputs, "OUTPUT_DIR": output_dir,
                             "display": display, "save_excel": save_excel, "save_figure": save_figure,
                             "split_table": split_table, "merge_tables": merge_tables}
                print("开始运行脚本…")
                exec(compile(code, str(script_path), "exec"), namespace)
                show()
                print("脚本运行完成。")
        except BaseException:
            error = traceback.format_exc()
            log.write(error)
        finally:
            log.flush()
            sys.stdin = original_stdin
            os.chdir(previous_cwd)
            if plt is not None:
                plt.close("all")
    artifacts = sorted(str(p.relative_to(output_dir)) for p in output_dir.rglob("*")
                       if p.is_file() and not any(part.startswith(".") for part in p.relative_to(output_dir).parts))
    return {"output": str(output_dir), "status": "error" if error else "success", "error": error,
            "artifacts": artifacts[:500], "artifactCount": len(artifacts), "rows": row_count,
            "duration": round(time.monotonic() - started, 2)}
