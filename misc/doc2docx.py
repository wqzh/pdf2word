
# -*- coding: utf-8 -*-
"""
DOC -> DOCX 高性能转换工具

特点：
1. 纯 Python，不调用 Word / LibreOffice / COM / soffice
2. 使用 msdoc2docx 转换引擎
3. 支持单文件转换
4. 支持目录批量转换
5. 支持多进程并行转换
6. 支持跳过已转换文件
7. 支持 DOCX ZIP 完整性检查
8. 统计每个文件及整体转换耗时

安装：
    python -m pip install -U msdoc2docx

运行：
    python doc2docx_fast.py input.doc
    python doc2docx_fast.py input.doc -o output_dir
    python doc2docx_fast.py input_dir -o output_dir --workers 4
"""

from __future__ import annotations

import argparse
import os
import sys
import time
import zipfile

from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass
from typing import Optional

try:
    from doc2docx import convert as _doc2docx_convert
except ImportError as exc:
    raise ImportError(
        "缺少转换引擎，请执行："
        "python -m pip install -U msdoc2docx"
    ) from exc


class DocConversionError(RuntimeError):
    """DOC 转 DOCX 失败。"""


@dataclass
class ConversionResult:
    src: str
    dst: str
    success: bool
    elapsed: float
    message: str = ""


def _verify_docx(path: Path) -> bool:
    """检查 DOCX 是否为完整、可读取的 ZIP 文件。"""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            if zf.testzip() is not None:
                return False

            names = set(zf.namelist())

            return (
                "[Content_Types].xml" in names
                and "word/document.xml" in names
            )
    except (OSError, zipfile.BadZipFile):
        return False


def convert_doc_to_docx(
    src: Path,
    workdir: Path,
    *,
    password: Optional[str] = None,
    overwrite: bool = True,
    verify: bool = False,
) -> Path:
    """
    将旧版 DOC 转换为 DOCX。

    Parameters
    ----------
    src:
        输入 DOC 文件。
    workdir:
        输出目录。
    password:
        文档密码。
    overwrite:
        是否覆盖已有 DOCX。
    verify:
        是否执行 ZIP 完整性检查。

    Returns
    -------
    Path:
        转换后的 DOCX 文件路径。
    """
    start = time.perf_counter()

    src = Path(src)
    workdir = Path(workdir)

    if not src.is_file():
        raise DocConversionError(f"源文件不存在: {src}")

    if src.suffix.lower() != ".doc":
        raise DocConversionError(f"不支持的文件格式: {src}")

    workdir.mkdir(parents=True, exist_ok=True)

    dst = workdir / f"{src.stem}.docx"

    # 已有文件时跳过
    if dst.exists() and not overwrite:
        if not verify or _verify_docx(dst):
            return dst

        raise DocConversionError(
            f"已有输出文件不完整: {dst}"
        )

    try:
        # 核心转换：直接调用纯 Python 引擎
        result = _doc2docx_convert(
            str(src),
            str(dst),
            password=password,
        )

    except Exception as exc:
        raise DocConversionError(
            f"DOC 转 DOCX 失败: {src}\n"
            f"原因: {exc}"
        ) from exc

    if not dst.is_file():
        raise DocConversionError(
            f"转换结束，但未生成文件: {dst}"
        )

    if dst.stat().st_size == 0:
        raise DocConversionError(
            f"生成的 DOCX 文件为空: {dst}"
        )

    if verify and not _verify_docx(dst):
        raise DocConversionError(
            f"DOCX 完整性检查失败: {dst}"
        )

    elapsed = time.perf_counter() - start

    print(
        f"[完成] {src.name} -> {dst.name} "
        f"| {elapsed:.2f}s "
        f"| {dst.stat().st_size / 1024:.1f} KB",
        flush=True,
    )

    return dst


def _worker_convert(args) -> ConversionResult:
    """
    多进程工作函数。

    必须定义在模块顶层，以兼容 Windows
    multiprocessing 的 spawn 模式。
    """
    src, dst_dir, password, overwrite, verify = args

    start = time.perf_counter()

    try:
        output = convert_doc_to_docx(
            Path(src),
            Path(dst_dir),
            password=password,
            overwrite=overwrite,
            verify=verify,
        )

        return ConversionResult(
            src=str(src),
            dst=str(output),
            success=True,
            elapsed=time.perf_counter() - start,
        )

    except Exception as exc:
        return ConversionResult(
            src=str(src),
            dst="",
            success=False,
            elapsed=time.perf_counter() - start,
            message=str(exc),
        )


def batch_convert(
    input_dir: Path,
    output_dir: Path,
    *,
    workers: Optional[int] = None,
    password: Optional[str] = None,
    overwrite: bool = False,
    verify: bool = False,
    recursive: bool = True,
) -> list[ConversionResult]:
    """
    批量转换 DOC 文件。

    使用多进程，而不是多线程：
    DOC 解析和 XML 构建主要是 CPU 密集型工作，
    多进程可以绕过 CPython GIL 的限制。
    """
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)

    if not input_dir.is_dir():
        raise DocConversionError(
            f"输入目录不存在: {input_dir}"
        )

    output_dir.mkdir(parents=True, exist_ok=True)

    # 收集 DOC 文件
    iterator = (
        input_dir.rglob("*")
        if recursive
        else input_dir.glob("*")
    )

    files = sorted(
        p for p in iterator
        if p.is_file() and p.suffix.lower() == ".doc"
    )

    if not files:
        print("没有找到 DOC 文件")
        return []

    # 保留原始目录结构，防止同名文件冲突
    tasks = []

    for src in files:
        relative = src.relative_to(input_dir)

        dst_dir = output_dir / relative.parent
        dst = dst_dir / f"{src.stem}.docx"

        if dst.exists() and not overwrite:
            if not verify or _verify_docx(dst):
                print(f"[跳过] {src.name}")
                continue

        tasks.append((
            str(src),
            str(dst_dir),
            password,
            overwrite,
            verify,
        ))

    total = len(tasks)

    if total == 0:
        print("所有文件均已转换，无需处理")
        return []

    if workers is None:
        workers = min(os.cpu_count() or 1, 4)

    workers = max(1, min(workers, total))

    print("=" * 60)
    print(f"DOC 文件数量: {total}")
    print(f"并行进程数量: {workers}")
    print(f"输入目录: {input_dir}")
    print(f"输出目录: {output_dir}")
    print("=" * 60)

    start = time.perf_counter()
    results = []

    with ProcessPoolExecutor(
        max_workers=workers
    ) as executor:

        futures = {
            executor.submit(_worker_convert, task): task
            for task in tasks
        }

        for future in as_completed(futures):
            result = future.result()
            results.append(result)

            if result.success:
                print(
                    f"[{len(results)}/{total}] "
                    f"[成功] {Path(result.src).name} "
                    f"| {result.elapsed:.2f}s",
                    flush=True,
                )
            else:
                print(
                    f"[{len(results)}/{total}] "
                    f"[失败] {Path(result.src).name} "
                    f"| {result.message}",
                    flush=True,
                )

    elapsed = time.perf_counter() - start

    success = sum(r.success for r in results)
    failed = total - success

    print("\n" + "=" * 60)
    print("批量转换完成")
    print(f"总文件数: {total}")
    print(f"成功: {success}")
    print(f"失败: {failed}")
    print(f"总耗时: {elapsed:.2f}s")

    if success:
        print(f"平均成功文件耗时: "
              f"{sum(r.elapsed for r in results if r.success) / success:.2f}s")

    print("=" * 60)

    return results


def main():
    parser = argparse.ArgumentParser(
        description="纯 Python DOC -> DOCX 高性能转换工具"
    )

    parser.add_argument(
        "input",
        help="输入 DOC 文件或目录",
    )

    parser.add_argument(
        "-o", "--output",
        help="输出文件或目录",
    )

    parser.add_argument(
        "-w", "--workers",
        type=int,
        default=None,
        help="批量转换并行进程数，默认最多 4 个",
    )

    parser.add_argument(
        "--password",
        default=None,
        help="DOC 文档密码",
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="覆盖已有 DOCX 文件",
    )

    parser.add_argument(
        "--verify",
        action="store_true",
        help="转换后检查 DOCX ZIP 完整性",
    )

    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="批量转换时不递归子目录",
    )

    args = parser.parse_args()

    src = Path(args.input)

    if not src.exists():
        parser.error(f"输入路径不存在: {src}")

    try:
        if src.is_file():
            if args.output:
                out = Path(args.output)

                # 如果输出路径是目录，生成同名 DOCX
                if out.exists() and out.is_dir():
                    output_dir = out
                elif str(args.output).endswith(
                    ("/", "\\")
                ):
                    output_dir = out
                else:
                    output_dir = out.parent
                    output_name = out

                if "output_name" in locals():
                    result = _doc2docx_convert(
                        str(src),
                        str(output_name),
                        password=args.password,
                    )
                    if not output_name.is_file():
                        raise DocConversionError("输出文件未生成")
                    if args.verify and not _verify_docx(output_name):
                        raise DocConversionError("DOCX 完整性检查失败")
                    print(f"已生成: {output_name}")
                    return

            else:
                output_dir = src.parent

            output = convert_doc_to_docx(
                src,
                output_dir,
                password=args.password,
                overwrite=args.overwrite or True,
                verify=args.verify,
            )

            print(f"已生成: {output}")

        else:
            batch_convert(
                src,
                Path(args.output) if args.output else src / "docx_output",
                workers=args.workers,
                password=args.password,
                overwrite=args.overwrite,
                verify=args.verify,
                recursive=not args.no_recursive,
            )

    except DocConversionError as exc:
        print(f"[错误] {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    from time import time as timeit
    start_t = timeit()
    main()
    end_t = timeit()
    total_t = end_t - start_t
    print(total_t, ' s')
