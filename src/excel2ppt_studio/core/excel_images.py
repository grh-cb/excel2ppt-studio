"""从 xlsx 包内提取 WPS「嵌入单元格图片」（DISPIMG）。

与 `excel2ppt` 技能同源，改为抛异常而非 ``sys.exit``，并支持内容哈希缓存。
"""

from __future__ import annotations

import hashlib
import os
import re
import zipfile


class ExcelImageError(Exception):
    """Excel 内嵌图片提取失败（可读中文信息）。"""


def extract_cell_images(
    xlsx_path: str, out_dir: str, only: set[str] | None = None
) -> dict[str, str]:
    """提取 ``xl/cellimages.xml`` 引用的图片，返回 ``{图片ID: 本地路径}``。

    ``only`` 非空时只落盘这些 ID，避免为预览一页解压整包 300MB 图片。
    图片按内容 SHA256 命名，重复内容只落盘一次。
    """
    if not os.path.isfile(xlsx_path):
        raise ExcelImageError(f"Excel 文件不存在：{xlsx_path}")
    os.makedirs(out_dir, exist_ok=True)

    id_to_file: dict[str, str] = {}
    skipped = 0
    with zipfile.ZipFile(xlsx_path, "r") as z:
        names = set(z.namelist())
        if "xl/cellimages.xml" not in names:
            raise ExcelImageError(
                "该 Excel 内没有 xl/cellimages.xml，未发现 WPS「嵌入单元格图片」。\n"
                "  典型原因：①图片是浮动图片而非嵌入单元格；②文件经其他软件转存丢失了 cellimages。\n"
                "  处理：在 WPS 中选中图片设为「嵌入单元格」后另存，或改用无图模式。"
            )
        if "xl/_rels/cellimages.xml.rels" not in names:
            raise ExcelImageError("Excel 缺少 xl/_rels/cellimages.xml.rels，cellimages 引用关系丢失。")

        ci = z.read("xl/cellimages.xml").decode("utf-8")
        ids = re.findall(r'name="(ID_[A-Fa-f0-9]+)"', ci)
        rids = re.findall(r'r:embed="(rId\d+)"', ci)

        rels = z.read("xl/_rels/cellimages.xml.rels").decode("utf-8")
        rid_to_media = dict(
            re.findall(r'Id="(rId\d+)"[^>]*Target="(media/image\d+\.\w+)"', rels)
        )

        for image_id, rid in zip(ids, rids, strict=False):
            if only is not None and image_id not in only:
                continue
            media = rid_to_media.get(rid)
            if not media:
                skipped += 1
                continue
            member = f"xl/{media}"
            if member not in names:
                skipped += 1
                continue
            data = z.read(member)
            digest = hashlib.sha256(data).hexdigest()[:16]
            ext = media.rsplit(".", 1)[-1].lower()
            path = os.path.join(out_dir, f"{digest}.{ext}")
            if not os.path.exists(path):
                with open(path, "wb") as f:
                    f.write(data)
            id_to_file[image_id] = path

    if skipped:
        # 调用方负责转成 WARNING；这里只保留计数，不静默
        id_to_file["__skipped__"] = str(skipped)
    return id_to_file


def list_dispimg_ids(xlsx_path: str) -> dict[str, str]:
    """只解析 ID → media 成员名，不解压图片（用于快速识别图片列）。"""
    with zipfile.ZipFile(xlsx_path, "r") as z:
        if "xl/cellimages.xml" not in z.namelist():
            return {}
        ci = z.read("xl/cellimages.xml").decode("utf-8")
        ids = re.findall(r'name="(ID_[A-Fa-f0-9]+)"', ci)
        return {i: "" for i in ids}
