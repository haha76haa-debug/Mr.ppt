# -*- coding: utf-8 -*-
"""渲染：抽单页预览 + 调 PowerPoint COM 真机导出。

LibreOffice 只能当快速预览：它忽略组合 3D 透视、文本描边/阴影等效果，
「预览对」不等于「交付对」。最终必须以 PowerPoint 导出为准。
"""
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import tempfile

from pptx import Presentation

HERE = os.path.dirname(os.path.abspath(__file__))
PS1 = os.path.join(HERE, "..", "..", "scripts", "export_slide.ps1")


def preview_deck(src, dst, keep):
    """只保留指定页，输出一个小 deck（渲染快，也让 PowerPoint 打得开超大文件）。

    keep: 1-based 页码列表。
    """
    keep = set(keep)
    prs = Presentation(src)
    lst = prs.slides._sldIdLst
    for i, sld in enumerate(list(lst), 1):
        if i in keep:
            continue
        rId = sld.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        prs.part.drop_rel(rId)
        lst.remove(sld)
    prs.save(dst)
    return dst


def has_powerpoint():
    return platform.system() == "Windows" and shutil.which("powershell") is not None


def export_slide_via_powerpoint(deck, out_png, slide=1, width=2001, height=1126,
                                script=PS1):
    """用 PowerPoint COM 导出单页 PNG（Windows + Office 才有）。"""
    if not has_powerpoint():
        raise RuntimeError("需要 Windows + Microsoft PowerPoint 才能做真机渲染")
    script = os.path.abspath(script)
    if not os.path.exists(script):
        raise FileNotFoundError(script)
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script,
           "-Path", os.path.abspath(deck), "-Out", os.path.abspath(out_png),
           "-Slide", str(slide), "-Width", str(width), "-Height", str(height)]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0 or not os.path.exists(out_png):
        raise RuntimeError("PowerPoint 导出失败: %s\n%s" % (proc.stdout, proc.stderr))
    return out_png


def export_slide_via_soffice(deck, out_png, slide=1):
    """备选的快速预览：用 LibreOffice headless 转 PDF/PNG（效果可能有差异）。"""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("未找到 LibreOffice（soffice）")
    tmp = tempfile.mkdtemp()
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp,
                    os.path.abspath(deck)], check=True, capture_output=True)
    pdfs = [os.path.join(tmp, f) for f in os.listdir(tmp) if f.endswith(".pdf")]
    if not pdfs:
        raise RuntimeError("LibreOffice 转换失败")
    # 调用者自行把 PDF 特定页转 PNG（例如 pdftoppm / PyMuPDF）
    return pdfs[0]
