# -*- coding: utf-8 -*-
"""
CATIA V5 零件模态分析（Frequency Analysis）自动化工具

功能：
  1. 连接正在运行的 CATIA V5（或按需启动新实例）
  2. 对指定零件（默认当前活动零件）自动赋予材料（默认 Steel，含
     ELFINI 各向同性分析数据；若零件已有材料则跳过）
  3. 创建 CATAnalysis 分析文档并导入该零件
  4. 创建频率分析 Case（模态分析），按配置施加约束
  5. 设置模态数 / 频率范围（接口允许时）
  6. 求解并提取固有频率，导出 CSV / XLSX，并生成 HTML 基础报告

依赖：pywin32（win32com）、openpyxl（可选，用于 XLSX 导出）
运行：python frequency_analysis.py [--part 零件路径] [--material Steel] ...

注意：
  - 需要 Generative Structural Analysis (GSA/GPS) 或 ELFINI Structural
    Analysis (EST) 许可，否则创建分析文档或求解会失败。
  - 脚本与被控的 CATIA 必须运行在同一 Windows 用户、同一权限级别，
    否则无法通过 COM 取到运行中的实例（可用 --launch 启动新实例）。
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import threading
import time
import traceback
from pathlib import Path
from typing import Any, Iterable, Optional

import win32com.client

# Windows 控制台输出统一用 UTF-8，避免中文日志乱码
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# ---------------------------------------------------------------- 常量

CATIA_PROGID = "CATIA.Application"

# CATAnalysisSetType 枚举（按 IDL 声明顺序）
CAT_ANALYSIS_SET_IN = 0       # catAnalysisSetIn
CAT_ANALYSIS_SET_OUT = 1      # catAnalysisSetOut
CAT_ANALYSIS_SET_NEUTRAL = 2  # catAnalysisSetNeutral

# SystemService.Evaluate 语言
CAT_VBSCRIPT_LANGUAGE = 0

DEFAULT_CONFIG = {
    "material": "Steel",
    "family": "Metal",
    "catalog": r"D:\Dassault Systemes\B32\win_b64\startup\materials\Catalog.CATMaterial",
    "n_modes": 10,
    "max_frequency_hz": 0.0,     # 0 表示不限
    "mesh_size": 1.0,            # 0 表示保持默认网格大小
    "restraint": "none",         # whole=约束整个零件 | none=自由模态
    "faces": [],                 # restraint=faces 时显式指定的约束面名
    "images": ["Disp_Iso", "StressVonMises_Iso_Smooth"],  # 求解后创建的云图
    "close_analysis": False,     # 求解保存后自动关闭分析文档
    "output_dir": "output",
    "save_part": True,           # 赋材料后保存零件
    "save_analysis": True,       # 求解后保存 CATAnalysis
    "report": True,              # 生成 HTML 基础报告
}

KNOWN_CATALOG_PATHS = [
    r"D:\Dassault Systemes\B32\win_b64\startup\materials\Catalog.CATMaterial",
    r"D:\Dassault Systemes\B33\win_b64\startup\materials\Catalog.CATMaterial",
    r"D:\Program Files\Dassault Systemes\B32\win_b64\startup\materials\Catalog.CATMaterial",
    r"C:\Program Files\Dassault Systemes\B32\win_b64\startup\materials\Catalog.CATMaterial",
]


def _com_item(coll: Any, index: int) -> Any:
    """兼容部分版本集合需要 (index, search_mode) 两个参数的 Item 调用。"""
    try:
        return coll.Item(index)
    except Exception:
        # catAnalysisSetSearchAll = 3
        return coll.Item(index, 3)


def _iter_com(coll: Any) -> list[Any]:
    items: list[Any] = []
    try:
        n = int(coll.Count)
    except Exception:
        return items
    for i in range(1, n + 1):
        try:
            items.append(_com_item(coll, i))
        except Exception:
            continue
    return items


# ---------------------------------------------------------------- 日志

class Logger:
    """带阶段前缀的控制台日志。"""

    def __init__(self, debug: bool = False) -> None:
        self.debug_on = debug

    def info(self, msg: str) -> None:
        print(f"[INFO ] {msg}", flush=True)

    def warn(self, msg: str) -> None:
        print(f"[WARN ] {msg}", flush=True)

    def error(self, msg: str) -> None:
        print(f"[ERROR] {msg}", flush=True)

    def debug(self, msg: str) -> None:
        if self.debug_on:
            print(f"[DEBUG] {msg}", flush=True)


# ---------------------------------------------------------------- COM 工具

def dyn(obj: Any) -> Any:
    """转为晚绑定（dynamic）dispatch，避免 typelib 缺失导致方法不可见。"""
    return win32com.client.dynamic.Dispatch(obj)


class AutoClickYes:
    """后台监视 CATIA 模态弹窗并自动点击“是(Y)”。

    子文档从临时目录另存到输出目录时，CATIA 会弹出
    "Save As ... is referenced by other documents ... Do you want to proceed?"
    确认框（Win32 模态对话框）。该弹窗无法通过 COM 控制，
    这里用 Win32 API 轮询查找标题为 "Save As" 的窗口，
    枚举其按钮并点击“是”，实现无人值守。
    """

    def __init__(self, timeout: float = 60.0) -> None:
        self.timeout = timeout
        self._stop = threading.Event()
        self._thread: Optional[threading.Thread] = None

    def start(self) -> None:
        if self._thread is not None and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=2)

    def _run(self) -> None:
        import win32con
        import win32gui
        import win32process
        import psutil
        start = time.time()
        while not self._stop.is_set():
            if time.time() - start > self.timeout:
                return

            def enum_cb(hwnd: int, _: Any) -> bool:
                if self._stop.is_set():
                    return False
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                title = win32gui.GetWindowText(hwnd)
                if "Save As" not in title:
                    return True
                # 安全：只处理 CATIA 进程的窗口，避免误点其他应用
                # （Word/Excel 等）的保存确认框。
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    proc_name = psutil.Process(pid).name().lower()
                except Exception:
                    proc_name = ""
                if "cnext" not in proc_name and "catia" not in proc_name:
                    return True
                # 找到按钮并点击“是”
                def child_cb(ch: int, __: Any) -> bool:
                    if win32gui.GetClassName(ch) != "Button":
                        return True
                    text = win32gui.GetWindowText(ch)
                    if text.strip().startswith("是") or text.strip().lower() == "yes":
                        win32gui.PostMessage(ch, win32con.BM_CLICK, 0, 0)
                    return True

                win32gui.EnumChildWindows(hwnd, child_cb, 0)
                return True

            try:
                win32gui.EnumWindows(enum_cb, 0)
            except Exception:
                pass
            time.sleep(0.3)


def connect_catia(launch: bool = False, logger: Optional[Logger] = None) -> tuple[Any, bool]:
    """连接正在运行的 CATIA；找不到实例且 launch=True 时启动新实例。"""
    logger = logger or Logger()
    try:
        app = win32com.client.GetActiveObject(CATIA_PROGID)
        logger.info("已连接到正在运行的 CATIA V5 实例")
        return dyn(app), False
    except Exception:
        if not launch:
            raise RuntimeError(
                "未找到正在运行的 CATIA 实例。请先打开 CATIA，"
                "或使用 --launch 参数启动新实例。"
            )
        logger.warn("未找到运行中的 CATIA，正在启动新实例...")
        app = win32com.client.Dispatch(CATIA_PROGID)
        app.Visible = True
        return dyn(app), True


def cleanup_leftover_analysis(app: Any, logger: Logger) -> None:
    """关闭会话中未保存的空分析文档残留（Analysis2/4/6...）。

    偶发导入重试可能留下从未保存到磁盘的空 CATAnalysis 文档
    （Name 形如 AnalysisN.CATAnalysis 且 FullName 为相对路径），
    运行前统一清理，避免会话越积越多。
    """
    svc = app.SystemService
    script = (
        "Function CloseDoc(name)\n"
        "  On Error Resume Next\n"
        "  CATIA.Documents.Item(name).Close\n"
        "  CloseDoc = Err.Description\n"
        "End Function"
    )
    for d in list(app.Documents):
        try:
            name = str(d.Name)
            full = str(d.FullName)
        except Exception:
            continue
        if not name.lower().endswith(".catanalysis"):
            continue
        if not re.match(r"^analysis\d+\.catanalysis$", name, re.IGNORECASE):
            continue
        # 未保存到磁盘的空壳：FullName 为相对路径（不含盘符）
        if re.match(r"^[A-Za-z]:", full):
            continue
        try:
            svc.Evaluate(script, CAT_VBSCRIPT_LANGUAGE, "CloseDoc", [name])
            logger.info(f"已清理未保存的空分析文档: {name}")
        except Exception:
            pass


def get_target_part(app: Any, part_path: Optional[str], logger: Logger) -> tuple[Any, Any]:
    """返回 (part_document, part)。未指定路径时使用活动文档。"""
    if part_path:
        p = Path(part_path)
        if not p.exists():
            raise FileNotFoundError(f"零件文件不存在: {p}")
        # 若零件已在会话中打开，直接复用，避免 CATIA 弹
        # "is already open, do you want to reopen" 确认框。
        target = os.path.normcase(str(p))
        doc = None
        for d in app.Documents:
            try:
                if os.path.normcase(str(d.FullName)) == target:
                    doc = d
                    break
            except Exception:
                continue
        if doc is None:
            doc = win32com.client.Dispatch(app.Documents.Open(str(p)))
            logger.info(f"已打开零件: {p}")
        else:
            logger.info(f"零件已在会话中打开，直接复用: {p}")
    else:
        # 优先活动文档；若活动文档不是零件（如残留的分析文档），
        # 则自动选择会话中打开的 CATPart。
        active = None
        try:
            active = app.ActiveDocument
        except Exception:
            active = None
        doc = None
        if active is not None:
            try:
                active.Part
                doc = active
            except Exception:
                doc = None
        if doc is None:
            for d in app.Documents:
                try:
                    d.Part
                    doc = d
                    break
                except Exception:
                    continue
        if doc is None:
            raise RuntimeError(
                "CATIA 中没有打开的 CATPart，请先打开一个零件或指定 --part。"
            )
        logger.info(f"使用零件文档: {doc.Name}")
    # 用 COM 原对象验证 Part 接口（动态包装后的 .Part 属性访问不可靠）
    svc = app.SystemService
    script = (
        "Function HasPart(doc)\n"
        "  On Error Resume Next\n"
        "  Set p = doc.Part\n"
        "  If Err.Number <> 0 Then\n"
        "    HasPart = False\n"
        "  ElseIf p Is Nothing Then\n"
        "    HasPart = False\n"
        "  Else\n"
        "    HasPart = True\n"
        "  End If\n"
        "End Function"
    )
    try:
        is_part = bool(svc.Evaluate(script, CAT_VBSCRIPT_LANGUAGE, "HasPart", [doc]))
    except Exception:
        is_part = False
    if not is_part:
        raise RuntimeError(
            f"文档 {doc.Name} 不是 CATPart 零件文档，无法进行模态分析。"
        )
    return dyn(doc), dyn(doc.Part)


def get_material_manager(owner: Any) -> Any:
    """通过 CATMatManagerVBExt 扩展获取材料管理器（晚绑定）。"""
    return dyn(owner.GetItem("CATMatManagerVBExt"))


def get_material_on_part(app: Any, mm: Any, part: Any) -> Any:
    """读取零件上的材料（COM 输出参数通过 VBScript 中转）。"""
    svc = app.SystemService
    script = (
        "Function GetMat(mm, part)\n"
        "  Dim m\n"
        "  mm.GetMaterialOnPart part, m\n"
        "  Set GetMat = m\n"
        "End Function"
    )
    try:
        return svc.Evaluate(script, CAT_VBSCRIPT_LANGUAGE, "GetMat", [mm, part])
    except Exception:
        return None


def open_catalog_material(app: Any, catalog: str, family: str, material: str,
                          logger: Logger) -> tuple[Any, Any]:
    """打开材料库并读取指定材料；调用方负责在赋值完成后关闭 mdoc。"""
    if not os.path.exists(catalog):
        raise FileNotFoundError(f"材料库文件不存在: {catalog}")
    mdoc = app.Documents.Read(catalog)
    mdoc2 = dyn(mdoc)
    fam = mdoc2.Families.Item(family)
    mat = fam.Materials.Item(material)
    logger.info(f"材料库中读取材料: {family}/{material}")
    return mdoc, mat


def ensure_material(app: Any, part_doc: Any, part: Any, cfg: dict, logger: Logger) -> str:
    """零件没有材料时自动赋予；返回材料名。"""
    mm = get_material_manager(part)
    existing = get_material_on_part(app, mm, part)
    if existing is not None:
        try:
            name = existing.Name
        except Exception:
            name = "<unknown>"
        logger.info(f"零件已有材料: {name}，跳过赋材料")
        return name

    mdoc = None
    try:
        # 注意：赋值期间必须保持材料库文档打开，否则材料对象失效、调用静默失败
        mdoc, material = open_catalog_material(
            app, cfg["catalog"], cfg["family"], cfg["material"], logger
        )
        link_mode = 1  # 与材料库建立链接（官方示例用法）
        mm.ApplyMaterialOnPart(part, material, link_mode)
        try:
            mm.ApplyMaterialOnBody(part.MainBody, material, link_mode)
        except Exception as exc:
            logger.warn(f"对 MainBody 赋材料失败（不影响 Part 级材料）: {exc}")
        part.Update()
    finally:
        if mdoc is not None:
            try:
                mdoc.Close()
            except Exception:
                pass

    # 校验：材料必须真正赋上且带分析数据
    applied = get_material_on_part(app, mm, part)
    if applied is None:
        raise RuntimeError(
            "材料未能成功赋予零件（ApplyMaterialOnPart 未生效），"
            "请检查材料库路径与 CATIA 权限设置。"
        )
    try:
        if not applied.ExistAnalysisData():
            raise RuntimeError(f"材料 {applied.Name} 缺少分析域数据，无法用于模态分析。")
    except RuntimeError:
        raise
    except Exception as exc:
        logger.warn(f"校验材料分析数据时出错: {exc}")
    logger.info(f"已赋予材料: {cfg['family']}/{cfg['material']}")

    if cfg.get("save_part", True):
        try:
            part_doc.Save()
            logger.info("零件已保存（材料已写入文件）")
        except Exception as exc:
            logger.warn(f"零件保存失败: {exc}")
    return cfg["material"]


# ---------------------------------------------------------------- 分析流程

def create_frequency_case(app: Any, part_doc: Any, part: Any, cfg: dict,
                          logger: Logger) -> tuple[Any, Any, Any, Any]:
    """创建分析文档 + 频率 Case + 约束；返回 (analysis_doc, case, solution_set, manager)。"""
    # 偶发情况下零件会被重复导入（双 3D/双 material/双 mesh），
    # 检测到后关闭文档重建，最多重试 3 次。
    for attempt in range(1, 4):
        analysis_doc = None
        try:
            analysis_doc = dyn(app.Documents.Add("Analysis"))
            logger.info(f"已创建 CATAnalysis 文档（第 {attempt} 次尝试）")
            try:
                app.StartWorkbench("GPSCfg")
            except Exception as exc:
                logger.warn(f"切换 GPSCfg 工作台失败（继续尝试）: {exc}")

            manager = dyn(analysis_doc.Analysis)

            # 导入零件（优先文件导入，回退内存导入）
            imported = False
            try:
                part_path = str(part_doc.FullName)
                if part_path:
                    manager.ImportDefineFile(part_path, "CATAnalysisImport", ())
                    imported = True
                    logger.info(f"已导入零件文件: {part_path}")
            except Exception as exc:
                logger.warn(f"ImportDefineFile 失败（{exc}），改用内存导入...")
            if not imported:
                manager.Import(part_doc)
                logger.info("已通过内存导入零件")
            logger.debug(f"[mesh] 导入后网格部件数: {mesh_part_count(manager)}")
            logger.debug(f"[mesh] 导入后链接零件: {linked_part_names(manager)}")

            linked = linked_part_names(manager)
            if len(linked) > 1:
                raise RuntimeError(f"零件被重复导入 {len(linked)} 次: {linked}")

            model = manager.AnalysisModels.Item(1)
            set_mesh_size(manager, cfg, logger)
            ensure_single_mesh_part(manager, logger)
            logger.debug(f"[mesh] 设置网格大小后部件数: {mesh_part_count(manager)}")
            restraint = cfg.get("restraint", "whole")
            if restraint == "none":
                # 自由模态：使用官方 Free Frequency Case 模板（与 UI 一致）
                before = int(model.AnalysisCases.Count)
                model.RunTransition("CATGPSFreeModalAnalysis_template")
                after = int(model.AnalysisCases.Count)
                if after <= before:
                    raise RuntimeError("自由频率 Case 创建失败（RunTransition 未生效）")
                case = model.AnalysisCases.Item(after)
                solution_set = find_solution_set(case)
                logger.info("已创建 Free Frequency Case（自由模态，含刚体模态）")
            elif restraint == "whole":
                case = model.AnalysisCases.Add()
                sets = case.AnalysisSets
                restraint_set = sets.Add("RestraintSet", CAT_ANALYSIS_SET_IN)
                sets.Add("MassSet", CAT_ANALYSIS_SET_IN)
                solution_set = case.AddSolution("FrequencySet")
                sets.Add("SensorSet", CAT_ANALYSIS_SET_OUT)
                add_clamp_on_whole_part(manager, restraint_set, part_doc, part, logger)
                logger.info("已创建 Frequency Case（含固定约束）")
            elif restraint == "faces":
                # 约束用户标记的提取面：默认只处理名字以 fix_ 开头的
                # 混合形状（如 fix_1、fix_2），也可用 cfg["faces"] 显式
                # 指定名字列表。其余提取面（设计用途）一律忽略。
                case = model.AnalysisCases.Add()
                sets = case.AnalysisSets
                restraint_set = sets.Add("RestraintSet", CAT_ANALYSIS_SET_IN)
                sets.Add("MassSet", CAT_ANALYSIS_SET_IN)
                solution_set = case.AddSolution("FrequencySet")
                sets.Add("SensorSet", CAT_ANALYSIS_SET_OUT)
                add_clamp_on_extracted_faces(
                    manager, restraint_set, part_doc, part, cfg, logger
                )
                logger.info(
                    "已创建 Frequency Case（固定 fix 几何图形集 / 指定面）"
                )
            else:
                raise ValueError(
                    f"不支持的 restraint 值: {restraint}（支持 whole|none|faces）"
                )

            logger.debug(f"[mesh] 创建 Case 后网格部件数: {mesh_part_count(manager)}")
            set_frequency_parameters(manager, solution_set, cfg, logger)
            return analysis_doc, case, solution_set, manager
        except Exception as exc:
            if analysis_doc is not None:
                try:
                    analysis_doc.Close()
                    logger.warn("已关闭本次创建的分析文档")
                except Exception:
                    pass
            if attempt < 3:
                logger.warn(f"创建分析文档失败（{exc}），正在重试...")
            else:
                raise


def find_solution_set(case: Any) -> Any:
    """在 Case 中查找频率解集合（名称含 Solution 或类型含 Frequency）。"""
    sets = case.AnalysisSets
    for s in _iter_com(sets):
        try:
            name = str(s.Name)
        except Exception:
            name = ""
        try:
            typ = str(s.Type)
        except Exception:
            typ = ""
        if ("solution" in name.lower()) or ("frequenc" in typ.lower()):
            return s
    return None


def mesh_part_count(manager: Any) -> int:
    """返回当前分析模型中网格部件数量（诊断用）。"""
    try:
        model = manager.AnalysisModels.Item(1)
        return int(model.MeshManager.AnalysisMeshParts.Count)
    except Exception:
        return -1


def linked_part_names(manager: Any) -> list[str]:
    """返回分析文档链接的零件文档名（不含结果/计算文档）。"""
    names: list[str] = []
    try:
        ld = manager.LinkedDocuments
        for i in range(1, int(ld.Count) + 1):
            try:
                nm = str(ld.Item(i).Name)
            except Exception:
                continue
            if nm.lower().endswith((".catpart", ".catproduct")):
                names.append(nm)
    except Exception:
        pass
    return names


def ensure_single_mesh_part(manager: Any, logger: Logger) -> None:
    """防御：同一几何体只保留一个网格部件，多余的停用（避免双网格失真）。"""
    try:
        model = manager.AnalysisModels.Item(1)
        parts = model.MeshManager.AnalysisMeshParts
        n = int(parts.Count)
    except Exception as exc:
        logger.warn(f"无法检查网格部件数量: {exc}")
        return
    if n <= 1:
        return
    logger.warn(f"检测到 {n} 个网格部件，将停用多余的（保留第 1 个）")
    for i in range(2, n + 1):
        try:
            p = parts.Item(i)
            p.Activity = False
            logger.warn(f"已停用多余网格部件: {p.Name}")
        except Exception as exc:
            logger.warn(f"停用网格部件 {i} 失败: {exc}")


def set_mesh_size(manager: Any, cfg: dict, logger: Logger) -> None:
    """设置网格大小（mm）。参数名以 'Mesh Size' 结尾，可能有多个网格部件。"""
    mesh_size = float(cfg.get("mesh_size", 0.0))
    if mesh_size <= 0:
        return
    try:
        params = manager.Parameters
    except Exception as exc:
        logger.warn(f"无法读取参数集以设置网格大小: {exc}")
        return
    n_set = 0
    for p in _iter_com(params):
        try:
            name = str(p.Name)
        except Exception:
            continue
        if name.lower().endswith("\\mesh size"):
            try:
                p.Value = mesh_size
                n_set += 1
                logger.info(f"网格大小参数 {name} 设置为 {mesh_size} mm")
            except Exception as exc:
                logger.warn(f"设置 {name} 失败: {exc}")
    if n_set == 0:
        logger.warn("未找到 Mesh Size 参数，网格保持默认")


def add_clamp_on_whole_part(manager: Any, restraint_set: Any, part_doc: Any,
                            part: Any, logger: Logger) -> None:
    """对整个零件实体施加固定约束（SAMClamp）。"""
    clamp = restraint_set.AnalysisEntities.Add("SAMClamp")
    product = dyn(part_doc).Product

    # 通过拓扑搜索枚举实体全部面
    sel = part_doc.Selection
    sel.Clear()
    try:
        sel.Add(part.MainBody)
    except Exception as exc:
        logger.debug(f"Selection.Add(MainBody) 失败: {exc}")
    try:
        # 注意：Search 的返回值可能是 None，但结果会填充到 Selection 中
        sel.Search("Topology.Face,sel")
    except Exception as exc:
        logger.debug(f"Topology.Face,sel 搜索失败: {exc}")

    n = sel.Count
    added = 0
    for i in range(1, n + 1):
        item = sel.Item(i)
        ref = None
        try:
            ref = item.Reference
        except Exception:
            try:
                ref = part.CreateReferenceFromObject(item.Value)
            except Exception:
                ref = None
        if ref is None:
            continue
        try:
            clamp.AddSupportFromProduct(product, ref)
            added += 1
        except Exception as exc:
            logger.debug(f"第 {i} 个支撑面添加失败: {exc}")
    try:
        sel.Clear()
    except Exception:
        pass

    if added == 0:
        raise RuntimeError("未能为零件添加任何约束支撑面（请检查几何/许可）。")
    logger.info(f"已对整个零件施加固定约束（SAMClamp，共 {added} 个面）")


def add_clamp_on_extracted_faces(manager: Any, restraint_set: Any,
                                 part_doc: Any, part: Any, cfg: dict,
                                 logger: Logger) -> None:
    """对用户标记的提取面施加固定约束。

    默认约束名为 `fix` 的几何图形集（Geometrical Set）下的所有面，
    面自身无需改名；若 cfg["faces"] 提供了名字列表，则只处理列表中的面。
    """
    clamp = restraint_set.AnalysisEntities.Add("SAMClamp")
    try:
        product = dyn(part_doc).Product
    except Exception as exc:
        raise RuntimeError(f"无法获取零件 Product 对象: {exc}")
    explicit = [str(x).strip() for x in (cfg.get("faces") or [])]
    added = 0
    names: list[str] = []
    try:
        hybrid_bodies = part.HybridBodies
    except Exception as exc:
        raise RuntimeError(f"无法读取零件的几何体集: {exc}")
    logger.info(f"零件几何体集数量: {hybrid_bodies.Count}")
    for i in range(1, int(hybrid_bodies.Count) + 1):
        try:
            hb = hybrid_bodies.Item(i)
            set_name = str(hb.Name)
        except Exception:
            continue
        if explicit:
            logger.debug(f"遍历几何体集（按名字匹配）: {set_name}")
        elif set_name.strip().lower() != "fix":
            logger.debug(f"跳过非 fix 几何体集: {set_name}")
            continue
        logger.info(f"几何体集: {set_name}，全部面参与约束")
        try:
            shapes = hb.HybridShapes
        except Exception:
            continue
        logger.info(f"  提取形状数量: {shapes.Count}")
        for j in range(1, int(shapes.Count) + 1):
            try:
                hs = shapes.Item(j)
                name = str(hs.Name)
            except Exception:
                continue
            if explicit and name not in explicit:
                logger.debug(f"  跳过未指定的提取面: {name}")
                continue
            logger.info(f"  处理约束面: {name}")
            try:
                ref = part.CreateReferenceFromObject(hs)
            except Exception as exc:
                logger.warn(f"提取面 {name} 创建引用失败: {exc}")
                continue
            try:
                clamp.AddSupportFromProduct(product, ref)
                added += 1
                names.append(name)
                logger.info(f"已对提取面 {name} 施加固定约束")
            except Exception as exc:
                logger.warn(f"提取面 {name} 添加约束失败: {exc}")
    if added == 0:
        raise RuntimeError(
            "未找到可约束的提取面：请把要固定的面放入名为 fix 的"
            "几何图形集（Geometrical Set），或用 --faces 指定面名后重试。"
        )
    logger.info(f"固定约束完成（SAMClamp，共 {added} 个提取面: {names}）")


def set_frequency_parameters(manager: Any, solution_set: Any, cfg: dict,
                             logger: Logger) -> None:
    """在分析管理器/解集参数中设置模态数 / 频率范围（找不到则告警）。"""
    n_modes = int(cfg.get("n_modes", 10))
    max_hz = float(cfg.get("max_frequency_hz", 0.0))

    # 优先扫描分析管理器聚合参数（AnalysisSet 本身不暴露 Parameters）
    params = None
    try:
        params = manager.Parameters
    except Exception as exc:
        logger.warn(f"无法读取分析管理器参数集: {exc}")
        return

    mode_param = None
    freq_param = None
    for p in _iter_com(params):
        try:
            name = str(p.Name)
        except Exception:
            continue
        lower = name.lower()
        if ("mode" in lower and "number" in lower) or (
            "mode" in lower and "nb" in lower
        ):
            if mode_param is None:
                mode_param = (name, p)
        if "frequenc" in lower and ("range" in lower or "max" in lower or "limit" in lower):
            if freq_param is None:
                freq_param = (name, p)

    if mode_param is not None:
        try:
            mode_param[1].Value = n_modes
            logger.info(f"模态数参数 {mode_param[0]} 设置为 {n_modes}")
        except Exception as exc:
            logger.warn(f"设置模态数失败: {exc}")
    else:
        names = []
        for p in _iter_com(params):
            try:
                names.append(str(p.Name))
            except Exception:
                pass
        logger.warn(f"未找到模态数参数（保持默认；参数列表: {names}）")

    if max_hz > 0 and freq_param is not None:
        try:
            freq_param[1].Value = max_hz
            logger.info(f"频率上限参数 {freq_param[0]} 设置为 {max_hz} Hz")
        except Exception as exc:
            logger.warn(f"设置频率上限失败: {exc}")


def compute_case(case: Any, logger: Logger) -> None:
    logger.info("开始求解（网格划分 + ELFINI 求解器）...")
    t0 = time.time()
    case.Compute()
    logger.info(f"求解完成，耗时 {time.time() - t0:.1f} s")


def activate_result_image(analysis_doc: Any, cfg: dict, logger: Logger) -> bool:
    """求解后在求解集下创建指定云图并激活显示。

    CATIA 求解后不会自动生成云图；这里对 cfg["images"] 中每个标识符
    调用 AnalysisImages.Add 创建云图（iDuplicate=True 复用已有），
    Update 计算后激活，让 3D 视口显示云图。
    返回是否至少创建成功一张。
    """
    image_ids = [str(x).strip() for x in (cfg.get("images") or [])]
    if not image_ids:
        logger.info("未配置云图类型，跳过云图生成")
        return False
    try:
        manager = dyn(analysis_doc.Analysis)
        model = manager.AnalysisModels.Item(1)
        case = model.AnalysisCases.Item(1)
        sets = case.AnalysisSets
        sol_set = None
        for i in range(1, int(sets.Count) + 1):
            try:
                s = _com_item(sets, i)
                if "solution" in str(s.Type).lower() or "frequenc" in str(s.Type).lower():
                    sol_set = s
                    break
            except Exception:
                continue
        if sol_set is None:
            logger.warn("未找到求解集，跳过云图生成")
            return False
        images = sol_set.AnalysisImages
        # 清理早期无效的 Disp 占位图像（若存在）
        try:
            images.RemoveImage("Disp")
        except Exception:
            pass
        created = 0
        for image_id in image_ids:
            try:
                # iDuplicate=True：已存在同名图像时直接复用，否则新建
                img = images.Add(image_id, True, False, True)
                img.Update()
                img.SetActivationStatus(True)
                created += 1
                logger.info(f"云图已生成并激活: {img.Name} ({image_id})")
            except Exception as exc:
                logger.warn(f"创建云图 {image_id} 失败: {exc}")
        if created == 0:
            return False
        try:
            analysis_doc.Activate()
            time.sleep(1)
            app = analysis_doc.Application
            app.ActiveWindow.ActiveViewer.Reframe()
        except Exception:
            pass
        logger.info(f"共生成 {created} 张云图（3D 视口可查看）")
        return True
    except Exception as exc:
        logger.warn(f"生成云图失败: {exc}")
        return False


# ---------------------------------------------------------------- 结果提取

def _walk_collection(obj: Any, path: list[str],
                     sink: list[tuple[list[str], Any]], depth: int = 0) -> None:
    """递归收集分析树中的集合/实体（带深度上限，防止循环引用死循环）。"""
    if depth > 6:
        return
    for attr in ("AnalysisEntities", "AnalysisSets"):
        try:
            coll = getattr(obj, attr)
        except Exception:
            continue
        for item in _iter_com(coll):
            try:
                name = str(item.Name)
            except Exception:
                name = "<item>"
            child_path = path + [name]
            sink.append((child_path, item))
            _walk_collection(item, child_path, sink, depth + 1)


def extract_frequencies(app: Any, analysis_doc: Any, logger: Logger) -> list[float]:
    """从求解后的频率解中提取固有频率（Hz）。"""
    manager = dyn(analysis_doc.Analysis)
    model = manager.AnalysisModels.Item(1)
    case = model.AnalysisCases.Item(1)

    freqs: list[float] = []

    # 方式 1：分析管理器聚合参数（如 "...\Frequency\Frequency 1"）
    try:
        params = manager.Parameters
        for p in _iter_com(params):
            try:
                name = str(p.Name)
            except Exception:
                continue
            if not re.search(r"\\frequency\\frequency\s*\d+$", name, re.IGNORECASE):
                continue
            try:
                freqs.append(float(p.Value))
            except Exception:
                continue
    except Exception as exc:
        logger.debug(f"参数集扫描失败: {exc}")

    # 方式 2：树形结构中的频率传感器（兜底）
    collected: list[tuple[list[str], Any]] = []
    _walk_collection(case, [str(case.Name)], collected)
    for path, item in collected:
        leaf = path[-1] if path else ""
        if not re.match(r"^frequency\s*\d+$", leaf, re.IGNORECASE):
            continue
        outp = None
        for attr in ("OutPutParameters", "Parameters"):
            try:
                outp = getattr(item, attr)
                break
            except Exception:
                continue
        if outp is None:
            continue
        for p in _iter_com(outp):
            try:
                freqs.append(float(p.Value))
            except Exception:
                continue

    # 保持模态顺序去重（同一频率被多个来源读到时不重复计数）
    seen: set[float] = set()
    ordered: list[float] = []
    for f in freqs:
        if f not in seen:
            seen.add(f)
            ordered.append(f)
    return ordered


def generate_report(app: Any, analysis_doc: Any, out_dir: Path,
                    logger: Logger) -> Optional[Path]:
    """生成 HTML 基础报告（BuildReport 需要 CATIA Folder 对象）。"""
    try:
        manager = dyn(analysis_doc.Analysis)
        model = manager.AnalysisModels.Item(1)
        case = model.AnalysisCases.Item(1)
        post = dyn(model.PostManager)
        post.AddExistingCaseForReport(case)
        out_dir.mkdir(parents=True, exist_ok=True)
        folder = app.FileSystem.GetFolder(str(out_dir))
        title = "Frequency_Report"
        post.BuildReport(folder, title, False)
        html = out_dir / "index.html"
        if not html.exists():
            html = out_dir / f"{title}.html"
        logger.info(f"HTML 报告已生成: {html}")
        return html
    except Exception as exc:
        logger.warn(f"生成 HTML 报告失败: {exc}")
        return None


def parse_report_frequencies(html: Path) -> list[float]:
    """从 HTML 报告中解析频率数值（Hz）。"""
    if not html.exists():
        return []
    text = html.read_text(encoding="utf-8", errors="ignore")
    # 频率表：表头 "Mode number / Frequency Hz / Stability"，数值为科学计数法
    idx = re.search(r"frequency\s*<br>", text, re.IGNORECASE)
    if idx is None:
        return []
    seg = text[idx.start(): idx.start() + 8000]
    end = seg.find("</TABLE>")
    if end != -1:
        seg = seg[:end]
    rows = re.findall(
        r"<TD ALIGN=CENTER>\s*\d+\s*</TD>\s*<TD ALIGN=RIGHT>\s*([-+0-9.eE]+)\s*</TD>",
        seg,
        re.IGNORECASE,
    )
    # 保持表格顺序，与 CATIA 模态编号一一对应（不排序、不去重）
    return [float(f) for f in rows]


def export_csv(freqs: Iterable[float], path: Path, logger: Logger) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["mode", "frequency_hz"])
        for i, fq in enumerate(freqs, start=1):
            writer.writerow([i, f"{fq:.6g}"])
    logger.info(f"频率已导出 CSV: {path}")


def export_xlsx(freqs: Iterable[float], path: Path, logger: Logger) -> None:
    try:
        from openpyxl import Workbook
    except ImportError:
        logger.warn("未安装 openpyxl，跳过 XLSX 导出")
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "frequencies"
    ws.append(["mode", "frequency_hz"])
    for i, fq in enumerate(freqs, start=1):
        ws.append([i, float(fq)])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))
    logger.info(f"频率已导出 XLSX: {path}")


# ---------------------------------------------------------------- 主流程

def load_config(path: Optional[str]) -> dict:
    cfg = dict(DEFAULT_CONFIG)
    # 未显式指定 --config 时，自动读取程序同目录下的 config.json
    candidates = [path] if path else []
    candidates.append(str(Path(__file__).resolve().parent / "config.json"))
    for cand in candidates:
        if cand and os.path.exists(cand):
            with open(cand, encoding="utf-8") as f:
                cfg.update(json.load(f))
            break
    return cfg


def choose_analysis_path(app: Any, out_dir: Path, stem: str) -> Path:
    """选择不与会话中已打开文档/磁盘现有文件冲突的分析文档路径。"""
    session = set()
    for d in app.Documents:
        try:
            session.add(os.path.normcase(str(d.FullName)))
        except Exception:
            try:
                session.add(os.path.normcase(str(d.Name)))
            except Exception:
                pass
    n = 1
    while True:
        suffix = "" if n == 1 else f"_{n}"
        cand = out_dir / f"{stem}_frequency_analysis{suffix}.CATAnalysis"
        if os.path.normcase(str(cand)) not in session and not cand.exists():
            return cand.resolve()
        n += 1


def save_analysis_outputs(app: Any, analysis_doc: Any, analysis_path: Optional[Path],
                          run_dir: Path, part_doc: Any, part: Any,
                          stem: str, logger: Logger) -> Optional[Path]:
    """求解后保存结果/计算子文档与分析文档本体。

    若求解前 SaveAs 未成功（analysis_path 为 None），这里补一次 SaveAs。
    子文档另存可能触发 CATIA 模态弹窗，由 AutoClickYes 自动点击“是”。
    返回最终保存路径。
    """
    try:
        prefix = str(analysis_doc.Name)
        if prefix.lower().endswith(".catanalysis"):
            prefix = prefix[:-len(".CATAnalysis")]
        auto_click = AutoClickYes(timeout=120.0)
        auto_click.start()
        saved_docs: set[str] = set()
        for d in list(app.Documents):
            nm = str(d.Name)
            if nm.lower().startswith(prefix.lower()) and nm.lower().endswith(
                ("results", "computations")
            ):
                try:
                    key = os.path.normcase(str(d.FullName))
                except Exception:
                    key = os.path.normcase(nm)
                # 会话中可能有同名旧子文档（上一次分析的残留），去重跳过
                if key in saved_docs:
                    continue
                saved_docs.add(key)
                try:
                    full = str(d.FullName)
                    if full and os.path.normcase(os.path.dirname(full)) == os.path.normcase(str(run_dir)):
                        # 已在输出目录：直接保存
                        d.Save()
                    else:
                        # 子文档仍在临时目录：另存到输出目录
                        d.SaveAs(str(run_dir / nm))
                    logger.info(f"求解后已保存子文档: {nm} -> {d.FullName}")
                except Exception as exc:
                    logger.warn(f"保存子文档 {nm} 失败: {exc}")
        auto_click.stop()
        if analysis_path is not None:
            analysis_doc.Save()
        else:
            analysis_path = choose_analysis_path(app, run_dir, stem)
            analysis_doc.SaveAs(str(analysis_path))
        logger.info(f"求解后分析文档已保存: {analysis_path}")
        return analysis_path
    except Exception as exc:
        logger.warn(f"求解后保存分析文档失败: {exc}")
        return analysis_path


def find_catalog(cfg: dict) -> str:
    if os.path.exists(cfg["catalog"]):
        return cfg["catalog"]
    for p in KNOWN_CATALOG_PATHS:
        if os.path.exists(p):
            return p
    return cfg["catalog"]


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="CATIA V5 零件模态分析自动化",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    ap.add_argument("--part", help="CATPart 路径（默认：活动文档）")
    ap.add_argument("--config", help="config.json 路径")
    ap.add_argument("--material", help="材料名（默认 Steel）")
    ap.add_argument("--family", help="材料族（默认 Metal）")
    ap.add_argument("--catalog", help="材料库 CATMaterial 路径")
    ap.add_argument("--n-modes", type=int, help="计算模态数")
    ap.add_argument("--max-frequency", type=float, help="频率上限 Hz（0 不限）")
    ap.add_argument("--mesh-size", type=float, help="网格大小 mm（0 保持默认）")
    ap.add_argument("--restraint", choices=["whole", "none", "faces"],
                    help="约束方式（whole=整体固定 none=自由 "
                         "faces=固定 fix 几何图形集内的面）")
    ap.add_argument("--faces", nargs="*",
                    help="--restraint faces 时指定要约束的提取面名（默认 fix_* 前缀）")
    ap.add_argument("--images", nargs="*",
                    help="求解后创建的云图标识符"
                         "（默认 Disp_Iso StressVonMises_Iso_Smooth）")
    ap.add_argument("--output-dir", help="输出目录")
    ap.add_argument("--launch", action="store_true", help="无运行实例时启动新 CATIA")
    ap.add_argument("--dry-run", action="store_true", help="只检查零件/材料状态，不做任何修改")
    ap.add_argument("--skip-solve", action="store_true", help="建 Case 但不求解（调试用）")
    ap.add_argument("--no-save-part", action="store_true", help="不保存零件")
    ap.add_argument("--no-save-analysis", action="store_true", help="不保存 CATAnalysis")
    ap.add_argument("--no-report", action="store_true", help="不生成 HTML 报告")
    ap.add_argument("--close-analysis", action="store_true",
                    help="求解保存完成后自动关闭分析文档（磁盘文件保留）")
    ap.add_argument("--debug", action="store_true", help="输出调试信息")
    return ap.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    cfg = load_config(args.config)
    for key, val in [
        ("material", args.material),
        ("family", args.family),
        ("catalog", args.catalog),
        ("n_modes", args.n_modes),
        ("max_frequency_hz", args.max_frequency),
        ("mesh_size", args.mesh_size),
        ("restraint", args.restraint),
        ("faces", args.faces),
        ("images", args.images),
        ("output_dir", args.output_dir),
    ]:
        if val is not None:
            cfg[key] = val
    if args.no_save_part:
        cfg["save_part"] = False
    if args.no_save_analysis:
        cfg["save_analysis"] = False
    if args.no_report:
        cfg["report"] = False
    if args.close_analysis:
        cfg["close_analysis"] = True

    logger = Logger(debug=args.debug)
    analysis_doc = None
    out_dir = Path(cfg["output_dir"]).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        app, spawned = connect_catia(args.launch, logger)
        cleanup_leftover_analysis(app, logger)
        part_doc, part = get_target_part(app, args.part, logger)

        logger.info(f"零件: {part.Name} | 主几何体: {part.MainBody.Name}")
        stem = Path(str(part_doc.FullName or part.Name)).stem

        # 每次模拟在 output 下新建独立子目录，本次所有输出都放这里
        run_dir = out_dir / f"{stem}_{time.strftime('%Y%m%d_%H%M%S')}"
        run_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"本次模拟输出目录: {run_dir}")

        mm = get_material_manager(part)
        existing = get_material_on_part(app, mm, part)
        if args.dry_run:
            if existing is not None:
                logger.info(f"dry-run：零件已有材料 {existing.Name}，无需赋材料")
            else:
                logger.info(
                    f"dry-run：零件无材料，将赋予 {cfg['family']}/{cfg['material']}"
                )
            logger.info("dry-run 完成（未创建分析文档、未求解）")
            return 0

        if existing is not None:
            material_name = existing.Name
            logger.info(f"零件已有材料: {material_name}，跳过赋材料")
        else:
            material_name = ensure_material(app, part_doc, part, cfg, logger)

        # 关键：先保存零件，确保后续保存分析文档时不会触发
        # "Activates other document save operations" 联动保存弹窗。
        if cfg.get("save_part", True):
            try:
                part_doc.Save()
                logger.info("零件已保存（材料写入文件，避免联动保存弹窗）")
            except Exception as exc:
                logger.warn(f"零件保存失败: {exc}")

        analysis_doc, case, solution_set, manager = create_frequency_case(
            app, part_doc, part, cfg, logger
        )

        # 求解前先保存分析文档结构（此时没有结果子文档，可避免联动保存弹窗）
        analysis_path = None
        if cfg.get("save_analysis", True):
            analysis_path = choose_analysis_path(app, run_dir, stem)
            try:
                analysis_doc.SaveAs(str(analysis_path))
                logger.info(f"分析文档结构已保存: {analysis_path}")
            except Exception as exc:
                logger.warn(f"求解前保存分析文档失败: {exc}")
                analysis_path = None

        if not args.skip_solve:
            compute_case(case, logger)
            activate_result_image(analysis_doc, cfg, logger)
            freqs = extract_frequencies(app, analysis_doc, logger)
            html = None
            if cfg.get("report"):
                html = generate_report(app, analysis_doc, run_dir, logger)
            if not freqs and html is not None:
                logger.info("参数集未直接暴露频率值，尝试从 HTML 报告解析")
                freqs = parse_report_frequencies(html)
            if not freqs:
                logger.warn("未从分析树中提取到频率值，请检查求解结果")
            else:
                logger.info("固有频率（Hz）: " + ", ".join(f"{f:.4g}" for f in freqs))
                export_csv(freqs, run_dir / f"{stem}_frequencies.csv", logger)
                export_xlsx(freqs, run_dir / f"{stem}_frequencies.xlsx", logger)
        else:
            logger.info("--skip-solve：已建 Case，未求解")
            freqs = []

        if not args.skip_solve:
            analysis_path = save_analysis_outputs(
                app, analysis_doc, analysis_path, run_dir,
                part_doc, part, stem, logger,
            )

        if analysis_path is not None and args.skip_solve:
            try:
                analysis_doc.Save()
                logger.info("分析文档已保存")
            except Exception as exc:
                logger.warn(f"保存分析文档失败: {exc}")

        if analysis_path is not None:
            logger.info(f"分析文档路径: {analysis_path}")

        if cfg.get("close_analysis", False) and analysis_doc is not None:
            try:
                analysis_doc.Close()
                logger.info("已按 --close-analysis 关闭分析文档（磁盘文件保留）")
            except Exception as exc:
                logger.warn(f"关闭分析文档失败: {exc}")

        logger.info("完成。")
        return 0
    except Exception as exc:
        logger.error(f"运行失败: {exc}")
        if analysis_doc is not None:
            try:
                analysis_doc.Close()
                logger.warn("已关闭本次创建的分析文档")
            except Exception:
                pass
        if args.debug:
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
