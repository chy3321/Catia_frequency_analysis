# CATIA V5 零件模态分析自动化（frequency_analysis）

通过 CATIA V5 的 COM 自动化接口（`CATAnalysisInterfaces` + `CATMatInterfaces`），
对零件自动完成：赋材料 → 建分析文档 → 频率分析 Case → 求解 → 提取固有频率
并导出 CSV/XLSX 与 HTML 报告。

## 环境要求

- Windows + CATIA V5（实测 V5-6R2024 / B32；B28/V5-6R2018 以上接口均适用）
- 许可：Generative Structural Analysis（GSA/GPS）或 ELFINI Structural
  Analysis（EST），否则创建分析文档或求解会失败
- Python 3.9+，依赖：

```powershell
pip install -r requirements.txt
```

## 用法

先打开 CATIA 和要分析的 CATPart（或直接指定路径），然后：

```powershell
# 分析当前活动零件
python frequency_analysis.py

# 指定零件与参数
python frequency_analysis.py --part "D:\parts\PLATE.CATPart" --n-modes 20

# 自由模态（不加约束）
python frequency_analysis.py --restraint none

# 约束模态：固定零件中名为 fix 的几何图形集内的所有面
python frequency_analysis.py --restraint faces --mesh-size 2.0

# 求解保存后自动关闭分析文档（磁盘文件保留，避免会话堆积）
python frequency_analysis.py --restraint faces --close-analysis

# 只检查零件与材料状态，不做修改
python frequency_analysis.py --dry-run
```

常用参数（详见 `python frequency_analysis.py --help`）：

| 参数 | 说明 | 默认 |
| --- | --- | --- |
| `--part` | 零件路径；缺省用 CATIA 活动文档 | 活动文档 |
| `--material` / `--family` | 材料名与材料族 | Steel / Metal |
| `--catalog` | 材料库 `.CATMaterial` 路径 | 自动探测 B32/B33 |
| `--n-modes` | 计算模态数 | 10 |
| `--max-frequency` | 频率上限 Hz（0=不限） | 0 |
| `--mesh-size` | 网格大小 mm（0=保持默认） | 1.0 |
| `--restraint` | `none`=自由模态 / `whole`=约束整个零件 / `faces`=固定提取面 | none |
| `--faces` | `--restraint faces` 时显式指定要约束的提取面名（默认固定名为 `fix` 的几何图形集内的所有面） | 空 |
| `--images` | 求解后创建的云图标识符（`Disp_Iso` 位移幅值 / `StressVonMises_Iso_Smooth` Von Mises 应力等） | `Disp_Iso StressVonMises_Iso_Smooth` |
| `--output-dir` | 输出目录 | output |
| `--launch` | 无运行实例时启动新 CATIA | 关 |
| `--dry-run` | 只读检查 | 关 |
| `--skip-solve` | 建 Case 不求解（调试） | 关 |
| `--close-analysis` | 求解保存完成后自动关闭分析文档（磁盘文件保留） | 关 |
| `--no-save-part` | 不保存零件（材料不写入文件） | 关 |
| `--no-save-analysis` | 不保存 CATAnalysis | 关 |
| `--no-report` | 不生成 HTML 报告 | 关 |
| `--debug` | 输出调试信息 | 关 |

输出文件（`output/`）：

- `<零件名>_frequencies.csv`：模态阶数与固有频率（Hz）
- `<零件名>_frequencies.xlsx`：同上（Excel）
- `<零件名>_frequency_analysis[_n].CATAnalysis`：可继续在 GSA 中查看的分析文档
  （若同名文件/文档已存在会自动加序号）
- `index.html`（+ 图片/css）：ELFINI 基础报告，频率表即解析来源

## 实现要点与限制

- 材料：程序通过 `Part.GetItem("CATMatManagerVBExt")` 获取材料管理器，从
  `Catalog.CATMaterial` 读取 Steel（自带 ELFINI 各向同性分析数据：
  E=200 GPa、ν=0.266、ρ=7860 kg/m³），赋给 Part 与 MainBody。若零件已有
  材料则跳过。赋材料后默认保存零件。
- 分析：`Documents.Add("Analysis")` 创建 CATAnalysis 文档，
  `StartWorkbench("GPSCfg")` 进入生成式结构分析，导入零件后自动生成
  3D 网格与材料属性；`AnalysisCases.Add()` + `AddSolution("FrequencySet")`
  创建模态分析 Case。
- 约束：默认自由模态，使用官方模板 `CATGPSFreeModalAnalysis_template`
  （与 UI 的 “Free Frequency Case” 一致，结果含刚体模态）；
  `--restraint whole` 会对整个零件实体施加固定约束
  （`SAMClamp`，通过拓扑搜索按面逐个添加支撑）。
- 约束模态（`--restraint faces`）：约束零件中**名为 `fix` 的几何图形集
  （Geometrical Set）内的所有面**，面自身无需命名规则；其他几何图形集
  （设计用途等）一律忽略。若不想建 `fix` 集，可用
  `--faces face1 face2` 显式指定面名。程序枚举几何图形集中的混合形状，
  对每个命中的面创建引用并施加 `SAMClamp` 固定约束。
- 结果：求解后生成 HTML 报告，并从中解析 “Mode number / Frequency Hz /
  Stability” 频率表（ELFINI 输出的科学计数法），导出 CSV/XLSX。
- 云图：求解后自动创建并激活位移幅值云图（`Disp_Iso`）与
  Von Mises 应力云图（`StressVonMises_Iso_Smooth`），3D 视口直接显示；
  可用 `--images` 自定义（如 `--images Disp_Iso` 只生成位移云图）。
  注意：模态分析通常只有位移解，Von Mises 应力云图在模态分析中可能
  创建失败（程序会告警并继续，不影响位移云图与结果导出）。
- 会话清理：每次运行前自动关闭会话中未保存的空分析文档残留
  （`AnalysisN.CATAnalysis` 空壳），避免越积越多；加 `--close-analysis`
  可在求解保存完成后关闭本次分析文档，只保留磁盘产物。
- 子文档保存：结果/计算子文档默认写入临时目录，程序会自动另存到本次
  输出子目录；CATIA 弹出的 “Save As 更新链接” 确认框由内置监视器自动
  点击“是”，全程无需人工干预。
- 网格防御：正常情况下导入零件只生成一个 3D 网格部件；个别会话状态下
  会出现两个相同的 OCTREE 网格同时作用于同一几何体（双网格），导致
  元素数量翻倍、频率结果严重失真。程序会在求解前检查
  `MeshManager.AnalysisMeshParts`，多于一个时自动停用多余的部件。
- COM 权限：脚本与 CATIA 必须同一 Windows 用户、同一权限级别运行；若
  CATIA 以管理员运行，脚本也需以管理员运行，否则取不到运行实例。

## 常见问题

- **“未找到正在运行的 CATIA 实例”**：CATIA 未开，或权限级别不同。打开
  CATIA 后重试；或加 `--launch`。
- **创建分析文档/求解失败（License）**：检查是否安装了 GSA/EST 许可。
- **没有提取到频率**：先打开保存的 `.CATAnalysis` 查看求解是否成功、
  频率传感器是否生成；再用 `--debug` 查看中间信息。
- **曾经的双 3D / 双 material / 双 mesh 问题**：早期运行结果只有 1e-3 Hz
  量级，根因是零件被偶发**重复导入两次**——`LinkedDocuments` 里出现两个
  METAL_PLATE.CATPart，每个实例各带一个 material（Property/Material 集
  各 2 个实体）和一个 OCTREE 网格（1mm 网格元素从约 70.9 万翻倍到
  142 万）。程序现在会在导入后检查链接零件数量，多于 1 个时自动关闭
  文档重建（最多 3 次），并把“只保留一个网格部件”作为第二道防线。
  修复后频率恢复正常量级（254/357/729/785 Hz）。
- **模态数**：频率解参数（Number of Modes）在当前版本未通过自动化接口
  暴露（`Frequency List\Size` 实测为 10），如需自定义请在求解前于 UI
  中设置 Frequency Solution Parameters，或调整模板默认。
- **“未找到可约束的提取面”**：`--restraint faces` 模式下程序只识别名为
  `fix` 的几何图形集（或 `--faces` 指定的面）。请先把要固定的面放入
  名为 `fix` 的 Geometrical Set（“提取（Extract）”后拖入该集），再重试。
