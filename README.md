# shu-academic-poster

上海大学量子科技研究院学术报告海报生成技能（DSH skill，**1 号模板**）：把一份学术报告的 Word 文档，填进研究院固定版式的单页海报，输出**可编辑 PPTX + PNG + PDF**。

> 2 号模板（物理系版式，20×30 cm，蓝色页眉 + 华文彩云抬头 + 页脚带链接）是另一个仓库：`shu-physics-poster`。

## 两个模板怎么选

| | **1 号（本仓库）** | 2 号（`shu-physics-poster`） |
|---|---|---|
| 抬头 | 华文隶书「量子科技研究院」+ Calibri 英文 | 华文彩云「上海大学理学院物理系」+ Times 英文 |
| 「学术报告」 | 微软雅黑 加粗 红色 | 华文行楷 白色 |
| 画布 | 20 × 32 cm | 20 × 30 cm |
| `摘要` / `报告人简介` 标签 | 方块 蓝底白字 | 圆角渐变 蓝底华文行楷 |
| 页脚 | 蓝色底条，只放邀请人 | 左链接 + 右邀请人 |
| 摘要分框 | 一个框 | 两个框（绕开右上角照片） |

报告挂量子科技研究院 → 用 1 号；挂物理学科 Seminar → 用 2 号。

## 它解决什么

- **字体不串。** 正文中文宋体、西文 Times New Roman，每个 run 成对写 `latin` + `ea`；题目与信息区微软雅黑；抬头英文钉成模板的 Calibri。`check_fonts.py` 逐段比对生成稿与母版的字体组合，出现 `DRIFT` 就说明跑偏了。
- **上下标不丢。** 纯文本抽取会丢上下标（`Li_nB_n+1C_n-1`、`La3Ni2O7`、`U_SS(t_c)`），`extract_runs.py` 回 DOCX 把上下标位置取出来，写成独立 run。
- **缺字形提前拦。** `⟨⟩〈〉∼` 宋体和 Times New Roman 都没有，`check_text.py` 会在生成前拦下来，并给出替代方案。
- **排版参数化。** 长短内容都用命令行调（`--body-size`、`--leading`、`--shift-abstract`、`--box-height`…），不改 XML。

## 环境要求

| 依赖 | 说明 |
|---|---|
| Python 3.8+ | `scripts/*.py` **只用标准库**，不需要 python-pptx / lxml |
| Windows + PowerPoint | `render_poster.ps1` 通过 COM 导出 PNG/PDF，并给出文本边界 `qa.json` |
| Windows PowerShell 5.1 | `.ps1` 保持纯 ASCII，可直接跑 |

没有 PowerPoint 也能产出可编辑 PPTX，只是拿不到 PNG/PDF 与视觉核对。

## 用法

```powershell
# 1. 抽字段
python -X utf8 scripts\extract_docx.py "报告.docx" --out fields.json

# 2. 文本预检（连字符、%、缺字形、多余空格）
python -X utf8 scripts\check_text.py --data fields.json

# 3. 上下标 run（有上下标时才需要）
python -X utf8 scripts\extract_runs.py "报告.docx" --fields fields.json --out runs.json

# 4. 生成
python -X utf8 scripts\build_poster.py --data fields.json --runs runs.json --out 海报.pptx

# 5. 渲染
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\render_poster.ps1 -InputPptx 海报.pptx -OutputDir render -Jpg

# 6. 字体审计
python -X utf8 scripts\check_fonts.py 海报.pptx
```

再打开 `render\poster.png` 核对全文：字段、校徽、换行、摘要末行、简介末行、页脚。

## 目录结构

```
├── SKILL.md                   技能正文：边界、工作流、验收标准
├── agents/openai.yaml         其他 agent 平台的入口描述
├── assets/poster-template.pptx   母版（已脱敏，版式与原件一致）
├── assets/example-input.docx     自测用示例输入（虚构报告）
├── assets/reference.jpg          母版渲染图，用于版式比对
├── references/fields.md          字段与证据规则
├── references/fonts.md           字体对照表（从母版读出来的）
├── references/layout.md          shape ID 与适配范围
├── references/example-fields.json 与 example-input.docx 配套的字段
└── scripts/                      extract / check / build / render
```

## 自测

```powershell
python -X utf8 scripts\build_poster.py --data references\example-fields.json --out test.pptx --force
python -X utf8 scripts\check_fonts.py test.pptx     # 应输出 OK
```

## 已知限制

- 母版已脱敏：题目、摘要、简介、邀请人是占位文字（`某某` / `name@example.com`），只有版式、字体、配色是真实的。
- 渲染依赖本机 PowerPoint；无 PowerPoint 时不得声称已完成视觉核对。
- 中文字体（华文隶书、微软雅黑、宋体）需本机已安装，否则会走字体替换。

## 许可

MIT，见 [LICENSE](LICENSE)。
