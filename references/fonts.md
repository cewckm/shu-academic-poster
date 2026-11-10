# 字体与排版：以校方 PPT 为唯一标准

`assets/poster-template.pptx` 取自用户给的 `1106 课题组讲座（内部）4pm.pptx`，**已脱敏**：题目、摘要、简介、邀请人换成占位文字，页眉页脚、标签、字号与坐标保持原件不动（重建已交付海报验证过，逐像素一致）。字体、字号、颜色、位置、行距全部照这份 PPT，**不引入外部字体政策**（曾经按 Nature 风格改成微软雅黑/Arial，已回退）。

## 从这份 PPT 读到的实际用字

| 对象 | shape | 中文 | 西文 / 数字 | 字号 |
|---|---|---|---|---|
| 抬头「量子科技研究院」 | 6 第 1 段 | 华文隶书 | 华文隶书 | 28 pt |
| 抬头英文 | 6 第 2 段 | — | Calibri | 14 pt |
| 红色「学术报告」 | 10 | 微软雅黑 加粗 | 微软雅黑 | 28 pt，`#C00000` |
| 报告题目 | 14 | 微软雅黑 加粗 | 微软雅黑 | 20 pt |
| 报告人 / 时间 / 地点 | 25 | 微软雅黑 | 微软雅黑 | 16 pt |
| 标签「摘  要」「报告人简介」 | 15、16 | 主题字体（等线）加粗 | — | 18 pt，蓝底白字 |
| 摘要、简介正文 | 12、19 | 宋体 | Times New Roman | 16 pt，20 pt 行距 |
| 页脚 / 邀请人 | 22 | 主题字体（等线） | Helvetica | 14 pt，白字右对齐 |

## 三处需要写死的原因

1. **正文成对写 TNR + 宋体。** 模板正文是逐 run 混排（中文宋体的 run、英文 Times New Roman 的 run）；脚本重排后只保留一组 run 属性，必须自己成对写入 `a:latin=Times New Roman`、`a:ea=宋体`，否则英文会掉进宋体的西文字形。
2. **抬头英文钉成 Calibri。** 模板那里是 `+mj-lt`（主题西文字体 Aptos Display），不同 Office 版本替换结果不同；成功版呈现为 Calibri，所以显式写 Calibri，保证任何机器导出的 PNG/PDF 与批准版一致。
3. **标签、页脚不写字体。** 它们继续走模板主题（等线），改了反而偏离原版观感。

## 不做的事

- 不把正文换成微软雅黑/Arial 等外部字体，不套用 Nature/其他 skill 的字体政策。
- 不动版式：画布、色块、标签位置、页脚蓝条、字号与行距都照模板；只有当新内容确实放不下时，才用 `--body-size`、`--title-lines`、`--bio-top` 在规定范围内微调，并重新渲染核对。
- 上下标（Li<sub>n</sub>B<sub>n+1</sub>C<sub>n-1</sub>、La<sub>3</sub>Ni<sub>2</sub>O<sub>7</sub>、U<sub>SS</sub>(t<sub>c</sub>)）沿用同字体、65% 字号、`baseline="-25000"`。

## 可选覆盖

`--cn-font` / `--en-font` 只在用户明确要求换字体时使用。一旦使用，`.build.json` 会记录 `fonts.source = override`，`scripts/check_fonts.py` 会按模板报 DRIFT——这是有意偏离的提醒，交付时要向用户说明。

## 审计

```powershell
python -X utf8 "<skill目录>\scripts\check_fonts.py" "<输出目录>\海报.pptx"
```

脚本逐段比对生成稿与 `assets/poster-template.pptx` 的 `(a:latin, a:ea)` 组合：模板里出现过的组合才允许。两个白名单例外写在脚本顶部——抬头英文的 Calibri（模板成功版的实际呈现）和缺字形字符的 Cambria。任何 `DRIFT` 都说明字体被换掉或主题字体漏了出来。

## 缺字形与断行

- **正文按中英文拆 run。** 一个 run 同时挂 `latin=Times New Roman`、`ea=宋体` 时，英文虽然渲染成 TNR，但 PowerPoint 字体框里的“中文字体”仍显示宋体，容易被看成“公式还是宋体”。生成器会对摘要/简介做 `split_by_script`：中文 run 保持 `ea=宋体`，纯英文/公式 run 的 `ea` 也设成 `Times New Roman`、`lang=en-US`，这样点开公式，西文和中文两栏都显示 Times New Roman。
- **正文只允许 宋体 + Times New Roman。** 希腊字母（`κ`、`σ`、`Δ`、`ε`）TNR 都有，直接用。少数符号两个字体都没有（`⟨` U+27E8、`⟩` U+27E9、`〈` U+2329、`〉` U+232A、`∼` U+223C），PowerPoint 会静默换成别的字体，字体框里也显示成非新罗马——这就是“公式字体不对”的真实来源。本项目的要求是正文（含公式、英文）全部 Times New Roman，所以**把符号换成 TNR 自带的等价字符**：`⟩`→`>`、`⟨`→`<`、`∼`→`~`。`check_text.py` 会直接把这类字符报成错误。
  - 只有用户明确要求“保留原符号、接受字体切换”时，才用 `extract_runs.py --glyph-font Cambria` 给这几个字符单独指定 Cambria（Windows 上唯一含该字形的衬线体）。
- **断行**：`-5 %` 这类“负号 + 数字 + 空格 + %”会在行尾被拆开。数字前的负号用 **U+2011 NON-BREAKING HYPHEN**，数字与 `%` 之间用 **U+00A0 NO-BREAK SPACE**（两者 TNR 都有对应字形，外观与普通符号一致）。普通英文连字符（`d-wave`、`electron-phonon`）保持不动。
- **数字与单位**：写成 `44.7 K`、`5 %`，中间留一个空格；单位前用普通空格即可，只有 `%` 需要不换行（因为它紧跟在数字后、容易被拆开）。

## 审计

```powershell
python -X utf8 "<skill目录>\scripts\check_fonts.py" "<输出目录>\海报.pptx"
```

脚本逐段比对生成稿与 `assets/poster-template.pptx` 的 `(a:latin, a:ea)` 组合：模板里出现过的组合才允许。两个白名单例外写在脚本顶部——抬头英文的 Calibri（模板成功版的实际呈现）和缺字形字符的 Cambria。任何 `DRIFT` 都说明字体被换掉或主题字体漏了出来。

## 缺字形与断行

- **缺字形**：正文只用 宋体 + Times New Roman。有些符号两个字体都没有（`⟨` U+27E8、`⟩` U+27E9、`〈` U+2329、`〉` U+232A、`∼` U+223C），PowerPoint 会静默换成别的字体，在字体框里也会显示成非新罗马——这就是“公式字体不对”的真实来源。两种处理：
  1. **保留原符号**（默认，推荐）：给这几个字符单独建 run，字体设成 **Cambria**（衬线、Office 自带、Windows 上唯一含该字形的衬线体），其余字符仍是 Times New Roman。bra-ket 的形状与 Word 里一致。
  2. **只允许一种字体**：把 `⟩`→`>`、`⟨`→`<`，全式都在 TNR 里，但符号形状变成 ASCII 尖括号。只在用户明确要求“全篇只用一个西文字体”时采用。
- **断行**：`-5%` 这类“负号 + 数字”会在行尾被拆成 `-` 与 `5%`。改用 **U+2011 NON-BREAKING HYPHEN**（Times New Roman 有该字形，外观与半角连字符相同）。普通英文连字符（`d-wave`、`electron-phonon`）保持不动。
- **数字与符号之间**：按本项目习惯写成 `5 %`、`44.7 K`，数字与单位/百分号之间留一个半角空格。`scripts/check_text.py` 会把 `5%` 报成错误。
- 生僻符号（`κ`、`σ`、`Δ`、`ε` 等希腊字母）Times New Roman 都有，直接用即可。
