# 排版助手

**把现有 Word 原稿和格式要求交给 Codex，得到另存的排版稿，并核对内容与页面。**

适合大学生整理课程报告、实验报告、论文初稿和作业。你可以提供学校规范、老师的文字要求或 Word 模板；没有规范时，也能按内置默认预设继续。使用时直接描述需求，不必自己写 Python。

这是一个供 Codex 使用的 **skill（技能指令与辅助工具）**。它指导 Codex 根据实际文档选择方法、调整格式和验收；具体效果仍取决于原稿、规范和运行环境。

## 能做什么

| 你遇到的问题 | 排版助手的处理范围 |
| --- | --- |
| 标题、正文格式不统一 | 按有效规则设置标题样式、字号、中英文字体、数字字体、对齐、缩进和行距，检查直接格式覆盖。 |
| 有规范但没有 Word 模板 | 直接依据规范排版，另存新的 DOCX。 |
| 图表缺题注、编号混乱 | 检查题注位置、编号顺序和正文引用；缺失时依据可见内容补中性题注与必要引用句，并记录新增内容。 |
| 公式展示和编号不一致 | 保留原有表达和数值，按规则处理独立展示、编号及相关引用，核对复杂公式对象。 |
| 图片后断页、页面大块空白 | 检查分页、段落关联和图片尺寸，保持纵横比，调整图题与正文接排。 |
| 担心排版改动正文 | 在原稿副本上修改，保全正文、数据、表格、公式、图片、注释、链接及其关系；比较内容，复核对象，说明差异。 |

排版范围包括必要的题注、图表引用句和编号呈现调整，这些会产生真实差异，交付时应单独记录。原有论述、数值与公式含义应保持；复杂对象的保全需要实际核对，不能仅凭“文字比较通过”认定整份文档无损。

## 三步上手

### 1. 安装技能

把下面这句话发给 Codex：

```text
请安装 https://github.com/01x01x/paiban-assistant 中的 word-format-from-content-template skill。
```

也可以下载本仓库的 [排版助手.zip](排版助手.zip)，告诉 Codex 文件保存的位置，让它从压缩包中安装 `word-format-from-content-template` 技能。安装后，按 Codex 的实际提示确认技能可用。

### 2. 准备原稿和依据

- **原稿**：优先提供 `.docx`，保留自己的原始文件。
- **格式依据**：学校规范、老师要求或 Word 模板，有什么就提供什么；没有也可以明确说明。
- **其他信息**：只有需要封面时才提供真实姓名、学号等信息。

旧 `.doc`、PDF 或扫描件需要先转换、提取并核对内容。只有模板时，能采用模板里可观察的格式，无法由此推断学校未提供的全部标准。

### 3. 发出排版需求并查看交付

把文件交给 Codex，复制下面的提示词，替换文件名即可。收到结果后，查看另存的 Word 文件、差异说明和页面检查情况。

```text
用 $word-format-from-content-template 排版这份课程报告。

原稿：课程报告.docx
格式规范：格式要求.pdf
Word 模板：学校模板.docx
输出：课程报告_排版完成.docx

请按我提供的要求排版，保留原稿，另存结果。
保全原有正文、数值、表格、公式、图片、脚注、批注和链接。
检查标题样式、字体、数字字体、行距、图表题注与正文引用、公式编号及页面空白。
缺失的图表题注和必要引用句，只根据现有内容补中性说明；记录所有新增文字与编号、公式表示变化。
完成内容比较和对象复核，并将最终稿渲染后逐页检查。
交付 Word 文件，同时说明采用的规则、已检查范围、差异和未完成的验收。
```

没有规范或模板时，用这个版本：

```text
用 $word-format-from-content-template 排版课程报告.docx。
这次没有额外规范和模板，请读取技能中的默认排版预设执行。
保留原稿，另存为课程报告_排版完成.docx；保全正文、数据与对象。
按预设补必要的中性题注、图表引用及公式编号，记录差异。
完成内容比较、对象复核和最终稿逐页视觉检查；无法完成的检查请明确说明。
```

## 没有规范时，采用什么格式？

默认的字号、字体、中文样式名、1.5 倍行距、题注、公式编号和图文布局规则，统一维护在 [默认排版预设](word-format-from-content-template/references/default-format.md)。你可以打开它查看或调整自己的偏好，无需在多份说明中重复修改。

同一属性冲突时，采用顺序为：

**本次明确要求 > 适用的明确规范 > 模板 > 默认预设 > 工具默认。**

因此，学校要求的字体、字号等会覆盖默认预设。只覆盖冲突的属性，其余属性仍按适用规则处理。

## 交付与验收

一次完整交付应包含：

- **另存的 DOCX**，保留输入原件。
- **采用的规则与材料**，说明规范、模板和默认预设分别如何使用。
- **内容与对象核对结果**，说明比较范围、授权差异及待复核项目。
- **最终页面检查结果**，将当前最终稿渲染为页面图片，确保所有页面都有视觉审阅依据，检查裁切、遮挡、乱码、编号、图题归属与异常空白。

无法渲染时，应明确标记“未完成视觉验收”。内容比较不能验证字体、行距、分页或最终页码，也不能代替看页面。详细边界见 [比较工具与验收说明](word-format-from-content-template/references/validation.md) 和 [质检清单](word-format-from-content-template/references/quality-checklists.md)。

## 示例与测试

[合成示例](word-format-from-content-template/assets/demo/) 包含 [原稿](word-format-from-content-template/assets/demo/source.docx)、[模板](word-format-from-content-template/assets/demo/template.docx)、[格式要求](word-format-from-content-template/assets/demo/format-rules.md) 和 [排版结果](word-format-from-content-template/assets/demo/formatted.docx)。它用于展示流程，数据为演示构造，未使用用户真实作业。

示例自带明确规范，其中一级标题 **18pt 黑体**等设置遵循示例要求，**不代表默认预设**。这是“明确规范优先”的演示；单页示例的 [验收记录](word-format-from-content-template/assets/demo/verification.json) 只适用于该示例。

当前校验记录包括：

- **63 个 CLI 回归案例**：验证内容比较工具的差异检测、复核提示、输入保护等行为，使用合成测试材料；这不是 63 篇真实文档的排版验证。
- **8 项 PowerShell 前置检查**：检查 Word 导出回退脚本的前置校验行为；这不是启动 Word 并完成导出、渲染和逐页检查的端到端验证。

可在解压后的仓库根目录运行内容比较回归测试：

```text
python -X utf8 word-format-tests/test_compare_text.py
```

需要手动比较原稿与结果时，在仓库根目录运行：

```text
python -X utf8 word-format-from-content-template/scripts/compare_text.py 原稿.docx 排版结果.docx --scope all --show-diff --report comparison.json
```

`python` 应替换为当前环境可用的解释器。退出码 `0` 表示指定范围的提取结果一致，`1` 表示有差异，`2` 表示输入或运行错误，`3` 表示需要复核；还需阅读报告中的范围和警告。

## 依赖与局限

- **使用环境**：需要能安装、调用本技能并访问文档文件的 Codex 环境。
- **内容比较**：`compare_text.py` 及其提取模块仅使用 Python 标准库；支持 UTF-8 文本和常见 Transitional OOXML DOCX。Strict OOXML、加密文件等不能直接按相同方式处理。
- **文档修改与渲染**：由 Codex 探测当前环境的工具。优先使用可用的文档运行时与渲染器；示例生成器另需 `python-docx`。
- **Windows 可选回退**：[render_word.ps1](word-format-from-content-template/scripts/render_word.ps1) 依赖本机 Microsoft Word 及可用的自动化权限，以隐藏实例只读导出 PDF，拒绝覆盖已有 PDF。导出后还需要 PDF 转图片工具和逐页审阅。
- **复杂文档**：自动编号、目录与引用字段、浮动对象、图表、嵌入对象及修订等需要额外复核；不会自行接受或拒绝修订。不同渲染环境也可能改变分页。
- **内容边界**：用于排版和验收，正文润色、论证扩写、补写结论或个人信息需要另外明确授权。无法可靠验证的部分应在交付中说明。

## 仓库内容

| 路径 | 内容 |
| --- | --- |
| [word-format-from-content-template/](word-format-from-content-template/) | 技能入口、执行流程、默认预设、质检说明、脚本和合成示例。 |
| [word-format-tests/](word-format-tests/) | 内容比较工具的合成回归测试。 |
| [排版助手.zip](排版助手.zip) | 便于下载和安装的技能发布包。 |

## 来源与改造

本项目基于 [Tang1206cc/codex-skill-paper-formatting-skill](https://github.com/Tang1206cc/codex-skill-paper-formatting-skill) 的 `word-format-from-content-template` 改造，参考上游提交 [25466da](https://github.com/Tang1206cc/codex-skill-paper-formatting-skill/commit/25466da8e879d1066c4ea9ddfd9b2fb99b875921)，感谢原作者的工作。本仓库围绕默认排版预设、内容与对象保全边界、差异校验、页面验收、回归测试和中文使用说明进行了整理与调整。

上游仓库目前未见 `LICENSE` 文件；本仓库暂未新增许可证。上游内容的作者归属仍属于原作者，相关使用与分发的许可范围需以权利人的明确许可为准。
