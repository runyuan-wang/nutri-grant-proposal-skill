# 申报书 docx 格式处理技术（python-docx）

## 模板结构认知

此类申报书通常"全文装在表格里"：body 下少量正文段落 + 十余张 1×1 大表格（每节一张）+ 信息/成员/预算等真实表格 + 嵌套表（进度安排）。遍历时必须同时处理 `d.paragraphs` 与 `d.tables`（含 `cell.tables` 嵌套）。

## 字体统一（推荐做法）

优先改**样式层**而非逐 run：

1. `styles.xml` 的 `docDefaults/rPrDefault/rPr` 设 `w:rFonts`：`ascii="Times New Roman"`、`hAnsi="Times New Roman"`、`eastAsia="宋体"`。
2. Normal 样式字号设 21 half-points（10.5pt = 五号）；图注、表内备注可用 19（9.5pt = 小五）。
3. 标题段显式 `eastAsia="黑体"`。
4. 逐 run 清理只在样式层管不到的残留（显式 rFonts 覆盖）时做；审计口径：统计所有 `w:r` 的 `(eastAsia, ascii)` 与 `sz` 分布，目标收敛为 `(None,None)` 继承 + 少量黑体标题。

字号对照（half-points）：五号 21、小五 18、小四 24、四号 28、三号 32、小二 36、二号 44。

## run 级跨 run 替换

Word 常把一句话切成多个 run（格式边界/拼写检查），直接 `run.text.replace` 会漏。做法：

```python
def repl(p, old, new):
    """段落内跨 run 替换：合并到首 run，其余清空。"""
    full = "".join(r.text for r in p.runs)
    if old not in full:
        return False
    p.runs[0].text = full.replace(old, new)
    for r in p.runs[1:]:
        r.text = ""
    return True
```

注意：合并后继承首 run 格式；替换前确认该段格式均一。

## 单元格与表格坑

- `tcPr` 子元素有顺序约束：`w:shd` 必须在 `w:vAlign` **之前**，否则 Word 打开报错。
- 嵌套表设 `w:tblLayout type="fixed"` 与固定 `tblW`；跨页长表设 `w:tblHeader` 让表头跨页重复。
- 合并单元格在 python-docx 中重复出现，遍历 cells 会去重失真，金额校验按"行"取唯一值。

## 技术路线图

- 竖版：3～4 个阶段框纵向排列 + 向下箭头 + 侧边注释文本框。
- 用 Word 原生图形（DrawingML `wpg:wsp`/`prst="rect"`、`prst="line"` + `tailEnd` 箭头、`wpg:txbx` 文本框），直角、黑线、白底、宋体；忌彩色填充与圆角。
- 也可直接插图片（用户最终接受图片版）；图注编号"图1/图2"全文引用一致。

## 版面自检

- 页面 A4 竖装，正文不小于 5 号。
- 估算版面：CJK 每字 1em、拉丁 0.52em，行高≈字号×1.25，版心宽约 416～438pt；表格加图后按行高累加预估是否溢出。
- macOS 无 Word 自动化权限时（-10004），用 Chrome headless 渲染 HTML 预览做几何自检。

## 备份纪律

每次大改前存链式备份（`.bak`、`.bak2_合并前`…），出错可逐级回滚。
