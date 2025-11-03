# Markdown 转换器 - Google Docs 插件

这是一个 Google Docs 插件，可以一键将文档中的 Markdown 语法转换为 Google Docs 的原生格式。

## 功能特性

### 支持的 Markdown 语法

#### 1. 标题 (Headers)
```markdown
# 一级标题
## 二级标题
### 三级标题
#### 四级标题
##### 五级标题
###### 六级标题
```

#### 2. 文本格式 (Text Formatting)
```markdown
**加粗文本** 或 __加粗文本__
*斜体文本* 或 _斜体文本_
***粗斜体*** 或 ___粗斜体___
~~删除线文本~~
```

#### 3. 列表 (Lists)
```markdown
- 无序列表项 1
- 无序列表项 2
* 也可以用星号

1. 有序列表项 1
2. 有序列表项 2
3. 有序列表项 3
```

#### 4. 引用块 (Blockquotes)
```markdown
> 这是一段引用文本
> 会显示为缩进的灰色文字
```

#### 5. 表格 (Tables)
```markdown
| 列标题1 | 列标题2 | 列标题3 |
|---------|---------|---------|
| 单元格1 | 单元格2 | 单元格3 |
| 数据A   | 数据B   | 数据C   |
```

#### 6. 代码 (Code)
```markdown
这是`行内代码`示例

```
这是代码块
可以多行
支持等宽字体显示
```
```

---

## 安装步骤

### 方法 1: 直接在 Google Docs 中创建

1. **打开 Google Docs**
   - 创建一个新文档或打开现有文档

2. **打开 Apps Script 编辑器**
   - 点击顶部菜单: `扩展程序` → `Apps Script`

3. **删除默认代码**
   - 删除编辑器中的默认 `function myFunction() {}` 代码

4. **粘贴插件代码**
   - 复制 `Code.gs` 文件中的所有代码
   - 粘贴到 Apps Script 编辑器中

5. **保存项目**
   - 点击保存图标（或按 Ctrl+S / Cmd+S）
   - 给项目命名为 "Markdown Converter"

6. **授权插件**
   - 点击运行按钮（▶️）或者返回文档刷新页面
   - 第一次运行会要求授权，点击"审核权限"
   - 选择你的 Google 账号
   - 点击"高级" → "前往 Markdown Converter（不安全）"
   - 点击"允许"

7. **完成！**
   - 返回 Google Docs，刷新页面
   - 你会在菜单栏看到新的 "Markdown 转换器" 菜单

---

## 使用方法

### 基本使用流程

1. **在 Google Docs 中输入 Markdown 格式的文本**

   例如：
   ```
   # 我的文档标题

   这是一段**加粗文本**和*斜体文本*。

   ## 功能列表

   - 功能 1
   - 功能 2
   - 功能 3

   ## 数据表格

   | 姓名 | 年龄 | 城市 |
   |------|------|------|
   | 张三 | 25   | 北京 |
   | 李四 | 30   | 上海 |
   ```

2. **点击转换**
   - 点击菜单栏的 `Markdown 转换器` → `转换整个文档`

3. **等待转换完成**
   - 转换过程中会自动处理所有 Markdown 语法
   - 完成后会弹出提示 "✅ Markdown 转换完成！"

4. **查看结果**
   - 所有 Markdown 语法已转换为 Google Docs 原生格式
   - 标题会显示为文档标题样式
   - 表格会显示为真实的表格
   - 列表会显示为项目符号列表

---

## 使用示例

### 示例 1: 文档大纲

**转换前（Markdown）:**
```markdown
# 项目计划文档

## 1. 项目概述

这是一个**重要的项目**，需要在*三个月内*完成。

## 2. 主要功能

- 用户管理
- 数据分析
- 报表生成

## 3. 时间表

| 阶段 | 时间 | 负责人 |
|------|------|--------|
| 需求分析 | 2周 | 张三 |
| 开发 | 6周 | 李四 |
| 测试 | 2周 | 王五 |
```

**转换后:**
- "项目计划文档" 显示为大标题（Heading 1）
- "1. 项目概述" 显示为二级标题（Heading 2）
- "重要的项目" 显示为加粗
- "三个月内" 显示为斜体
- 主要功能显示为项目符号列表
- 时间表显示为格式化的表格

---

### 示例 2: 技术文档

**转换前（Markdown）:**
```markdown
# API 文档

## 用户登录接口

调用 `POST /api/login` 接口进行登录。

**请求参数:**

| 参数名 | 类型 | 必填 | 说明 |
|--------|------|------|------|
| username | string | 是 | 用户名 |
| password | string | 是 | 密码 |

**示例代码:**

```
fetch('/api/login', {
  method: 'POST',
  body: JSON.stringify({username, password})
})
```
```

**转换后:**
- "API 文档" 和 "用户登录接口" 显示为标题
- `POST /api/login` 显示为行内代码格式（等宽字体）
- 请求参数表格正确渲染
- 示例代码显示为代码块（等宽字体 + 灰色背景）

---

## 注意事项

### ⚠️ 重要提示

1. **转换是不可逆的**
   - 转换后，Markdown 标记（如 `**`、`#` 等）会被移除
   - 建议在转换前保存一份原始文档备份

2. **转换顺序**
   - 插件会按特定顺序处理（代码块 → 标题 → 表格 → 列表 → 文本格式）
   - 这样可以避免嵌套格式的冲突

3. **代码块不会被转换**
   - 代码块内的 Markdown 语法不会被转换
   - 例如代码示例中的 `**` 会保持原样

4. **表格格式要求**
   - 表格每行必须以 `|` 开头和结尾
   - 分隔行（`|-----|`）会被自动跳过
   - 第一行会自动设置为表头（加粗 + 灰色背景）

5. **性能考虑**
   - 对于大型文档（超过 50 页），转换可能需要几秒钟
   - 转换过程中请耐心等待，不要关闭文档

---

## 技术说明

### 转换逻辑

插件使用以下方式进行转换：

1. **文本查找和替换**
   - 使用正则表达式查找 Markdown 模式
   - 例如: `\*\*(.+?)\*\*` 匹配 `**文本**`

2. **格式应用**
   - 移除 Markdown 标记
   - 应用 Google Docs 原生格式 API
   - 例如: `text.setBold(start, end, true)`

3. **元素创建**
   - 对于表格，创建真实的 Google Docs 表格对象
   - 对于列表，使用 `setGlyphType()` 设置项目符号

### 代码结构

```
Code.gs
├── onOpen()                    // 创建菜单
├── convertMarkdownToFormat()   // 主转换函数
├── convertHeaders()            // 标题转换
├── convertBoldItalic()         // 加粗和斜体
├── convertStrikethrough()      // 删除线
├── convertLists()              // 列表
├── convertBlockquotes()        // 引用块
├── convertTables()             // 表格
├── convertCodeBlocks()         // 代码块
└── convertInlineCode()         // 行内代码
```

---

## 常见问题 (FAQ)

### Q1: 出现 "para.setGlyphType is not a function" 错误怎么办？

**A:** 这个错误已在最新版本中修复：
- 原因：旧版 API `setGlyphType()` 已被弃用
- 解决方案：使用 `setAttributes()` 方法替代
- 如果还遇到此错误，请确保使用最新版本的 Code.gs 文件

### Q2: 为什么转换后某些格式没有变化？

**A:** 可能的原因：
- Markdown 语法不正确（例如 `**加粗**` 必须是成对的）
- 格式嵌套过于复杂
- 建议检查原始 Markdown 语法是否符合标准

### Q3: 可以撤销转换吗？

**A:**
- 可以使用 Google Docs 的撤销功能（Ctrl+Z / Cmd+Z）
- 或者在转换前创建文档副本作为备份

### Q4: 转换失败怎么办？

**A:**
1. 查看浏览器控制台是否有错误信息
2. 打开 Apps Script 编辑器，查看日志（`查看` → `日志`）
3. 尝试分段转换（先转换一部分内容）

### Q5: 支持嵌套列表吗？

**A:**
- 当前版本暂不支持多级嵌套列表
- 建议将嵌套列表拆分为多个单级列表

### Q6: 能否自定义转换样式？

**A:**
- 可以！打开 `Code.gs` 修改代码
- 例如修改引用块的缩进量: `para.setIndentStart(36)` 改为其他数值
- 例如修改代码块的背景颜色: `codePara.setBackgroundColor('#f5f5f5')`

---

## 自定义和扩展

### 修改样式

如果你想修改转换后的样式，可以编辑 `Code.gs` 中的相关函数：

**示例 1: 修改引用块颜色**
```javascript
// 在 convertBlockquotes() 函数中
textElement.setForegroundColor('#666666');  // 改为你想要的颜色
```

**示例 2: 修改代码块背景**
```javascript
// 在 convertCodeBlocks() 函数中
codePara.setBackgroundColor('#f5f5f5');  // 改为你想要的颜色
```

**示例 3: 修改表头样式**
```javascript
// 在 convertTables() 函数中
if (row === 0) {
  cell.editAsText().setBold(true);
  cell.setBackgroundColor('#f0f0f0');  // 改为你想要的颜色
}
```

### 添加新功能

你可以添加新的转换函数来支持更多 Markdown 语法：

```javascript
// 例如：支持高亮文本 ==高亮==
function convertHighlight(body) {
  var found = body.findText('==(.+?)==');
  while (found) {
    // ... 转换逻辑
    text.setBackgroundColor(startOffset, endOffset, '#ffff00');
  }
}

// 然后在 convertMarkdownToFormat() 中调用
convertHighlight(body);
```

---

## 版本历史

### v1.0.0 (2025-01-01)
- ✅ 支持标题（1-6 级）
- ✅ 支持加粗、斜体、删除线
- ✅ 支持有序和无序列表
- ✅ 支持引用块
- ✅ 支持表格
- ✅ 支持代码块和行内代码

---

## 许可证

本项目为开源项目，可自由使用和修改。

---

## 联系方式

如有问题或建议，欢迎反馈！

---

## 致谢

感谢使用 Markdown 转换器插件！希望它能提高你的文档编辑效率。
