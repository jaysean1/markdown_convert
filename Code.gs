// location: markdown_convert/Code.gs
// Purpose: Google Docs Add-on to convert Markdown syntax to Google Docs formatting
// This file contains all the conversion functions for headers, bold, italic, lists, tables, code, etc.

/**
 * 当文档打开时运行，在菜单栏添加自定义菜单
 * This function runs when the document is opened and adds a custom menu
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Markdown 转换器')
    .addItem('转换整个文档', 'convertMarkdownToFormat')
    .addToUi();
}

/**
 * 主转换函数 - 按顺序调用各个转换模块
 * Main conversion function - calls each conversion module in order
 *
 * 转换顺序很重要:
 * 1. 先处理代码块（避免代码中的 Markdown 被转换）
 * 2. 再处理标题
 * 3. 然后处理表格
 * 4. 再处理列表和引用
 * 5. 最后处理行内格式（加粗、斜体等）
 */
function convertMarkdownToFormat() {
  try {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();

    // 按顺序执行转换
    convertCodeBlocks(body);     // 代码块（最先处理）
    convertHeaders(body);         // 标题
    convertTables(body);          // 表格
    convertLists(body);           // 列表
    convertBlockquotes(body);     // 引用块
    convertBoldItalic(body);      // 加粗和斜体
    convertStrikethrough(body);   // 删除线
    convertInlineCode(body);      // 行内代码（最后处理）

    DocumentApp.getUi().alert('✅ Markdown 转换完成！');
  } catch (e) {
    DocumentApp.getUi().alert('❌ 转换失败: ' + e.message);
    Logger.log('Error: ' + e.toString());
  }
}

/**
 * 转换标题 - 将 # 转换为 Google Docs 标题样式
 * Convert headers - transform # to Google Docs heading styles
 *
 * 支持:
 * # 一级标题 -> Heading 1
 * ## 二级标题 -> Heading 2
 * ### 三级标题 -> Heading 3
 * 等等...
 */
function convertHeaders(body) {
  var paragraphs = body.getParagraphs();

  for (var i = 0; i < paragraphs.length; i++) {
    var para = paragraphs[i];
    var text = para.getText();

    // 匹配行首的 # 符号（1-6个）
    // Match # symbols at the start of the line (1-6)
    var headerMatch = text.match(/^(#{1,6})\s+(.+)$/);

    if (headerMatch) {
      var level = headerMatch[1].length;  // # 的数量
      var cleanText = headerMatch[2];     // 移除 # 后的文本

      // 更新段落文本
      para.setText(cleanText);

      // 根据 # 数量设置对应的标题级别
      // Set heading level based on number of #
      switch(level) {
        case 1:
          para.setHeading(DocumentApp.ParagraphHeading.HEADING1);
          break;
        case 2:
          para.setHeading(DocumentApp.ParagraphHeading.HEADING2);
          break;
        case 3:
          para.setHeading(DocumentApp.ParagraphHeading.HEADING3);
          break;
        case 4:
          para.setHeading(DocumentApp.ParagraphHeading.HEADING4);
          break;
        case 5:
          para.setHeading(DocumentApp.ParagraphHeading.HEADING5);
          break;
        case 6:
          para.setHeading(DocumentApp.ParagraphHeading.HEADING6);
          break;
      }

      Logger.log('Converted heading: ' + cleanText);
    }
  }
}

/**
 * 转换加粗和斜体
 * Convert bold and italic text
 *
 * 支持:
 * **粗体** 或 __粗体__ -> 加粗
 * *斜体* 或 _斜体_ -> 斜体
 * ***粗斜体*** -> 加粗+斜体
 */
function convertBoldItalic(body) {
  // 先处理粗斜体 ***text***（避免被单独的加粗或斜体匹配）
  // Process bold+italic first to avoid separate matches
  processBoldItalic(body, '\\*\\*\\*(.+?)\\*\\*\\*', true, true);
  processBoldItalic(body, '___(.+?)___', true, true);

  // 处理加粗 **text** 或 __text__
  // Process bold
  processBoldItalic(body, '\\*\\*(.+?)\\*\\*', true, false);
  processBoldItalic(body, '__(.+?)__', true, false);

  // 处理斜体 *text* 或 _text_
  // Process italic
  processBoldItalic(body, '\\*(.+?)\\*', false, true);
  processBoldItalic(body, '_(.+?)_', false, true);
}

/**
 * 辅助函数：处理加粗和斜体的实际转换
 * Helper function: performs the actual bold/italic conversion
 */
function processBoldItalic(body, pattern, isBold, isItalic) {
  var found = body.findText(pattern);

  while (found) {
    var element = found.getElement();
    var startOffset = found.getStartOffset();
    var endOffset = found.getEndOffsetInclusive();

    var text = element.asText();
    var matchText = text.getText().substring(startOffset, endOffset + 1);

    // 提取标记符号的长度（* 或 _ 的数量）
    var markerLength = isBold && isItalic ? 3 : (isBold || isItalic ? 2 : 1);
    if (pattern.indexOf('_') !== -1 && pattern.indexOf('\\*') === -1) {
      // 如果是下划线模式
      markerLength = matchText.match(/^_+/)[0].length;
    } else {
      // 如果是星号模式
      markerLength = matchText.match(/^\*+/)[0].length;
    }

    // 提取内容（去掉标记符号）
    var content = matchText.substring(markerLength, matchText.length - markerLength);

    // 删除原文本，插入新文本
    text.deleteText(startOffset, endOffset);
    text.insertText(startOffset, content);

    // 应用格式
    var contentEnd = startOffset + content.length - 1;
    if (isBold) {
      text.setBold(startOffset, contentEnd, true);
    }
    if (isItalic) {
      text.setItalic(startOffset, contentEnd, true);
    }

    Logger.log('Applied format to: ' + content);

    // 继续查找下一个匹配
    found = body.findText(pattern);
  }
}

/**
 * 转换删除线
 * Convert strikethrough text
 *
 * 支持: ~~删除线~~ -> 删除线
 */
function convertStrikethrough(body) {
  var found = body.findText('~~(.+?)~~');

  while (found) {
    var element = found.getElement();
    var startOffset = found.getStartOffset();
    var endOffset = found.getEndOffsetInclusive();

    var text = element.asText();
    var matchText = text.getText().substring(startOffset, endOffset + 1);

    // 提取内容（去掉 ~~）
    var content = matchText.substring(2, matchText.length - 2);

    // 删除原文本，插入新文本
    text.deleteText(startOffset, endOffset);
    text.insertText(startOffset, content);

    // 应用删除线格式
    text.setStrikethrough(startOffset, startOffset + content.length - 1, true);

    Logger.log('Applied strikethrough to: ' + content);

    found = body.findText('~~(.+?)~~');
  }
}

/**
 * 转换列表
 * Convert lists (ordered and unordered)
 *
 * 支持:
 * - 无序列表 -> 项目符号列表
 * * 无序列表 -> 项目符号列表
 * 1. 有序列表 -> 编号列表
 */
function convertLists(body) {
  var paragraphs = body.getParagraphs();

  for (var i = 0; i < paragraphs.length; i++) {
    var para = paragraphs[i];
    var text = para.getText();

    // 无序列表: - 或 *
    var unorderedMatch = text.match(/^[\-\*]\s+(.+)$/);
    if (unorderedMatch) {
      var content = unorderedMatch[1];
      para.setText(content);

      // 使用 setAttributes 方法设置列表样式（正确的 API）
      // Use setAttributes method to set list style (correct API)
      var attributes = {};
      attributes[DocumentApp.Attribute.GLYPH_TYPE] = DocumentApp.GlyphType.BULLET;
      para.setAttributes(attributes);

      Logger.log('Converted unordered list: ' + content);
      continue;
    }

    // 有序列表: 1. 2. 3. 等
    var orderedMatch = text.match(/^\d+\.\s+(.+)$/);
    if (orderedMatch) {
      var content = orderedMatch[1];
      para.setText(content);

      // 使用 setAttributes 方法设置列表样式（正确的 API）
      // Use setAttributes method to set list style (correct API)
      var attributes = {};
      attributes[DocumentApp.Attribute.GLYPH_TYPE] = DocumentApp.GlyphType.NUMBER;
      para.setAttributes(attributes);

      Logger.log('Converted ordered list: ' + content);
    }
  }
}

/**
 * 转换引用块
 * Convert blockquotes
 *
 * 支持: > 引用文本 -> 缩进格式
 */
function convertBlockquotes(body) {
  var paragraphs = body.getParagraphs();

  for (var i = 0; i < paragraphs.length; i++) {
    var para = paragraphs[i];
    var text = para.getText();

    // 匹配 > 开头的引用
    var quoteMatch = text.match(/^>\s+(.+)$/);

    if (quoteMatch) {
      var content = quoteMatch[1];
      para.setText(content);

      // 设置左缩进和灰色字体来模拟引用样式
      para.setIndentStart(36);  // 缩进 36 磅
      para.setIndentFirstLine(36);

      // 设置字体颜色为灰色
      var textElement = para.editAsText();
      textElement.setForegroundColor('#666666');

      Logger.log('Converted blockquote: ' + content);
    }
  }
}

/**
 * 转换表格
 * Convert Markdown tables to Google Docs tables
 *
 * 支持:
 * | 列1 | 列2 |
 * |-----|-----|
 * | 值1 | 值2 |
 */
function convertTables(body) {
  var paragraphs = body.getParagraphs();
  var i = 0;

  while (i < paragraphs.length) {
    var para = paragraphs[i];
    var text = para.getText();

    // 检测表格行: | col1 | col2 |
    if (text.match(/^\|(.+)\|$/)) {
      var tableRows = [];
      var startIndex = i;
      var tableParagraphs = [];

      // 收集连续的表格行
      while (i < paragraphs.length) {
        var currentPara = paragraphs[i];
        var currentText = currentPara.getText();

        if (!currentText.match(/^\|(.+)\|$/)) {
          break;  // 不再是表格行，退出
        }

        tableParagraphs.push(currentPara);

        // 分割单元格
        var cells = currentText
          .substring(1, currentText.length - 1)  // 去掉首尾的 |
          .split('|')
          .map(function(cell) { return cell.trim(); });

        // 跳过分隔行 (|-----|-----|)
        if (!cells[0].match(/^-+$/)) {
          tableRows.push(cells);
        }

        i++;
      }

      // 如果收集到了表格数据，创建 Google Docs 表格
      if (tableRows.length > 0) {
        // 在第一个表格段落位置插入表格
        var childIndex = body.getChildIndex(tableParagraphs[0]);
        var table = body.insertTable(childIndex);

        // 填充表格数据
        for (var row = 0; row < tableRows.length; row++) {
          var tableRow;
          if (row === 0) {
            tableRow = table.appendTableRow();
          } else {
            tableRow = table.appendTableRow();
          }

          for (var col = 0; col < tableRows[row].length; col++) {
            var cell = tableRow.appendTableCell(tableRows[row][col]);

            // 第一行设置为表头样式（加粗）
            if (row === 0) {
              cell.editAsText().setBold(true);
              cell.setBackgroundColor('#f0f0f0');
            }
          }
        }

        // 删除原始 Markdown 表格文本
        for (var j = tableParagraphs.length - 1; j >= 0; j--) {
          body.removeChild(tableParagraphs[j]);
        }

        Logger.log('Converted table with ' + tableRows.length + ' rows');
      }

      continue;  // 继续处理下一段
    }

    i++;
  }
}

/**
 * 转换代码块
 * Convert code blocks
 *
 * 支持: ```代码``` -> 等宽字体 + 灰色背景
 */
function convertCodeBlocks(body) {
  var paragraphs = body.getParagraphs();
  var i = 0;

  while (i < paragraphs.length) {
    var para = paragraphs[i];
    var text = para.getText();

    // 检测代码块开始: ```
    if (text.match(/^```/)) {
      var codeLines = [];
      var codeParagraphs = [];
      var startIndex = i;

      i++;  // 跳过开始的 ```

      // 收集代码内容直到遇到结束的 ```
      while (i < paragraphs.length) {
        var currentPara = paragraphs[i];
        var currentText = currentPara.getText();

        if (currentText.match(/^```$/)) {
          // 遇到结束标记
          codeParagraphs.push(currentPara);
          i++;
          break;
        }

        codeLines.push(currentText);
        codeParagraphs.push(currentPara);
        i++;
      }

      if (codeLines.length > 0) {
        // 删除开始的 ``` 段落
        body.removeChild(paragraphs[startIndex]);

        // 为代码行应用格式
        for (var j = 0; j < codeParagraphs.length - 1; j++) {
          var codePara = codeParagraphs[j];
          if (codePara.getParent()) {
            var codeText = codePara.editAsText();
            codeText.setFontFamily('Courier New');  // 等宽字体
            codeText.setForegroundColor('#333333');
            codePara.setBackgroundColor('#f5f5f5');  // 浅灰背景
          }
        }

        // 删除结束的 ``` 段落
        if (codeParagraphs.length > 0) {
          var lastPara = codeParagraphs[codeParagraphs.length - 1];
          if (lastPara.getParent()) {
            body.removeChild(lastPara);
          }
        }

        Logger.log('Converted code block with ' + codeLines.length + ' lines');
      }

      continue;
    }

    i++;
  }
}

/**
 * 转换行内代码
 * Convert inline code
 *
 * 支持: `代码` -> 等宽字体
 */
function convertInlineCode(body) {
  var found = body.findText('`(.+?)`');

  while (found) {
    var element = found.getElement();
    var startOffset = found.getStartOffset();
    var endOffset = found.getEndOffsetInclusive();

    var text = element.asText();
    var matchText = text.getText().substring(startOffset, endOffset + 1);

    // 提取内容（去掉 `）
    var content = matchText.substring(1, matchText.length - 1);

    // 删除原文本，插入新文本
    text.deleteText(startOffset, endOffset);
    text.insertText(startOffset, content);

    // 应用代码格式
    var contentEnd = startOffset + content.length - 1;
    text.setFontFamily(startOffset, contentEnd, 'Courier New');
    text.setForegroundColor(startOffset, contentEnd, '#c7254e');
    text.setBackgroundColor(startOffset, contentEnd, '#f9f2f4');

    Logger.log('Applied inline code format to: ' + content);

    found = body.findText('`(.+?)`');
  }
}
