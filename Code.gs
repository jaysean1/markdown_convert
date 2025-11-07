// location: markdown_convert/Code.gs
// Purpose: Google Docs Add-on to convert Markdown syntax to Google Docs formatting
// This file contains all the conversion functions for headers, bold, italic, lists, tables, horizontal rules, code, etc.

/**
 * Runs when the document is opened and adds a custom menu to the menu bar
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Markdown Converter')
    .addItem('Convert Entire Document', 'convertMarkdownToFormat')
    .addToUi();
}

/**
 * Main conversion function - calls each conversion module in order
 *
 * Conversion order is important:
 * 1. Process code blocks first (prevent Markdown inside code from being converted)
 * 2. Then process headers
 * 3. Then process tables
 * 4. Then process horizontal rules (after tables to avoid converting table separators)
 * 5. Then process lists and blockquotes
 * 6. Finally process inline formatting (bold, italic, etc.)
 */
function convertMarkdownToFormat() {
  try {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();

    // Execute conversions in order
    convertCodeBlocks(body);     // Code blocks (process first)
    convertHeaders(body);         // Headers
    convertTables(body);          // Tables
    convertHorizontalRules(body); // Horizontal rules (after tables, before lists)
    convertLists(body);           // Lists
    convertBlockquotes(body);     // Blockquotes
    convertBoldItalic(body);      // Bold and italic
    convertStrikethrough(body);   // Strikethrough
    convertInlineCode(body);      // Inline code (process last)

    DocumentApp.getUi().alert('✅ Markdown conversion complete!');
  } catch (e) {
    DocumentApp.getUi().alert('❌ Conversion failed: ' + e.message);
    Logger.log('Error: ' + e.toString());
  }
}

/**
 * Convert headers - transform # to Google Docs heading styles
 *
 * Supports:
 * # Heading 1 -> Heading 1
 * ## Heading 2 -> Heading 2
 * ### Heading 3 -> Heading 3
 * etc...
 */
function convertHeaders(body) {
  var paragraphs = body.getParagraphs();

  for (var i = 0; i < paragraphs.length; i++) {
    var para = paragraphs[i];
    var text = para.getText();

    // Match # symbols at the start of the line (1-6)
    var headerMatch = text.match(/^(#{1,6})\s+(.+)$/);

    if (headerMatch) {
      var level = headerMatch[1].length;  // Number of # symbols
      var cleanText = headerMatch[2];     // Text after removing #

      // Update paragraph text
      para.setText(cleanText);

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
 * Convert bold and italic text
 *
 * Supports:
 * **Bold** or __Bold__ -> Bold
 * *Italic* or _Italic_ -> Italic
 * ***Bold+Italic*** -> Bold+Italic
 */
function convertBoldItalic(body) {
  // Process bold+italic first to avoid separate matches
  processBoldItalic(body, '\\*\\*\\*([^\\*]+)\\*\\*\\*', true, true);
  processBoldItalic(body, '___([^_]+)___', true, true);

  // Process bold **text** or __text__
  processBoldItalic(body, '\\*\\*([^\\*]+)\\*\\*', true, false);
  processBoldItalic(body, '__([^_]+)__', true, false);

  // Process italic *text* or _text_
  processBoldItalic(body, '\\*([^\\*]+)\\*', false, true);
  processBoldItalic(body, '_([^_]+)_', false, true);
}

/**
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

    // Extract marker symbol length (number of * or _)
    var markerLength = isBold && isItalic ? 3 : (isBold || isItalic ? 2 : 1);
    if (pattern.indexOf('_') !== -1 && pattern.indexOf('\\*') === -1) {
      // If underscore pattern
      markerLength = matchText.match(/^_+/)[0].length;
    } else {
      // If asterisk pattern
      markerLength = matchText.match(/^\*+/)[0].length;
    }

    // Extract content (remove marker symbols)
    var content = matchText.substring(markerLength, matchText.length - markerLength);

    // Delete original text and insert new text
    text.deleteText(startOffset, endOffset);
    text.insertText(startOffset, content);

    // Apply formatting
    var contentEnd = startOffset + content.length - 1;
    if (isBold) {
      text.setBold(startOffset, contentEnd, true);
    }
    if (isItalic) {
      text.setItalic(startOffset, contentEnd, true);
    }

    Logger.log('Applied format to: ' + content);

    // Continue searching for next match
    found = body.findText(pattern);
  }
}

/**
 * Convert strikethrough text
 *
 * Supports: ~~strikethrough~~ -> strikethrough
 */
function convertStrikethrough(body) {
  var found = body.findText('~~(.+?)~~');

  while (found) {
    var element = found.getElement();
    var startOffset = found.getStartOffset();
    var endOffset = found.getEndOffsetInclusive();

    var text = element.asText();
    var matchText = text.getText().substring(startOffset, endOffset + 1);

    // Extract content (remove ~~)
    var content = matchText.substring(2, matchText.length - 2);

    // Delete original text and insert new text
    text.deleteText(startOffset, endOffset);
    text.insertText(startOffset, content);

    // Apply strikethrough format
    text.setStrikethrough(startOffset, startOffset + content.length - 1, true);

    Logger.log('Applied strikethrough to: ' + content);

    found = body.findText('~~(.+?)~~');
  }
}

/**
 * Convert lists (ordered and unordered)
 *
 * Supports:
 * - Unordered list -> Bullet list
 * * Unordered list -> Bullet list
 * 1. Ordered list -> Numbered list
 */
function convertLists(body) {
  var paragraphs = body.getParagraphs();

  for (var i = 0; i < paragraphs.length; i++) {
    var para = paragraphs[i];
    var text = para.getText();

    // Unordered list: - or *
    var unorderedMatch = text.match(/^[\-\*]\s+(.+)$/);
    if (unorderedMatch) {
      var content = unorderedMatch[1];
      para.setText(content);

      // Use setAttributes method to set list style (correct API)
      var attributes = {};
      attributes[DocumentApp.Attribute.GLYPH_TYPE] = DocumentApp.GlyphType.BULLET;
      para.setAttributes(attributes);

      Logger.log('Converted unordered list: ' + content);
      continue;
    }

    // Ordered list: 1. 2. 3. etc.
    var orderedMatch = text.match(/^\d+\.\s+(.+)$/);
    if (orderedMatch) {
      var content = orderedMatch[1];
      para.setText(content);

      // Use setAttributes method to set list style (correct API)
      var attributes = {};
      attributes[DocumentApp.Attribute.GLYPH_TYPE] = DocumentApp.GlyphType.NUMBER;
      para.setAttributes(attributes);

      Logger.log('Converted ordered list: ' + content);
    }
  }
}

/**
 * Convert blockquotes
 *
 * Supports: > quoted text -> indented format
 */
function convertBlockquotes(body) {
  var paragraphs = body.getParagraphs();

  for (var i = 0; i < paragraphs.length; i++) {
    var para = paragraphs[i];
    var text = para.getText();

    // Match quotes starting with >
    var quoteMatch = text.match(/^>\s+(.+)$/);

    if (quoteMatch) {
      var content = quoteMatch[1];
      para.setText(content);

      // Set left indent and gray font to simulate quote style
      para.setIndentStart(36);  // Indent 36 points
      para.setIndentFirstLine(36);

      // Set font color to gray
      var textElement = para.editAsText();
      textElement.setForegroundColor('#666666');

      Logger.log('Converted blockquote: ' + content);
    }
  }
}

/**
 * Convert horizontal rules (dividers)
 *
 * Supports: --- -> Google Docs horizontal line
 * Must be on its own line (spaces before/after are allowed)
 */
function convertHorizontalRules(body) {
  var paragraphs = body.getParagraphs();

  // Process in reverse order to avoid index issues when removing paragraphs
  for (var i = paragraphs.length - 1; i >= 0; i--) {
    var para = paragraphs[i];
    var text = para.getText();

    // Match --- on its own line (allow spaces before/after)
    if (text.match(/^\s*---\s*$/)) {
      // Get the position of this paragraph in the body
      var childIndex = body.getChildIndex(para);

      // Insert horizontal rule at this position
      body.insertHorizontalRule(childIndex);

      // Remove the original paragraph with ---
      body.removeChild(para);

      Logger.log('Converted horizontal rule at position: ' + childIndex);
    }
  }
}

/**
 * Convert Markdown tables to Google Docs tables
 *
 * Supports:
 * | Column 1 | Column 2 |
 * |----------|----------|
 * | Value 1  | Value 2  |
 */
function convertTables(body) {
  var paragraphs = body.getParagraphs();
  var i = 0;

  while (i < paragraphs.length) {
    var para = paragraphs[i];
    var text = para.getText();

    // Detect table row: | col1 | col2 |
    if (text.match(/^\|(.+)\|$/)) {
      var tableRows = [];
      var startIndex = i;
      var tableParagraphs = [];

      // Collect consecutive table rows
      while (i < paragraphs.length) {
        var currentPara = paragraphs[i];
        var currentText = currentPara.getText();

        if (!currentText.match(/^\|(.+)\|$/)) {
          break;  // Not a table row anymore, exit
        }

        tableParagraphs.push(currentPara);

        // Split cells
        var cells = currentText
          .substring(1, currentText.length - 1)  // Remove leading and trailing |
          .split('|')
          .map(function(cell) { return cell.trim(); });

        // Skip separator row (|-----|-----|)
        if (!cells[0].match(/^-+$/)) {
          tableRows.push(cells);
        }

        i++;
      }

      // If table data was collected, create Google Docs table
      if (tableRows.length > 0) {
        // Insert table at the first table paragraph position
        var childIndex = body.getChildIndex(tableParagraphs[0]);
        var table = body.insertTable(childIndex);

        // Fill table data
        for (var row = 0; row < tableRows.length; row++) {
          var tableRow;
          if (row === 0) {
            tableRow = table.appendTableRow();
          } else {
            tableRow = table.appendTableRow();
          }

          for (var col = 0; col < tableRows[row].length; col++) {
            var cell = tableRow.appendTableCell(tableRows[row][col]);

            // Set first row as header style (bold)
            if (row === 0) {
              cell.editAsText().setBold(true);
              cell.setBackgroundColor('#f0f0f0');
            }
          }
        }

        // Delete original Markdown table text
        for (var j = tableParagraphs.length - 1; j >= 0; j--) {
          body.removeChild(tableParagraphs[j]);
        }

        Logger.log('Converted table with ' + tableRows.length + ' rows');
      }

      continue;  // Continue processing next paragraph
    }

    i++;
  }
}

/**
 * Convert code blocks
 *
 * Supports: ```code``` -> monospace font + gray background
 */
function convertCodeBlocks(body) {
  var paragraphs = body.getParagraphs();
  var i = 0;

  while (i < paragraphs.length) {
    var para = paragraphs[i];
    var text = para.getText();

    // Detect code block start: ```
    if (text.match(/^```/)) {
      var codeLines = [];
      var codeParagraphs = [];
      var startIndex = i;

      i++;  // Skip opening ```

      // Collect code content until closing ```
      while (i < paragraphs.length) {
        var currentPara = paragraphs[i];
        var currentText = currentPara.getText();

        if (currentText.match(/^```$/)) {
          // Found closing marker
          codeParagraphs.push(currentPara);
          i++;
          break;
        }

        codeLines.push(currentText);
        codeParagraphs.push(currentPara);
        i++;
      }

      if (codeLines.length > 0) {
        // Delete opening ``` paragraph
        body.removeChild(paragraphs[startIndex]);

        // Apply formatting to code lines
        for (var j = 0; j < codeParagraphs.length - 1; j++) {
          var codePara = codeParagraphs[j];
          if (codePara.getParent()) {
            var codeText = codePara.editAsText();
            codeText.setFontFamily('Courier New');  // Monospace font
            codeText.setForegroundColor('#333333');
            codePara.setBackgroundColor('#f5f5f5');  // Light gray background
          }
        }

        // Delete closing ``` paragraph
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
 * Convert inline code
 *
 * Supports: `code` -> monospace font
 */
function convertInlineCode(body) {
  var found = body.findText('`(.+?)`');

  while (found) {
    var element = found.getElement();
    var startOffset = found.getStartOffset();
    var endOffset = found.getEndOffsetInclusive();

    var text = element.asText();
    var matchText = text.getText().substring(startOffset, endOffset + 1);

    // Extract content (remove `)
    var content = matchText.substring(1, matchText.length - 1);

    // Delete original text and insert new text
    text.deleteText(startOffset, endOffset);
    text.insertText(startOffset, content);

    // Apply code formatting
    var contentEnd = startOffset + content.length - 1;
    text.setFontFamily(startOffset, contentEnd, 'Courier New');
    text.setForegroundColor(startOffset, contentEnd, '#c7254e');
    text.setBackgroundColor(startOffset, contentEnd, '#f9f2f4');

    Logger.log('Applied inline code format to: ' + content);

    found = body.findText('`(.+?)`');
  }
}
