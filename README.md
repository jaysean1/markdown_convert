# Markdown Converter - Google Docs Add-on

This is a Google Docs add-on that converts Markdown syntax to Google Docs native formatting with one click.

## Features

### Supported Markdown Syntax

#### 1. Headers
```markdown
# Heading 1
## Heading 2
### Heading 3
#### Heading 4
##### Heading 5
###### Heading 6
```

#### 2. Text Formatting
```markdown
**Bold text** or __Bold text__
*Italic text* or _Italic text_
***Bold and italic*** or ___Bold and italic___
~~Strikethrough text~~
```

#### 3. Lists
```markdown
- Unordered list item 1
- Unordered list item 2
* Can also use asterisks

1. Ordered list item 1
2. Ordered list item 2
3. Ordered list item 3
```

#### 4. Blockquotes
```markdown
> This is a quoted text
> Will display as indented gray text
```

#### 5. Tables
```markdown
| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Data A   | Data B   | Data C   |
```

#### 6. Code
```markdown
This is an `inline code` example

```
This is a code block
Can span multiple lines
Displays in monospace font
```
```

---

## Installation Steps

### Method 1: Create Directly in Google Docs

1. **Open Google Docs**
   - Create a new document or open an existing one

2. **Open Apps Script Editor**
   - Click the top menu: `Extensions` → `Apps Script`

3. **Delete Default Code**
   - Delete the default `function myFunction() {}` code in the editor

4. **Paste Add-on Code**
   - Copy all code from the `Code.gs` file
   - Paste it into the Apps Script editor

5. **Save Project**
   - Click the save icon (or press Ctrl+S / Cmd+S)
   - Name the project "Markdown Converter"

6. **Authorize Add-on**
   - Click the run button (▶️) or return to the document and refresh the page
   - First run will request authorization, click "Review permissions"
   - Select your Google account
   - Click "Advanced" → "Go to Markdown Converter (unsafe)"
   - Click "Allow"

7. **Done!**
   - Return to Google Docs and refresh the page
   - You will see a new "Markdown Converter" menu in the menu bar

---

## Usage

1. **Enter Markdown-formatted text in Google Docs**
2. **Click Convert** - Click menu bar `Markdown Converter` → `Convert Entire Document`
3. **Wait for conversion to complete** - A message will appear: "✅ Markdown conversion complete!"
