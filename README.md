# docx2html
Word to HTML


Export an HTLM table to Excel, or export an Excel table to HTML.

## How to use

```python
#!/usr/bin/env python3

import docx2html

generator = DocxHTMLGenerator('path/file.docx')

with open('output.html', 'w') as f:
    f.write(generator.to_html())
```
