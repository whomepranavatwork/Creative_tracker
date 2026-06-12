import markdown
from weasyprint import HTML, CSS
import re

with open('/home/user/Creative_tracker/S3_Hair_Regrowth_Kit_Master_Brief.md', 'r') as f:
    md_content = f.read()

html_body = markdown.markdown(md_content, extensions=['tables', 'nl2br'])

html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

  * {{ box-sizing: border-box; margin: 0; padding: 0; }}

  body {{
    font-family: 'Inter', 'Helvetica Neue', Arial, sans-serif;
    font-size: 9.5pt;
    line-height: 1.55;
    color: #1a1a1a;
    background: #fff;
    padding: 0;
  }}

  /* PAGE BREAKS */
  h1 {{ page-break-before: always; }}
  h1:first-of-type {{ page-break-before: avoid; }}

  /* COVER / TITLE */
  h1.doc-title {{
    font-size: 20pt;
    font-weight: 700;
    color: #1a1a1a;
    border-bottom: 3px solid #1a1a1a;
    padding-bottom: 8px;
    margin-bottom: 4px;
    page-break-before: avoid;
  }}

  /* TAB HEADERS (h1) */
  h1 {{
    font-size: 14pt;
    font-weight: 700;
    color: #fff;
    background: #1e3a2f;
    padding: 10px 16px;
    margin: 28px 0 18px 0;
    letter-spacing: 0.5px;
  }}

  /* SECTION HEADERS (h2) */
  h2 {{
    font-size: 11.5pt;
    font-weight: 700;
    color: #1e3a2f;
    border-bottom: 2px solid #1e3a2f;
    padding-bottom: 5px;
    margin: 22px 0 12px 0;
  }}

  /* SUB-SECTION (h3) */
  h3 {{
    font-size: 10pt;
    font-weight: 700;
    color: #fff;
    background: #2d6a4f;
    padding: 6px 12px;
    margin: 18px 0 10px 0;
  }}

  /* SUB-SUB SECTION (h4) */
  h4 {{
    font-size: 9.5pt;
    font-weight: 700;
    color: #1e3a2f;
    margin: 14px 0 6px 0;
    text-transform: uppercase;
    letter-spacing: 0.3px;
  }}

  /* TABLES */
  table {{
    width: 100%;
    border-collapse: collapse;
    margin: 8px 0 16px 0;
    font-size: 9pt;
  }}

  th {{
    background: #1e3a2f;
    color: #fff;
    font-weight: 600;
    padding: 7px 10px;
    text-align: left;
    border: 1px solid #1e3a2f;
  }}

  td {{
    padding: 6px 10px;
    border: 1px solid #d0d0d0;
    vertical-align: top;
  }}

  tr:nth-child(even) td {{
    background: #f4f8f6;
  }}

  /* First column bold in 2-col tables */
  td:first-child {{
    font-weight: 600;
    color: #1e3a2f;
    width: 28%;
  }}

  /* 3-col comparison tables — first col narrower */
  table td:first-child + td + td {{
    width: auto;
  }}

  /* BLOCKQUOTE = NEVER MENTION box */
  blockquote {{
    background: #fff3f3;
    border-left: 4px solid #c0392b;
    padding: 10px 14px;
    margin: 10px 0 18px 0;
    font-size: 9pt;
    color: #333;
    font-style: normal;
  }}

  /* LISTS */
  ul, ol {{
    margin: 6px 0 10px 18px;
    padding: 0;
  }}

  li {{
    margin-bottom: 3px;
  }}

  /* PARAGRAPHS */
  p {{
    margin: 6px 0 10px 0;
  }}

  strong {{
    font-weight: 700;
    color: #1e3a2f;
  }}

  em {{
    font-style: italic;
    color: #444;
  }}

  hr {{
    border: none;
    border-top: 1px solid #e0e0e0;
    margin: 16px 0;
  }}

  /* FOOTER */
  @page {{
    size: A4;
    margin: 18mm 16mm 18mm 16mm;
    @bottom-right {{
      content: counter(page);
      font-size: 8pt;
      color: #999;
    }}
    @bottom-left {{
      content: "Man Matters | S3 Hair Regrowth Kit — Master Brief";
      font-size: 8pt;
      color: #999;
    }}
  }}
</style>
</head>
<body>
{html_body}
</body>
</html>"""

HTML(string=html).write_pdf('/home/user/Creative_tracker/S3_Hair_Regrowth_Kit_Master_Brief.pdf')
print("PDF generated successfully.")
