import os
import sys
import re

# Self-bootstrap dependencies if missing
try:
    import docx
    from bs4 import BeautifulSoup
except ImportError:
    import subprocess
    print("Dependencies 'python-docx' or 'beautifulsoup4' not found. Installing now...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx", "beautifulsoup4"])
    import docx
    from bs4 import BeautifulSoup

from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- XML Helper Utilities for Word Styling ---

def set_cell_shading(cell, color_hex):
    """Apply background fill color to a cell."""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:val'), 'clear')
    shading.set(qn('w:color'), 'auto')
    shading.set(qn('w:fill'), color_hex)
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_left_border(cell, color_hex, size="36"):
    """Apply a thick left border to a cell (4.5 pt) and clear other borders."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    
    # Left Border
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single')
    left.set(qn('w:sz'), size)
    left.set(qn('w:space'), '0')
    left.set(qn('w:color'), color_hex)
    tcBorders.append(left)
    
    # Clear top, bottom, and right borders
    for border_name in ['top', 'bottom', 'right']:
        b = OxmlElement(f'w:{border_name}')
        b.set(qn('w:val'), 'nil')
        tcBorders.append(b)
        
    tcPr.append(tcBorders)

# --- Dynamic Parser for data.js ---

def load_data_js(filepath):
    """Parse the javascript data file into a list of page dicts."""
    if not os.path.exists(filepath):
        print(f"Error: File {filepath} not found.")
        sys.exit(1)
        
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # Find the array block in DOCS_DATA
    docs_match = re.search(r'const\s+DOCS_DATA\s*=\s*\[([\s\S]*)\];\s*$', content)
    if not docs_match:
        # Fallback to search any part of the file
        docs_match = re.search(r'const\s+DOCS_DATA\s*=\s*\[([\s\S]*)\]', content)
        if not docs_match:
            print("Error: Could not locate DOCS_DATA array inside data.js.")
            sys.exit(1)
            
    array_content = docs_match.group(1)

    # Regex mapping properties of each object. Handles template strings `content`.
    item_pattern = re.compile(
        r'\{\s*'
        r'id:\s*"([^"]+)",\s*'
        r'title:\s*"([^"]+)",\s*'
        r'icon:\s*"([^"]+)",\s*'
        r'path:\s*\[([^\]]*)\],\s*'
        r'lastUpdated:\s*"([^"]+)",\s*'
        r'content:\s*`([\s\S]*?)(?<!\\)`\s*'
        r'\}',
        re.MULTILINE
    )

    matches = item_pattern.findall(content)
    docs = []
    for m in matches:
        doc_id, title, icon, path_str, last_updated, doc_content = m
        # Extract path items
        path = [p.strip().strip('"').strip("'") for p in path_str.split(',') if p.strip()]
        docs.append({
            'id': doc_id,
            'title': title,
            'icon': icon,
            'path': path,
            'lastUpdated': last_updated,
            'content': doc_content
        })
        
    return docs

def save_data_js(filepath, docs):
    """Write the page dicts back to js/data.js in template string format."""
    file_content = "// Unified data storage for HomeFront Confluence documentation system\nconst DOCS_DATA = [\n"
    
    for idx, doc in enumerate(docs):
        file_content += "  {\n"
        file_content += f'    id: "{doc["id"]}",\n'
        file_content += f'    title: "{doc["title"]}",\n'
        file_content += f'    icon: "{doc["icon"]}",\n'
        # Format path string list
        paths = ", ".join([f'"{p}"' for p in doc["path"]])
        file_content += f'    path: [{paths}],\n'
        file_content += f'    lastUpdated: "{doc["lastUpdated"]}",\n'
        
        # Escape any unescaped backticks inside the content template strings
        sanitized_content = doc["content"].replace('`', '\\`').replace('${', '\\${')
        file_content += f'    content: `{sanitized_content}`\n'
        file_content += f'  }}{"," if idx < len(docs) - 1 else ""}\n'
        
    file_content += "];\n"
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(file_content)
    print(f"Successfully wrote documentation pages to {filepath}")

# --- HTML ➔ DOCX Formatting Exporter ---

def add_nested_runs(run, element, bold=False, italic=False, code=False):
    """Helper to write nested formatting styles."""
    run.text = element.get_text()
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    if code:
        run.font.name = 'Courier New'
        run.font.size = Pt(10)

def add_runs_from_html(paragraph, element):
    """Parse inline text nodes and apply bold, italic, code or link properties in Word."""
    for child in element.children:
        if isinstance(child, str):
            if child.strip('\n'):
                paragraph.add_run(child)
        elif child.name in ['strong', 'b']:
            run = paragraph.add_run()
            add_nested_runs(run, child, bold=True)
        elif child.name in ['em', 'i']:
            run = paragraph.add_run()
            add_nested_runs(run, child, italic=True)
        elif child.name == 'code':
            run = paragraph.add_run()
            add_nested_runs(run, child, code=True)
        elif child.name == 'a':
            run = paragraph.add_run()
            run.font.underline = True
            run.font.color.rgb = RGBColor(0, 82, 204) # Confluence blue links
            add_nested_runs(run, child)
        else:
            add_runs_from_html(paragraph, child)

def html_to_docx(doc, html_content):
    """Parse page content and add it to the Word document structure."""
    soup = BeautifulSoup(html_content, 'html.parser')
    for child in soup.children:
        if child.name is None:
            continue
            
        if child.name in ['h2', 'h3', 'h4']:
            level = int(child.name[1]) - 1 # h2 maps to Heading 2 in Word
            doc.add_heading(child.get_text(), level=level)
            
        elif child.name == 'p':
            p = doc.add_paragraph()
            add_runs_from_html(p, child)
            
        elif child.name in ['ul', 'ol']:
            style = 'List Bullet' if child.name == 'ul' else 'List Number'
            for li in child.find_all('li', recursive=False):
                p = doc.add_paragraph(style=style)
                add_runs_from_html(p, li)
                
        elif child.name == 'pre':
            # Format block code
            code_tag = child.find('code')
            code_text = code_tag.get_text() if code_tag else child.get_text()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.25)
            # Add grey shading and monospaced font
            run = p.add_run(code_text)
            run.font.name = 'Courier New'
            run.font.size = Pt(9.5)
            
        elif child.name == 'div' and 'callout' in child.get('class', []):
            # Form callout block as a bordered, shaded single-cell table in Word
            classes = child.get('class', [])
            callout_type = 'note'
            for c in ['note', 'tip', 'warn', 'danger']:
                if c in classes:
                    callout_type = c
                    break
                    
            colors = {
                'note': ('DEEBFF', '0052CC'),
                'tip': ('E3FCEF', '36B37E'),
                'warn': ('FFFAE6', 'FFAB00'),
                'danger': ('FFEBE6', 'FF5630')
            }
            bg_color, border_color = colors.get(callout_type, ('F4F5F7', 'DFE1E6'))
            
            table = doc.add_table(rows=1, cols=1)
            table.autofit = False
            cell = table.cell(0, 0)
            set_cell_shading(cell, bg_color)
            set_cell_left_border(cell, border_color)
            
            cell_p = cell.paragraphs[0]
            # Strip outer headers if callout-body wrapper exists
            cell_soup = child.find('div', class_='callout-body') or child
            add_runs_from_html(cell_p, cell_soup)
            
        elif child.name == 'table':
            rows = child.find_all('tr')
            if not rows:
                continue
                
            max_cols = 0
            for r in rows:
                cols = r.find_all(['td', 'th'])
                max_cols = max(max_cols, len(cols))
                
            w_table = doc.add_table(rows=len(rows), cols=max_cols)
            w_table.style = 'Table Grid'
            
            for row_idx, r in enumerate(rows):
                cols = r.find_all(['td', 'th'])
                for col_idx, c in enumerate(cols):
                    w_cell = w_table.cell(row_idx, col_idx)
                    w_cell_p = w_cell.paragraphs[0]
                    add_runs_from_html(w_cell_p, c)
                    # Headers shading
                    if c.name == 'th':
                        set_cell_shading(w_cell, 'F4F5F7')

# --- DOCX ➔ HTML Compilation Importer ---

def iter_block_items(doc):
    """Traverse document elements in chronological order."""
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    
    parent_element = doc.element.body
    for child in parent_element.iterchildren():
        if child.tag.endswith('p'):
            yield Paragraph(child, doc)
        elif child.tag.endswith('tbl'):
            yield Table(child, doc)

def compile_paragraph_runs_to_html(p):
    """Convert a Word paragraph with bold/italic runs into clean HTML string."""
    html_parts = []
    for run in p.runs:
        text = run.text
        if not text:
            continue
            
        # Escape HTML symbols
        text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        
        is_bold = run.bold or ('bold' in run.font.name.lower() if run.font.name else False)
        is_italic = run.italic
        is_code = run.font.name == 'Courier New'
        
        if is_code:
            text = f'<code>{text}</code>'
        if is_bold:
            text = f'<strong>{text}</strong>'
        if is_italic:
            text = f'<em>{text}</em>'
            
        html_parts.append(text)
        
    return "".join(html_parts)

def compile_callout_table_to_html(table):
    """Translate a 1x1 styled cell back into a styled HTML callout block."""
    cell = table.cell(0, 0)
    
    # Try to identify callout class type based on cell background color
    bg_color = 'note'
    try:
        shading = cell._tc.get_or_add_tcPr().find(qn('w:shd'))
        if shading is not None:
            fill = shading.get(qn('w:fill')).upper()
            if fill == 'DEEBFF': bg_color = 'note'
            elif fill == 'E3FCEF': bg_color = 'tip'
            elif fill == 'FFFAE6': bg_color = 'warn'
            elif fill == 'FFEBE6': bg_color = 'danger'
    except Exception:
        pass
        
    icon_map = {'note': 'ℹ️', 'tip': '💡', 'warn': '⚠️', 'danger': '🚨'}
    icon = icon_map.get(bg_color, 'ℹ️')
    
    # Render all paragraphs inside the cell
    cell_html = []
    for p in cell.paragraphs:
        p_html = compile_paragraph_runs_to_html(p)
        if p_html:
            cell_html.append(p_html)
            
    content = " ".join(cell_html)
    return (
        f'<div class="callout {bg_color}">\n'
        f'  <div class="callout-icon">{icon}</div>\n'
        f'  <div class="callout-body">\n'
        f'    {content}\n'
        f'  </div>\n'
        f'</div>'
    )

def compile_table_to_html(table):
    """Translate standard Word grid tables back to semantic HTML tables."""
    html = '<table>\n'
    # Check if first row is all bold / header shaded
    first_row_cells = table.rows[0].cells
    is_header = False
    try:
        # Check background of cell 1
        shading = first_row_cells[0]._tc.get_or_add_tcPr().find(qn('w:shd'))
        if shading is not None and shading.get(qn('w:fill')).upper() == 'F4F5F7':
            is_header = True
    except Exception:
        pass
        
    for r_idx, row in enumerate(table.rows):
        html += '  <tr>\n'
        cell_tag = 'th' if (r_idx == 0 and is_header) else 'td'
        for cell in row.cells:
            cell_content = []
            for p in cell.paragraphs:
                p_text = compile_paragraph_runs_to_html(p)
                if p_text:
                    cell_content.append(p_text)
            html += f'    <{cell_tag}>{" ".join(cell_content)}</{cell_tag}>\n'
        html += '  </tr>\n'
    html += '</table>'
    return html

# --- Orchestrators: Import & Export ---

def do_export(data_js_path, docx_path):
    """Orchestrate exporting js/data.js data structure into a formatted MS Word document."""
    print(f"Reading documentation from {data_js_path}...")
    docs = load_data_js(data_js_path)
    
    doc = docx.Document()
    
    # Document Style setup
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10.5)
    
    # Title Cover Page
    doc.add_heading("HomeFront Suite Documentation", level=0)
    doc.add_paragraph("This document compiles the codebase, architecture, hierarchy, database schemas, and migration registry profiles of the HomeFront suite of programs. This document can be edited and synchronized back to the web application.")
    doc.add_page_break()
    
    for idx, d in enumerate(docs):
        # Heading 1 representing the page
        h1 = doc.add_heading(d["title"], level=1)
        
        # Metadata block
        paths = " > ".join(d["path"])
        p_meta = doc.add_paragraph()
        run_meta = p_meta.add_run(f'Metadata: ID={d["id"]} | Icon={d["icon"]} | Path={paths}')
        run_meta.italic = True
        run_meta.font.color.rgb = RGBColor(110, 110, 110)
        run_meta.font.size = Pt(9.5)
        
        # Convert HTML to Word elements
        html_to_docx(doc, d["content"])
        
        if idx < len(docs) - 1:
            doc.add_page_break()
            
    doc.save(docx_path)
    print(f"Exported successfully to {docx_path}!")

def do_import(docx_path, data_js_path):
    """Orchestrate reading MS Word document, compiling HTML contents, and writing back to js/data.js."""
    if not os.path.exists(docx_path):
        print(f"Error: File {docx_path} not found.")
        sys.exit(1)
        
    print(f"Reading Word document from {docx_path}...")
    doc = docx.Document(docx_path)
    
    docs = []
    current_page = None
    page_blocks = []
    in_list = False
    
    # Skip cover page heading (only process pages starting from Heading 1 and having Metadata)
    started = False
    
    for item in iter_block_items(doc):
        if isinstance(item, docx.text.paragraph.Paragraph):
            style_name = item.style.name.lower() if item.style else ''
            text = item.text.strip()
            
            # Identify Heading 1 (Word Heading 1)
            if 'heading 1' in style_name:
                # Save previous page
                if current_page:
                    if in_list:
                        page_blocks.append('</ul>')
                        in_list = False
                    current_page['content'] = "\n".join(page_blocks).strip()
                    docs.append(current_page)
                    
                current_page = {
                    'id': '', 'title': item.text.strip(), 'icon': '📄', 'path': [], 'lastUpdated': 'June 15, 2026', 'content': ''
                }
                page_blocks = []
                started = True
                
            elif text.startswith('Metadata:') and started:
                # Parse metadata fields
                meta = {}
                parts = text.replace('Metadata:', '').split('|')
                for p in parts:
                    if '=' in p:
                        k, v = p.split('=', 1)
                        meta[k.strip().lower()] = v.strip()
                        
                if current_page:
                    current_page['id'] = meta.get('id', '')
                    current_page['icon'] = meta.get('icon', '📄')
                    current_page['path'] = [x.strip() for x in meta.get('path', '').split('>')]
                    
            elif started and current_page:
                # Process content
                if 'heading 2' in style_name:
                    if in_list:
                        page_blocks.append('</ul>')
                        in_list = False
                    page_blocks.append(f'<h2>{text}</h2>')
                    
                elif 'heading 3' in style_name:
                    if in_list:
                        page_blocks.append('</ul>')
                        in_list = False
                    page_blocks.append(f'<h3>{text}</h3>')
                    
                elif 'list bullet' in style_name:
                    if not in_list:
                        page_blocks.append('<ul>')
                        in_list = True
                    li_html = compile_paragraph_runs_to_html(item)
                    page_blocks.append(f'  <li>{li_html}</li>')
                    
                else:
                    # Normal paragraph or code block indent
                    if in_list:
                        page_blocks.append('</ul>')
                        in_list = False
                        
                    p_html = compile_paragraph_runs_to_html(item)
                    if p_html:
                        # Check if it was formatted as a code block (Courier New font)
                        is_code_block = False
                        try:
                            # If all runs in paragraph are courier
                            if len(item.runs) > 0 and all(r.font.name == 'Courier New' for r in item.runs):
                                is_code_block = True
                        except Exception:
                            pass
                            
                        if is_code_block:
                            # Strip nested code tags if they got wrapped
                            clean_text = item.text
                            page_blocks.append(f'<pre><code>{clean_text}</code></pre>')
                        else:
                            page_blocks.append(f'<p>{p_html}</p>')
                            
        elif isinstance(item, docx.table.Table) and started and current_page:
            # Check if this table has 1 row and 1 column (Callout block)
            if in_list:
                page_blocks.append('</ul>')
                in_list = False
                
            if len(item.rows) == 1 and len(item.rows[0].cells) == 1:
                callout_html = compile_callout_table_to_html(item)
                page_blocks.append(callout_html)
            else:
                table_html = compile_table_to_html(item)
                page_blocks.append(table_html)
                
    # Save final page
    if current_page:
        if in_list:
            page_blocks.append('</ul>')
        current_page['content'] = "\n".join(page_blocks).strip()
        docs.append(current_page)
        
    if not docs:
        print("Error: No page segments parsed from the Word document.")
        sys.exit(1)
        
    save_data_js(data_js_path, docs)

# --- CLI Entrypoint ---

def main():
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python doc_sync.py export    - Export data.js to Word docx")
        print("  python doc_sync.py import    - Import Word docx back to data.js")
        sys.exit(1)
        
    action = sys.argv[1].lower()
    
    # Path mappings
    script_dir = os.path.dirname(os.path.abspath(__file__))
    doc_dir = os.path.dirname(script_dir)
    data_js_path = os.path.join(doc_dir, "js", "data.js")
    docx_path = os.path.join(doc_dir, "HomeFront_Documentation.docx")
    
    if action == "export":
        do_export(data_js_path, docx_path)
    elif action == "import":
        do_import(docx_path, data_js_path)
    else:
        print(f"Unknown action: {action}")
        print("Must be 'export' or 'import'")
        sys.exit(1)

if __name__ == "__main__":
    main()
