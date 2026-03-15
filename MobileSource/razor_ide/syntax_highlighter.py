"""
Razor syntax highlighter for tkinter Text widget.
Supports HTML, C# (@code blocks, inline expressions), and Razor directives.
"""

import re

# Tag configuration: (pattern, tag_name)
# Order matters — later patterns can override earlier ones.
PATTERNS = [
    # Razor directives: @page, @using, @inherits, @inject, etc.
    (r'@(page|using|inherits|inject|implements|layout|namespace|attribute|typeparam)\b[^\n]*', 'directive'),

    # C# @code { ... } block content
    # Handled specially in the highlighter

    # Razor inline expressions: @variableName, @DateTime.Now, @someMethod()
    (r'@(?!code\b)\w[\w.()"\s]*', 'razor_expr'),

    # HTML tags: <tag ...> and </tag>
    (r'</?[\w][\w.-]*', 'html_tag'),
    (r'/?\s*>', 'html_tag'),

    # HTML attribute names
    (r'\b[\w@:-]+(?==)', 'html_attr'),

    # HTML attribute values (quoted strings)
    (r'"[^"]*"', 'string'),
    (r"'[^']*'", 'string'),

    # C# keywords (applied inside @code blocks)
    (r'\b(abstract|as|base|bool|break|byte|case|catch|char|checked|class|const|continue|decimal|default|delegate|do|double|else|enum|event|explicit|extern|false|finally|fixed|float|for|foreach|goto|if|implicit|in|int|interface|internal|is|lock|long|namespace|new|null|object|operator|out|override|params|private|protected|public|readonly|ref|return|sbyte|sealed|short|sizeof|stackalloc|static|string|struct|switch|this|throw|true|try|typeof|uint|ulong|unchecked|unsafe|ushort|using|var|virtual|void|volatile|while)\b', 'cs_keyword'),

    # C# types
    (r'\b(String|Int32|Int64|Boolean|DateTime|List|Dictionary|Task|Action|Func|IEnumerable|IList)\b', 'cs_type'),

    # Numbers
    (r'\b\d+\.?\d*\b', 'number'),

    # Single-line comments
    (r'//[^\n]*', 'comment'),

    # Multi-line comments
    (r'/\*[\s\S]*?\*/', 'comment'),

    # HTML comments
    (r'<!--[\s\S]*?-->', 'comment'),
]

# Color themes
THEME_DARK = {
    'directive':   {'foreground': '#C586C0'},  # purple
    'razor_expr':  {'foreground': '#DCDCAA'},  # yellow
    'html_tag':    {'foreground': '#569CD6'},  # blue
    'html_attr':   {'foreground': '#9CDCFE'},  # light blue
    'string':      {'foreground': '#CE9178'},  # orange
    'cs_keyword':  {'foreground': '#569CD6'},  # blue
    'cs_type':     {'foreground': '#4EC9B0'},  # teal
    'number':      {'foreground': '#B5CEA8'},  # green
    'comment':     {'foreground': '#6A9955'},  # green
    'code_block':  {'background': '#1E1E2E'},  # slightly different bg
}

THEME_LIGHT = {
    'directive':   {'foreground': '#AF00DB'},
    'razor_expr':  {'foreground': '#795E26'},
    'html_tag':    {'foreground': '#800000'},
    'html_attr':   {'foreground': '#FF0000'},
    'string':      {'foreground': '#A31515'},
    'cs_keyword':  {'foreground': '#0000FF'},
    'cs_type':     {'foreground': '#267F99'},
    'number':      {'foreground': '#098658'},
    'comment':     {'foreground': '#008000'},
    'code_block':  {'background': '#F5F5F0'},
}


def configure_tags(text_widget, theme='dark'):
    """Apply tag styles to a tkinter Text widget."""
    colors = THEME_DARK if theme == 'dark' else THEME_LIGHT
    for tag_name, config in colors.items():
        text_widget.tag_configure(tag_name, **config)
    # Ensure later-defined tags have higher priority
    for i, (_, tag_name) in enumerate(PATTERNS):
        text_widget.tag_raise(tag_name)
    text_widget.tag_raise('comment')


def highlight(text_widget):
    """Apply syntax highlighting to all text in the widget."""
    content = text_widget.get('1.0', 'end-1c')

    # Remove old tags
    for _, tag_name in PATTERNS:
        text_widget.tag_remove(tag_name, '1.0', 'end')
    text_widget.tag_remove('code_block', '1.0', 'end')

    # Find @code { ... } blocks and highlight the background
    for m in re.finditer(r'@code\s*\{', content):
        start = m.start()
        # Find matching closing brace
        brace_depth = 1
        pos = m.end()
        while pos < len(content) and brace_depth > 0:
            if content[pos] == '{':
                brace_depth += 1
            elif content[pos] == '}':
                brace_depth -= 1
            pos += 1
        end = pos
        start_idx = f"1.0+{start}c"
        end_idx = f"1.0+{end}c"
        text_widget.tag_add('code_block', start_idx, end_idx)

    # Apply all patterns
    for pattern, tag_name in PATTERNS:
        for m in re.finditer(pattern, content):
            start = m.start()
            end = m.end()
            start_idx = f"1.0+{start}c"
            end_idx = f"1.0+{end}c"
            text_widget.tag_add(tag_name, start_idx, end_idx)
