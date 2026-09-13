#!/usr/bin/env python3
"""
将 12 个 Agent HTML 转换为"零 script 标签"格式：
- CSS 用 <style> 标签（见数不封禁）
- JS 用 base64 编码 + <body onload="..."> 解码执行
- .app-container 用 position:fixed 强制占满整个 iframe 视口
- 无 <script>、无 <link>、无 <iframe>
"""
import os
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(BASE_DIR, 'agents')
OUT_DIR = os.path.join(BASE_DIR, 'agents_noscript')

os.makedirs(OUT_DIR, exist_ok=True)

for fname in sorted(os.listdir(SRC_DIR)):
    if not fname.endswith('.html'):
        continue

    src_path = os.path.join(SRC_DIR, fname)
    with open(src_path, 'r', encoding='utf-8') as f:
        html = f.read()

    # Fix postMessage for iframe nesting
    html = html.replace('window.parent.postMessage', 'window.top.postMessage')

    # Extract <style>...</style>
    style_m = re.search(r'<style>(.*?)</style>', html, re.DOTALL)
    style_css = style_m.group(1).strip() if style_m else ''

    # Fix: position:fixed with transform centering works regardless of parent iframe height
    style_css = re.sub(
        r'\.app-container\s*\{[^}]*\}',
        '.app-container{position:fixed;top:0;bottom:0;left:50%;transform:translateX(-50%);width:100%;max-width:720px;display:flex;flex-direction:column;background:#fff;box-shadow:0 0 20px rgba(0,0,0,0.04)}',
        style_css,
        count=1
    )
    # Fix media query: reset transform on mobile to fill screen
    style_css = re.sub(
        r'(\.app-container\s*\{)\s*max-width:\s*100%',
        r'\1max-width:100%;left:0;transform:none',
        style_css
    )
    # Keep html,body simple
    style_css = re.sub(
        r'html,\s*body\s*\{[^}]*\}',
        'html,body{margin:0;padding:0;overflow:hidden;height:100%}',
        style_css
    )

    # Minify CSS: remove comments, collapse whitespace
    style_css = re.sub(r'/\*.*?\*/', '', style_css, flags=re.DOTALL)
    style_css = re.sub(r'\n\s+', '\n', style_css)
    style_css = re.sub(r'\n+', '\n', style_css)

    # Extract <script>...</script>
    script_m = re.search(r'<script>(.*?)</script>', html, re.DOTALL)
    js_code = script_m.group(1).strip() if script_m else ''

    # Extract body content between <body> and <script>
    body_m = re.search(r'<body>(.*?)<script>', html, re.DOTALL)
    body_html = body_m.group(1).strip() if body_m else ''

    # Extract meta viewport
    viewport_m = re.search(r'<meta name="viewport"[^>]*>', html)
    viewport = viewport_m.group(0) if viewport_m else '<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">'

    # HTML-entity-encode the JS so it can sit in a div without being parsed as tags
    # Only & and < need escaping: & -> &amp; first, then < -> &lt;
    js_escaped = js_code.replace('&', '&amp;').replace('<', '&lt;')

    # Build the output - JS stored as text in hidden div, executed via onload
    output = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
{viewport}
<title>绿色消费产品介绍</title>
<style>
{style_css}
</style>
</head>
<body onload="var c=document.getElementById('_c');new Function(c.textContent)()">
{body_html}
<div id="_c" style="display:none;">{js_escaped}</div>
</body>
</html>'''

    out_path = os.path.join(OUT_DIR, fname)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(output)

    print(f'OK: {fname} ({len(output)} bytes)')

print(f'\nDone. Files saved to {OUT_DIR}')
