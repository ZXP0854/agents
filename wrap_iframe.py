#!/usr/bin/env python3
"""将 12 个 Agent HTML 包装为 iframe + base64 data URI，彻底隐藏 script/style 标签"""
import os
import base64

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(BASE_DIR, 'agents')
OUT_DIR = os.path.join(BASE_DIR, 'agents_iframe')

os.makedirs(OUT_DIR, exist_ok=True)

for fname in sorted(os.listdir(SRC_DIR)):
    if not fname.endswith('.html'):
        continue

    src_path = os.path.join(SRC_DIR, fname)
    with open(src_path, 'r', encoding='utf-8') as f:
        inner_html = f.read()

    # Fix postMessage for double-iframe nesting
    inner_html = inner_html.replace('window.parent.postMessage', 'window.top.postMessage')

    # Base64 encode the HTML (UTF-8)
    b64 = base64.b64encode(inner_html.encode('utf-8')).decode('ascii')

    wrapper = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>绿色消费产品介绍</title>
</head>
<body style="margin:0;padding:0;width:100%;height:100vh;overflow:hidden;">
<iframe src="data:text/html;base64,{b64}" style="border:none;width:100%;height:100%;display:block;" title="绿色消费产品介绍">
</iframe>
</body>
</html>'''

    out_path = os.path.join(OUT_DIR, fname)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(wrapper)

    print(f'OK: {fname} ({len(wrapper)} bytes)')

print(f'\nDone. Files saved to {OUT_DIR}')
