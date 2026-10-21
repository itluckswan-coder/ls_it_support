"""Build a one-file copy of the manual page with the three content files inlined.

    python3 docs/manual/build.py <out-dir>

Writes <out-dir>/manual.html; the pictures stay in docs/manual/img and are published beside it.
After editing content.*.js or index.html, run this and publish the result again.
"""
import pathlib, re, sys

here = pathlib.Path(__file__).parent
out = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else here / 'dist')
out.mkdir(parents=True, exist_ok=True)
html = (here / 'index.html').read_text(encoding='utf-8')

def inline(m):
    src = here / m.group(1)
    return '<script>\n' + src.read_text(encoding='utf-8') + '\n</script>' if src.exists() else ''

html = re.sub(r'<script src="(content\.[a-z]+\.js)"></script>', inline, html)
(out / 'manual.html').write_text(html, encoding='utf-8')
print(out / 'manual.html')
