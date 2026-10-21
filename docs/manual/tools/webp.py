"""Turn the PNG screenshots from shots.js into the manual's 1100px WebP pictures.

    python3 docs/manual/tools/webp.py /tmp/manual-raw
"""
import glob, os, sys
from PIL import Image

src = sys.argv[1]
dst = os.path.join(os.path.dirname(__file__), '..', 'img')
for p in sorted(glob.glob(os.path.join(src, '*.png'))):
    name = os.path.basename(p)[:-4]
    im = Image.open(p).convert('RGB')
    if name.startswith('staff-plus'):
        im = im.crop((0, 0, 1280, min(im.height, 1100)))
    im = im.resize((1100, int(im.height * 1100 / im.width)), Image.LANCZOS)
    im.save(os.path.join(dst, name + '.webp'), 'WEBP', quality=78, method=6)
    print(name)
