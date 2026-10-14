# Crops each team logo to a square and writes web versions to assets/teams.
# Usage: python3 make-logos.py [team_logos folder] [output folder]   (needs Pillow)
import sys, os
from PIL import Image
here = os.path.dirname(os.path.abspath(__file__))
src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, "..", "team_logos")
out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(here, "assets", "teams")
# team id -> (file, top offset for the square crop as a fraction of spare height; None = centre)
files = {
  1: ('Team1_Rajarajeshwari _United.JPG', None),
  2: ('Team2_Jeem_Boom_Baa.jpg', None),
  3: ('Team3_Dr._D.jpg', 0.42),     # portrait: keep the ring and the title
  4: ('Team4_Thoofan.jpg', 0.45),   # portrait: centre on the shield
  5: ('Team5_Dhurandhar.JPG', None),
  6: ('Team6_Aaram_Thamburan.JPG', None),
}
for tid, (name, frac) in files.items():
    im = Image.open(os.path.join(src, name)).convert('RGB')
    w, h = im.size; s = min(w, h)
    x0 = (w - s) // 2
    y0 = int((h - s) * (0.5 if frac is None else frac))
    sq = im.crop((x0, y0, x0 + s, y0 + s))
    for size, suffix, q in ((400, '', 82), (128, '-sm', 80)):
        p = os.path.join(out, f'team-{tid}{suffix}.webp')
        sq.resize((size, size), Image.LANCZOS).save(p, 'WEBP', quality=q, method=6)
    print(tid, name, (w, h), '-> crop', (x0, y0, s))
