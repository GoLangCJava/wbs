# -*- coding: utf-8 -*-
"""Render /tmp/arch_ops.json (produced by gen_*_ppt.py) to PNG for layout preview.
CJK glyphs unavailable (DejaVu only) -> Chinese shows as boxes; geometry & Latin text still readable."""
import json, math, os
from PIL import Image, ImageDraw, ImageFont

SCALE = 150
W, H = int(13.333 * SCALE), int(7.5 * SCALE)
REG = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
BOLD = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
_fcache = {}

def font(sz_pt, bold=False):
    px = max(6, int(round(sz_pt / 72 * SCALE)))
    key = (px, bold)
    if key not in _fcache:
        _fcache[key] = ImageFont.truetype(BOLD if bold else REG, px)
    return _fcache[key]

def hx(c):
    return tuple(int(c[i:i+2], 16) for i in (0, 2, 4)) if c else None

def dashes(d, p1, p2, w):
    L = math.dist(p1, p2)
    n = max(2, int(L / (w * 4 + 6)))
    return [((p1[0] + (p2[0]-p1[0]) * i / n, p1[1] + (p2[1]-p1[1]) * i / n),
             (p1[0] + (p2[0]-p1[0]) * (i+1) / n, p1[1] + (p2[1]-p1[1]) * (i+1) / n))
            for i in range(0, n, 2)]

def arrow(d, p1, p2, c, w):
    d.line([p1, p2], fill=c, width=max(1, int(w * 1.3)))
    ang = math.atan2(p2[1]-p1[1], p2[0]-p1[0])
    L = 5 + w * 2.2
    for s in (2.6, -2.6):
        d.line([p2, (p2[0] - L*math.cos(ang - s), p2[1] - L*math.sin(ang - s))], fill=c, width=max(1, int(w*1.3)))

def wrap(runs, w_px):
    """runs: [(text, style)] -> lines: [[(txt, font, color)...]]"""
    lines, cur, curw = [], [], 0
    for text, st in runs:
        f = font(st.get('sz', 8), st.get('b', False))
        for ch in text:
            cw = f.getlength(ch)
            if ch == '\n' or curw + cw > w_px:
                lines.append(cur); cur, curw = [], 0
                if ch == '\n':
                    continue
            cur.append((ch, f, hx(st.get('c', '0F172A')))); curw += cw
    if cur:
        lines.append(cur)
    return lines

def draw_text(d, op):
    x, y, w, h = [v * SCALE for v in (op['x'], op['y'], op['w'], op['h'])]
    align, anchor, leading = op.get('align', 'l'), op.get('anchor', 't'), op.get('leading', 1.12)
    all_lines = []
    for para in op['paras']:
        runs = [(s, st) for s, st in para.get('runs', [])]
        if not runs:
            all_lines.append([]); continue
        ls = para.get('ls', leading)
        lw = wrap(runs, w)
        for i, ln in enumerate(lw):
            all_lines.append((ln, ls if i == 0 else ls))
    if not all_lines:
        return
    # line heights
    hs, maxsz = [], 7
    for item in all_lines:
        ln, ls = (item if isinstance(item, tuple) else (item, leading))
        sz = max([f.size for _, f, _ in ln] or [8])
        maxsz = max(maxsz, sz)
        hs.append(sz * 1.25 * ls)
    total = sum(hs)
    if anchor == 'm':
        yy = y + (h - total) / 2
    elif anchor == 'b':
        yy = y + h - total
    else:
        yy = y
    vert = op.get('vert', False)
    for item, hh in zip(all_lines, hs):
        ln, _ = (item if isinstance(item, tuple) else (item, leading))
        lw_px = sum(f.getlength(ch) for ch, f, _ in ln)
        if vert:
            cx = x + w / 2
            sy = yy + hh / 2 - (sum(f.size for f in {id(f): f for _, f, _ in ln}.values()) * len(ln)) / 2
            tot = sum(f.size for _, f, _ in ln)
            cy = yy + (h - tot) / 2 if anchor == 'm' else yy
            for ch, f, c in ln:
                d.text((cx - f.size / 2, cy), ch, font=f, fill=c or (0, 0, 0))
                cy += f.size
            continue
        if align == 'c':
            xx = x + (w - lw_px) / 2
        elif align == 'r':
            xx = x + w - lw_px
        else:
            xx = x
        for ch, f, c in ln:
            d.text((xx, yy + (hh - f.size * 1.25) / 2 + f.size * 0.08), ch, font=f, fill=c or (0, 0, 0))
            xx += f.getlength(ch)
        yy += hh

def main():
    ops = json.load(open('/tmp/arch_ops.json', encoding='utf-8'))
    ns = max(op['s'] for op in ops) + 1
    os.makedirs('/tmp/preview', exist_ok=True)
    for s in range(ns):
        img = Image.new('RGB', (W, H), 'white')
        d = ImageDraw.Draw(img)
        for op in ops:
            if op['s'] != s:
                continue
            t = op['t']
            if t == 'r':
                x, y, w, h = [v * SCALE for v in (op['x'], op['y'], op['w'], op['h'])]
                fill, line = hx(op.get('fill')), hx(op.get('line'))
                lw = max(1, int((op.get('lw') or 1) * 1.4))
                rad = (op.get('rad') or 0) * SCALE
                if op.get('dash'):
                    d.rectangle([x, y, x + w, y + h], fill=fill, outline=line, width=lw)
                elif rad:
                    d.rounded_rectangle([x, y, x + w, y + h], radius=rad, fill=fill, outline=line, width=lw)
                else:
                    d.rectangle([x, y, x + w, y + h], fill=fill, outline=line, width=lw)
            elif t == 'l':
                p1 = (op['x1'] * SCALE, op['y1'] * SCALE)
                p2 = (op['x2'] * SCALE, op['y2'] * SCALE)
                c = hx(op.get('c', '64748B')) or (100, 100, 100)
                w = op.get('w', 1)
                segs = dashes(None, p1, p2, w) if op.get('dash') else [(p1, p2)]
                for a, b in segs:
                    d.line([a, b], fill=c, width=max(1, int(w * 1.3)))
                if op.get('ar'):
                    ang = math.atan2(p2[1]-p1[1], p2[0]-p1[0])
                    L = 5 + w * 2.2
                    for s2 in (2.6, -2.6):
                        d.line([p2, (p2[0]-L*math.cos(ang-s2), p2[1]-L*math.sin(ang-s2))], fill=c, width=max(1, int(w*1.3)))
            elif t == 'x':
                draw_text(d, op)
            elif t == 'arc':
                cx, cy, r = op['cx'] * SCALE, op['cy'] * SCALE, op.get('r', 0.03) * SCALE
                d.arc([cx - r, cy - r, cx + r, cy + r], 180, 360, fill=hx(op.get('c', '0078D4')) or (0, 120, 212), width=1)
        img.save(f'/tmp/preview/slide_{s}.png')
    print('rendered', ns, 'slides ->', '/tmp/preview/')

if __name__ == '__main__':
    main()
