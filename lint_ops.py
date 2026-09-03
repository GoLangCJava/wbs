# -*- coding: utf-8 -*-
"""Layout linter for /tmp/arch_ops.json: estimate text extents (CJK=1em, Latin~0.58em)
and flag overflow beyond declared box or slide bounds, plus text-text overlaps."""
import json, sys

SLIDE_W, SLIDE_H = 13.333, 7.5

def char_w(ch, sz):
    o = ord(ch)
    if o < 128:
        if ch == ' ': return sz * 0.30
        if ch in 'iljI.,:;\'"|!': return sz * 0.30
        if ch.isdigit() or ch.isalpha(): return sz * 0.60
        return sz * 0.42          # ascii punctuation
    if 0x2000 <= o <= 0x206F: return sz * 0.5
    return sz * 1.02              # CJK / fullwidth

def run_w(text, sz):
    return sum(char_w(c, sz) / 72.0 for c in text)   # inches

def para_lines(para, box_w):
    runs = para.get('runs', [])
    if not runs:
        return 0, 0
    maxsz = max(st.get('sz', 8) for _, st in runs) or 8
    # greedy wrap over concatenated chars
    lines, cur = 1, 0.0
    for s, st in runs:
        sz = st.get('sz', 8)
        for ch in s.replace('\n', '\n'):
            if ch == '\n':
                lines += 1; cur = 0; continue
            cw = char_w(ch, sz) / 72.0
            if cur + cw > box_w + 1e-6:
                lines += 1; cur = cw
            else:
                cur += cw
    return lines, maxsz

def est_text_h(op):
    h = 0.0
    for para in op['paras']:
        ls = para.get('ls', op.get('leading', 1.12))
        n, sz = para_lines(para, op['w'])
        h += n * (sz / 72.0) * 1.24 * (ls or 1.12)
    return h

def main():
    ops = json.load(open('/tmp/arch_ops.json', encoding='utf-8'))
    problems = 0
    boxes = []   # (slide, x, y, w, est_h, label)
    for op in ops:
        s = op['s']
        if op['t'] == 'x':
            x, y, w, h = op['x'], op['y'], op['w'], op['h']
            eh = est_text_h(op)
            over_w = []
            for para in op['paras']:
                if op.get('vert'):
                    continue
                total = sum(run_w(txt, st.get('sz', 8)) for txt, st in para.get('runs', []))
                # single-line width check only if no wrap expected: flag if > w * 1.02 and fits in 1 line? report anyway
                if total > w * 1.03 and para_lines(para, 1e9)[0] == 1:
                    over_w.append((total, w, ''.join(t for t, _ in para['runs'])[:40]))
            label = ''.join(t for p in op['paras'] for t, _ in p.get('runs', []))[:38]
            issues = []
            if eh > h * 1.12 + 0.02:
                issues.append(f'H overflow: box {h:.2f}" est {eh:.2f}"')
            for tw, bw, t in over_w:
                issues.append(f'W overflow: need {tw:.2f}" have {bw:.2f}" [{t}]')
            if x < 0.05 or y < 0.02 or x + w > SLIDE_W - 0.05 or y + max(h, eh) > SLIDE_H - 0.02:
                issues.append(f'out of slide bounds (x={x:.2f} y={y:.2f} w={w:.2f} h={max(h,eh):.2f})')
            if issues:
                problems += 1
                print(f'[S{s}] TEXT ({x:.2f},{y:.2f}) "{label}"')
                for i in issues:
                    print('   -', i)
            boxes.append((s, x, y, w, max(h, eh), label))
        elif op['t'] == 'r':
            x2, y2 = op['x'] + op['w'], op['y'] + op['h']
            if op['x'] < -0.01 or op['y'] < -0.01 or x2 > SLIDE_W + 0.01 or y2 > SLIDE_H + 0.01:
                problems += 1
                print(f"[S{s}] RECT out of bounds ({op['x']:.2f},{op['y']:.2f},{op['w']:.2f},{op['h']:.2f}) fill={op.get('fill')}")
    # text-text overlap (same slide, rough)
    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            a, b = boxes[i], boxes[j]
            if a[0] != b[0]:
                continue
            ox = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
            oy = min(a[2] + a[4], b[2] + b[4]) - max(a[2], b[2])
            if ox > 0.05 and oy > 0.05:
                problems += 1
                print(f'[S{a[0]}] TEXT-OVERLAP "{a[5]}" <-> "{b[5]}" (ox={ox:.2f} oy={oy:.2f})')
    print('---', problems, 'problems')

if __name__ == '__main__':
    main()
