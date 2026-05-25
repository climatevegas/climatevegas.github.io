#!/usr/bin/env python3
"""Soft Berry Bloom — Gentle feminine curves with warm rose-plum tones.
Nurturing, hopeful, elegant. Like petals unfurling at dawn."""
import os, sys

PAGE_W, PAGE_H = 794, 1123
C = {'bg':'#faf7f5','blush':'#c4a07a','plum':'#8e6b7a','rose':'#b8897a',
     'cream':'#e8ddd0','mist':'#d4c8be','gold':'#a08050','soft':'#c9a99a'}

DEFS = f'''<defs>
<linearGradient id="gw" x1="0" y1="0" x2="1" y2="0.3">
 <stop offset="0%" stop-color="{C['blush']}" stop-opacity="0.22"/>
 <stop offset="50%" stop-color="{C['rose']}" stop-opacity="0.14"/>
 <stop offset="100%" stop-color="{C['plum']}" stop-opacity="0.10"/>
</linearGradient>
<linearGradient id="gc" x1="0" y1="0" x2="1" y2="0">
 <stop offset="0%" stop-color="{C['cream']}" stop-opacity="0.32"/>
 <stop offset="100%" stop-color="{C['mist']}" stop-opacity="0.18"/>
</linearGradient>
<linearGradient id="gv" x1="0" y1="0" x2="0" y2="1">
 <stop offset="0%" stop-color="{C['blush']}" stop-opacity="0.07"/>
 <stop offset="40%" stop-color="{C['bg']}" stop-opacity="0"/>
 <stop offset="100%" stop-color="{C['plum']}" stop-opacity="0.06"/>
</linearGradient>
<filter id="gr"><feTurbulence type="fractalNoise" baseFrequency="0.5"
 numOctaves="3" stitchTiles="stitch" result="n"/>
 <feColorMatrix type="saturate" values="0" in="n" result="m"/>
 <feBlend in="SourceGraphic" in2="m" mode="multiply"/></filter>
</defs>'''

SVG = lambda body: (f'<!DOCTYPE html><html><head><meta charset="utf-8">'
 f'<style>*{{margin:0;padding:0}}body{{width:{PAGE_W}px;height:{PAGE_H}px;'
 f'background:{C["bg"]}}}</style></head><body>'
 f'<svg width="{PAGE_W}" height="{PAGE_H}" xmlns="http://www.w3.org/2000/svg">'
 f'{DEFS}<rect width="100%" height="100%" fill="{C["bg"]}"/>'
 f'{body}<rect width="100%" height="100%" filter="url(#gr)" opacity="0.035"/>'
 f'</svg></body></html>')

W = PAGE_W
COVER = SVG(f'''
<!-- Full-page vertical wash -->
<rect width="100%" height="100%" fill="url(#gv)"/>
<!-- Soft organic top band -->
<path d="M0,0 L{W},0 L{W},240 Q{W*0.8},252 {W*0.6},238
 Q{W*0.35},222 {W*0.15},242 Q{W*0.05},248 0,245 Z" fill="url(#gw)"/>
<!-- Gentle secondary ribbon -->
<path d="M0,242 Q{W*0.1},248 {W*0.2},238 Q{W*0.4},220 {W*0.65},236
 Q{W*0.85},250 {W},234 L{W},258 Q{W*0.8},270 {W*0.55},258
 Q{W*0.3},244 {W*0.1},262 Q{W*0.03},268 0,264 Z" fill="url(#gc)"/>
<!-- Organic bottom bloom -->
<path d="M0,920 Q{W*0.25},908 {W*0.5},918 Q{W*0.75},928 {W},912
 L{W},{PAGE_H} L0,{PAGE_H} Z" fill="url(#gw)"/>
<!-- Soft accent at very bottom -->
<path d="M0,980 Q{W*0.3},970 {W*0.6},978 Q{W*0.85},986 {W},974
 L{W},{PAGE_H} L0,{PAGE_H} Z" fill="url(#gc)" opacity="0.5"/>
<!-- Delicate gold accent -->
<line x1="60" y1="880" x2="280" y2="880" stroke="{C['gold']}"
 stroke-width="0.8" opacity="0.35"/>
<circle cx="286" cy="880" r="2" fill="{C['gold']}" opacity="0.30"/>''')

BODY = SVG(f'''
<!-- Whisper of warmth at top -->
<rect x="0" y="0" width="{W}" height="6" fill="url(#gw)" opacity="0.55"/>
<!-- Whisper of cool at bottom -->
<rect x="0" y="{PAGE_H-5}" width="{W}" height="5" fill="url(#gc)" opacity="0.45"/>
<!-- Full page vertical wash barely visible -->
<rect width="100%" height="100%" fill="url(#gv)" opacity="0.45"/>''')

BACK = SVG(f'''
<rect width="100%" height="100%" fill="url(#gv)"/>
<!-- Top soft band -->
<path d="M0,0 L{W},0 L{W},290 Q{W*0.75},306 {W*0.5},290
 Q{W*0.25},274 0,292 Z" fill="url(#gw)"/>
<!-- Bottom band -->
<path d="M0,800 Q{W*0.3},790 {W*0.55},798 Q{W*0.8},808 {W},794
 L{W},{PAGE_H} L0,{PAGE_H} Z" fill="url(#gc)"/>
<!-- Centered accent line -->
<line x1="{W//2-70}" y1="750" x2="{W//2+70}" y2="750"
 stroke="{C['gold']}" stroke-width="0.6" opacity="0.22"/>
<circle cx="{W//2}" cy="750" r="1.2" fill="{C['gold']}" opacity="0.20"/>''')

def _render(tpl, out):
    from playwright.sync_api import sync_playwright
    os.makedirs(out, exist_ok=True)
    pairs = list(tpl.items())
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(viewport={'width': PAGE_W, 'height': PAGE_H}, device_scale_factor=2)
        for n, h in pairs:
            pg.set_content(h)
            pg.screenshot(path=os.path.join(out, n), type='png')
            print(n)
        b.close()

if __name__=='__main__':
    out=sys.argv[1] if len(sys.argv)>1 else '/mnt/agents/output/bg'
    _render({'cover_bg.png':COVER,'backcover_bg.png':BACK,'body_bg.png':BODY},out)
    print("Done - Soft Berry Bloom backgrounds")
