#!/usr/bin/env python3
"""Audit PowerPoint text runs for native gradient or solid fill.

Usage:
  python audit_text_fill.py deck.pptx
  python audit_text_fill.py deck.pptx expectations.json

expectations.json example:
[
  {"contains": "器件更高效", "fill": "gradient"},
  {"contains": "系统更好协同", "fill": "gradient"}
]
"""
from __future__ import annotations
import json, sys, zipfile
from pathlib import Path
from lxml import etree

A='http://schemas.openxmlformats.org/drawingml/2006/main'
NS={'a':A}

def run_fill(r):
    rpr=r.find('a:rPr',NS)
    if rpr is None: return 'inherited'
    if rpr.find('a:gradFill',NS) is not None: return 'gradient'
    if rpr.find('a:solidFill',NS) is not None: return 'solid'
    if rpr.find('a:noFill',NS) is not None: return 'none'
    return 'inherited'

def main():
    if len(sys.argv)<2:
        raise SystemExit('usage: audit_text_fill.py deck.pptx [expectations.json]')
    deck=Path(sys.argv[1]); expected=[]
    if len(sys.argv)>2: expected=json.loads(Path(sys.argv[2]).read_text(encoding='utf-8'))
    rows=[]
    with zipfile.ZipFile(deck) as z:
        for name in sorted(n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')):
            root=etree.fromstring(z.read(name))
            for r in root.findall('.//a:r',NS):
                t=r.find('a:t',NS)
                if t is not None and (t.text or '').strip():
                    rows.append({'slide':name,'text':t.text,'fill':run_fill(r)})
    for row in rows: print(f"{row['slide']}\t{row['fill']}\t{row['text']}")
    failed=[]
    for e in expected:
        hits=[r for r in rows if e['contains'] in r['text']]
        if not hits or not any(r['fill']==e['fill'] for r in hits): failed.append(e)
    if failed:
        print('FAILED expectations:', json.dumps(failed,ensure_ascii=False), file=sys.stderr)
        return 2
    return 0
if __name__=='__main__': raise SystemExit(main())
