#!/usr/bin/env python3
"""Validate that generated PPT text exactly preserves the locked source inventory."""
import json, sys
from collections import Counter
from pathlib import Path


def norm(s: str) -> str:
    # Preserve wording and punctuation; normalize only line endings and outer whitespace.
    return s.replace('\r\n', '\n').replace('\r', '\n').strip()


def fail(msg: str) -> None:
    print(f"FAIL: {msg}", file=sys.stderr)
    raise SystemExit(1)

if len(sys.argv) != 3:
    fail("usage: validate_content_lock.py page-plan.json generated-texts.json")
plan = json.loads(Path(sys.argv[1]).read_text(encoding='utf-8'))
gen = json.loads(Path(sys.argv[2]).read_text(encoding='utf-8'))
if plan.get('allow_text_rewrite', False):
    fail('allow_text_rewrite must be false for image-to-editable restoration')
source = [norm(x['content']) for x in plan.get('source_text_inventory', []) if norm(x.get('content',''))]
generated = [norm(x) for x in gen.get('texts', []) if norm(x)]
if not source:
    fail('source_text_inventory is empty')
cs, cg = Counter(source), Counter(generated)
missing = list((cs-cg).elements())
extra = list((cg-cs).elements())
if missing: fail(f'missing or altered source text: {missing}')
if extra: fail(f'new or duplicated text not in source: {extra}')
print('PASS: exact text inventory preserved')
