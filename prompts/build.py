#!/usr/bin/env python3
"""Assemble prompt scenes from base.md + scenes/*.md into dist/.

Edit base.md or scenes/*.md, then run:
    python3 prompts/build.py
Do not edit dist/ by hand.
"""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BASE = ROOT / 'base.md'
SCENES = ROOT / 'scenes'
DIST = ROOT / 'dist'


def metadata(text: str, name: str) -> str:
    match = re.search(rf'<!--\s*{name}:\s*(.*?)\s*-->', text)
    if not match:
        raise ValueError(f'missing metadata: {name}')
    return match.group(1).strip()


def block(text: str, name: str) -> str:
    match = re.search(rf'<!--\s*{name}\s*-->\n(.*?)(?=\n<!--\s*[A-Z_]+\s*-->|\Z)', text, re.S)
    if not match:
        raise ValueError(f'missing block: {name}')
    return match.group(1).strip()


def main() -> int:
    base = BASE.read_text('utf-8')
    DIST.mkdir(exist_ok=True)
    scenes = sorted(SCENES.glob('*.md'))
    if not scenes:
        raise SystemExit('no scenes found')

    for scene_path in scenes:
        scene = scene_path.read_text('utf-8')
        values = {
            'SCENE_TITLE': metadata(scene, 'TITLE'),
            'CATEGORY': metadata(scene, 'CATEGORY'),
            'SCENE_FOCUS': block(scene, 'FOCUS'),
            'SCENE_TASK': block(scene, 'TASK'),
            'SCENE_RULES': block(scene, 'RULES'),
            'SCENE_EXAMPLES': block(scene, 'EXAMPLES'),
            'SCENE_CHECKS': block(scene, 'CHECKS'),
        }
        text = base
        for key, value in values.items():
            text = text.replace('{{' + key + '}}', value)

        missing = sorted(set(re.findall(r'\{\{[A-Z_]+\}\}', text)))
        if missing:
            raise ValueError(f'{scene_path.name}: unresolved placeholders: {missing}')

        output_name = metadata(scene, 'OUTPUT')
        output = DIST / output_name
        output.write_text(text.rstrip() + '\n', 'utf-8')
        print(f'{scene_path.name} -> dist/{output_name} ({len(text)} chars)')

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
