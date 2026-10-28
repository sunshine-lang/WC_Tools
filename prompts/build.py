#!/usr/bin/env python3
"""Assemble prompt scenes from base.md + scenes/*.md into dist/.

Edit base.md or scenes/*.md, then run:
    python3 prompts/build.py
Do not edit dist/ by hand.
"""
from __future__ import annotations

import re
import argparse
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent))
from core_tools.language import language_profiles, prompt_locations


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


def main(argv=None) -> int:
    profiles = language_profiles()
    parser = argparse.ArgumentParser(description='生成各场景的完整翻译提示词')
    parser.add_argument('--lang', choices=tuple(profiles),
                        default=next(key for key, profile in profiles.items() if profile['default']))
    args = parser.parse_args(argv)
    source, destination = prompt_locations(profiles[args.lang], ROOT)
    base = (source / 'base.md').read_text('utf-8')
    destination.mkdir(exist_ok=True, parents=True)
    scenes = sorted((source / 'scenes').glob('*.md'))
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
        if Path(output_name).name != output_name or not output_name.endswith('.md'):
            raise ValueError('OUTPUT 必须是当前输出目录下的 Markdown 文件名')
        output = destination / output_name
        output.write_text(text.rstrip() + '\n', 'utf-8')
        print(f'{scene_path.name} -> {output.relative_to(ROOT.resolve())} ({len(text)} chars)')

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
