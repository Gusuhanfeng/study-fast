#!/usr/bin/env python3
"""Build the opt-in offline quiz; independent of the Word pack contract."""
from __future__ import annotations

import argparse
import html
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def validate(data):
    if not isinstance(data, dict) or set(data) != {'bank_id', 'title', 'questions'}:
        raise ValueError('Expected bank_id, title and questions only.')
    if not isinstance(data['bank_id'], str) or not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_-]{0,79}', data['bank_id']):
        raise ValueError('Invalid bank_id.')
    def nonempty(value):
        return isinstance(value, str) and bool(value.strip())
    if not nonempty(data['title']):
        raise ValueError('title must not be blank.')
    if not isinstance(data['questions'], list) or not data['questions']:
        raise ValueError('questions must be a nonempty array.')
    seen = set()
    for q in data['questions']:
        if not isinstance(q, dict) or not {'id', 'category', 'question', 'answer'} <= set(q) or set(q) - {'id', 'category', 'question', 'answer', 'criteria', 'title', 'learning'}:
            raise ValueError('Invalid question fields.')
        if not all(nonempty(q[k]) for k in ('id', 'category', 'question', 'answer')):
            raise ValueError('Question fields must not be blank.')
        if q['id'] in seen:
            raise ValueError('Duplicate question ID: ' + q['id'])
        seen.add(q['id'])
        if 'criteria' in q and (not isinstance(q['criteria'], list) or not all(nonempty(x) for x in q['criteria'])):
            raise ValueError('criteria must contain nonblank strings.')
        if 'title' in q and not nonempty(q['title']):
            raise ValueError('title must be a nonblank string.')
        if 'learning' in q:
            learning = q['learning']
            fields = {'concept', 'method', 'explanation', 'example', 'pitfalls'}
            if not isinstance(learning, dict) or not learning or set(learning) - fields or not all(nonempty(x) for x in learning.values()):
                raise ValueError('learning must contain supported nonblank text fields.')
    return data


def render(data):
    validate(data)
    template = (ROOT / 'assets' / 'quiz.html').read_text(encoding='utf-8')
    payload = json.dumps(data, ensure_ascii=False).replace('<', '\\u003c').replace('>', '\\u003e').replace('&', '\\u0026')
    replacements = {'__QUIZ_TITLE__': html.escape(data['title']), '__QUIZ_DATA__': payload}
    # A single pass leaves token-shaped user text untouched.
    return re.sub(r'__QUIZ_TITLE__|__QUIZ_DATA__', lambda match: replacements[match.group()], template)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--config', type=Path, required=True)
    parser.add_argument('--output', type=Path)
    parser.add_argument('--check-only', action='store_true')
    parser.add_argument('--overwrite', action='store_true')
    args = parser.parse_args()
    try:
        data = validate(json.loads(args.config.read_text(encoding='utf-8-sig')))
        if not args.check_only:
            if args.output is None or args.output.suffix.lower() != '.html':
                raise ValueError('--output must name an HTML file.')
            if args.output.resolve() == args.config.resolve():
                raise ValueError('Output must differ from input.')
            if args.output.exists() and not args.overwrite:
                raise ValueError('Output exists; use --overwrite only after preserving user edits.')
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(render(data), encoding='utf-8')
        print(json.dumps({'valid': True, 'questions': len(data['questions']), 'output': str(args.output) if not args.check_only else None}, ensure_ascii=False))
    except (ValueError, OSError) as exc:
        parser.exit(2, str(exc) + '\n')


if __name__ == '__main__':
    main()
