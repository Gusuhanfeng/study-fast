# Contributing

Contributions should preserve the skill's evidence boundaries, privacy guarantees, and selective output policy.

## Development process

1. Keep generated DOCX files, private task configurations, and render outputs outside the repository.
2. Make focused changes and update the relevant reference instead of duplicating rules in `SKILL.md`.
3. Add tests for observable behavior and failure boundaries.
4. Run `python scripts/self_check.py` and `python -m unittest discover -s tests -v`.
5. For layout changes, generate a sample pack, audit every DOCX, render every page, and inspect the PNGs.

Do not add real resumes, employer documents, meeting links, credentials, customer data, unpublished research data, or generated private artifacts as fixtures. Use synthetic examples and reserved domains such as `example.com`.

Bug reports should include the operating system, Python version, selected template, command, sanitized configuration, observed result, and expected result. Remove all private information before posting.
