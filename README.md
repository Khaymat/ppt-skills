# ppt-skills

Skills for AI agents that build PowerPoint decks in a fixed visual style. Each style lives in its own folder with a
`SKILL.md`, a Python toolkit, and an example, so you can install one style without the others.

## Styles

Only one style exists for now.

| Folder | Style | Preview |
|---|---|---|
| [`etching-ppt/`](etching-ppt/) | Monochrome pen-and-ink: warm paper, EB Garamond, spaced-caps titles, Roman numerals | [demo PDF](etching-ppt/examples/demo.pdf), [contact sheet](etching-ppt/examples/preview/contact_sheet.png) |

More styles will be added as sibling folders.

## Quick start

```bash
cd etching-ppt
pip install -r requirements.txt
python scripts/build_deck.py examples/demo_spec.json -o demo.pptx
```

Install instructions for agents (Claude Code and others) are in each style's own README.

## License

MIT, see `LICENSE`.
