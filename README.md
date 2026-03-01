# Office Design Toolkit Skill

A production-ready Office skill system for AI agents to create **high-quality DOCX/XLSX/PPTX/PDF** outputs with a strict workflow:

1. **Content first**
2. **Structure second**
3. **Design third**
4. **QA before delivery**

This repository packages the Office workflow and design references used to produce professional client-facing documents consistently.

---

## Why this project exists

Most agent-generated documents fail because they style too early, skip structure, or output visually inconsistent files.

This toolkit fixes that by enforcing:

- deterministic document workflow
- reusable design system tokens
- QA gates before final delivery
- no leakage of internal style-guidance text into user-facing content

---

## Key capabilities

- DOCX-focused design system and delivery standard
- Modular references for workflow policy, design tokens, and QA
- Brand-aware styling rules (including primary brand color systems)
- Support for alternate style packs from approved sample docs
- Explicit operating mode for fast tasks without skipping phase order

---

## Skill structure

```text
office-design-toolkit/
├── SKILL.md
├── README.md
└── references/
    ├── workflow-policy.md
    ├── design-system-docx.md
    └── qa-checklists.md
```

### `SKILL.md`
Core skill contract and execution standard.

### `references/workflow-policy.md`
Mandatory phase sequence and operating policy.

### `references/design-system-docx.md`
Typography, color tokens, table rules, brand overrides, adaptive palette modes.

### `references/qa-checklists.md`
Content / structure / design / technical quality gates.

---

## Workflow standard (mandatory)

For Office outputs, always run in this order:

1. **Content Draft** — objective, audience, key facts, actions
2. **Structure Design** — heading hierarchy, information architecture, table/list decisions
3. **Visual Design & Polish** — style tokens, spacing, consistency, readability
4. **QA & Delivery** — content + structure + design + technical checks

---

## Design philosophy

- Green-led brand identity can be preserved while adapting dark/light tone by document context.
- Pastel accent colors may be used selectively for creative documents while maintaining a green anchor.
- Readability and hierarchy are prioritized over decoration.

---

## Installation

### Option A — Use as local skill folder
Place this folder inside your skills directory:

```bash
~/.agents/skills/office-design-toolkit
```

### Option B — Clone from GitHub

```bash
git clone https://github.com/<your-username>/office-design-toolkit.git
```

Then connect it to your agent's skill loading path.

---

## Usage notes

- Do not insert internal style instructions in final documents unless explicitly requested by the user.
- Keep phase order even when user asks for faster turnaround.
- Apply brand tokens before final styling pass.

---

## Roadmap

- Add PPTX and XLSX visual QA presets
- Add optional template assets for common business document types
- Add packaging script for one-command skill distribution

---

## License

MIT (or your preferred license; update as needed).

---

## Credits

Created with product direction and quality standards by:

**Lê Huy Đức Anh — Founder, Vidtory.ai**

