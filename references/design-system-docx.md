# DOCX Design System (Default)

## Typography
- Body: Calibri 11pt
- H1: 18pt bold
- H2: 14pt bold
- H3: 12pt bold
- Note text: 10–11pt italic, muted

## Primary Token Set
- Primary: 1A2332
- Secondary: 2D8B8B
- Body text: 1F2937
- Muted: 475569
- Border: CBD5E1
- Zebra row: F4F8FA

## Brand Override (User Brand)
- Brand primary color: `ABDF00` (lime green)
- Brand visual direction: green-led system with flexible dark/light balance
- Recommended gradient pairings:
  - `ABDF00` → `D8FF72` (light lime)
  - `7FB800` → `ABDF00` (deeper to bright lime)
- Core rule: green stays the anchor, but tone can shift by document context.

### Adaptive Palette Modes
1. **Corporate Green-Dark** (formal docs)
   - Primary: `2F4F2F`
   - Accent: `ABDF00`
   - Support: `EAF7C8`, `F5F7FA`
   - Text: `1F2937`

2. **Balanced Green-Light** (general proposals/reports)
   - Primary: `ABDF00`
   - Support dark: `3E5C2A`
   - Support light: `F7FFE6`, `ECFDF3`
   - Text: `1F2937`

3. **Green + Pastel Mix** (creative/social marketing docs)
   - Green anchor: `ABDF00`
   - Pastel supports: `FFD6E7`, `CDEBFF`, `FFE9B3`, `E6D9FF`
   - Use pastel only as secondary highlights; avoid overpowering green identity.

### Usage rules
- Use green as the dominant visual identity in headers/key accents.
- Flex between dark and light tones based on readability and tone.
- Pastel is optional for variety; keep one primary green anchor on every page.
- Keep body text dark (`1F2937`/`2D2D2D`) for readability.
- Do not use strong gradients behind dense paragraphs.
- Preserve contrast (white or very dark text depending on background intensity).

## Alternate Client Style Pack (when asked)
- H1: 2A7A78
- H2: 8B6B4A
- H3: 1F4D78
- Body: 2D2D2D / 333333
- Accents (limited): 76C893, FF9E8A, 5B4FA0

## Table Rules
- Set table width with DXA.
- Set columnWidths and each cell width.
- Header row: dark fill + white bold text + centered.
- Body: zebra rows, left-align descriptive text, center short/numeric cells.
- Always add cell margins/padding.

## Prohibitions
- No random fonts/colors.
- No leaked internal guidance text (e.g., "Theme: ...") unless user explicitly requests it in output.
