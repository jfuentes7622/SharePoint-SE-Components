# GitHub Pages Single-Page Design Document (SPFx Dynamic Form)

Date: 2026-03-11

## Goal
Provide SharePoint developers with a clean, professional, bilingual (Chinese/English) single-page introduction that clearly communicates the project's value, key features, and quick-start steps, with a clear CTA/link entry point.

## Audience
- Primary: SharePoint / SPFx developers
- Secondary: tech leads, PMs, IT leads (quickly understand the capability boundaries)

## Scope (single page, simplified information)
Only 6 sections:
1. Project intro (Hero + CTA)
2. Key features (cards)
3. Quick start (steps)
4. Tech stack/compatibility (tags/brief table)
5. FAQ (2-3 items)
6. CTA/links

## Bilingual presentation strategy
- Within the same section, "Chinese primary text + English secondary text" side by side.
- Two columns on desktop, stacked on mobile.
- Chinese text slightly larger, English secondary text slightly smaller and de-emphasized in color.

## Visual direction (Fluent enterprise feel)
- Tone: clean, rational, light layering
- Primary color: deep blue (primary buttons/accents) + neutral gray
- Background: very light gray gradient or subtle noise texture, avoiding a stark empty white
- Cards: 12px rounded corners, light shadow

## Layout and typography
- 12-column grid, max width 1120px
- Section vertical spacing 56-72px
- Font: `Segoe UI` / `Microsoft YaHei` combination
- Font sizes:
  - Hero title 36-40px
  - Body text 15-16px
  - English secondary text 13-14px

## Content draft (Chinese/English side by side)

### 1) Project intro (Hero)
- Chinese title: SharePoint Dynamic Form Engine (SPFx 1.22)
- English title: SharePoint Dynamic Form Engine (SPFx 1.22)
- Chinese subtext: A visual form designer and runtime renderer for SharePoint lists.
- English subtext: A visual form designer and runtime renderer for SharePoint lists.
- CTA: `Get Started` (anchor to quick start) / `View on GitHub`

### 2) Key features (4 cards)
1. Visual Designer
2. Conditional Rules & Validation
3. Rich Field Types
4. SharePoint List Integration

### 3) Quick start (4 steps)
1. Install dependencies `npm install`
2. Start local dev `heft start`
3. Add the web part to a page
4. Pick a list and enter designer mode

### 4) Tech stack/compatibility
- SPFx 1.22
- React 17
- Fluent UI 8
- TypeScript 5.3
- Heft toolchain

### 5) FAQ (3 items)
- How does the designer work?
- Does it support attachments and advanced fields?
- How to handle PnP control style issues?

### 6) CTA/links
- GitHub Repo
- Design Spec document
- Issue / Feedback

## Illustrations and placeholders
- Use geometric abstract art or simplified UI wireframes as placeholders
- Can be replaced with real screenshots/animations later

## Implementation notes (to be implemented)
- Static single page `docs/index.html` + `docs/assets/*`
- GitHub Pages points to the `/docs` directory
- Pure static HTML/CSS/minimal JS, no build chain

## Success criteria
- Project positioning and core capabilities understood within 10 seconds
- Quick-start path completed within 30 seconds
- Bilingual reading feels natural, no switching cost
