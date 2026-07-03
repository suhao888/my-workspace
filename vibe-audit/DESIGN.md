# Douyin — Design

> **Hard-gate (BUILD flow only).** This file is the visual contract for
> the scene-project flow (`vibe build`, `vibe scene ...`, composition
> HTML, backdrop image-gen). Author it before authoring scene HTML; the
> Hyperframes `hyperframes` skill enforces it at composition time.
>
> **Single-asset requests (`vibe generate image|video|speech|...`) do
> NOT consult this file.** Run the generate command directly with the
> user's prompt. See AGENTS.md → "Route by the user's actual request".

Visual identity for **Douyin**, scaffolded from the **Velvet Standard** style (after Massimo Vignelli). Customise freely — this file is the single source of truth for every scene's palette, typography, and motion.

## Style

**Mood:** Premium, timeless · **Best for:** Luxury products, enterprise software, keynotes, investor decks

## Palette

- `#000000`
- `#ffffff`
- `#1a237e`

Black, white, ONE rich accent — deep navy (#1a237e) or gold (#c9a84c).

## Typography

Thin sans-serif, ALL CAPS, wide letter-spacing (0.15em+). Sequential reveals only.

## Composition

Generous negative space. Symmetrical, centered, architectural precision. Nothing busy.

## Motion

Slow, deliberate. Sequential reveals with long holds. No frantic motion.

**GSAP signature:** sine.inOut, power1 — nothing snaps, everything glides

## Transition

Cross-Warp Morph (elegant, organic flow between scenes)

## What NOT to do

- Tight letter-spacing — kills the premium register
- Bouncy or elastic easings — too playful
- Multiple elements arriving at once — break sequence

---

_Browse other named styles: `vibe scene list-styles`_
_This file was seeded by `vibe scene init --visual-style "Velvet Standard"`._
