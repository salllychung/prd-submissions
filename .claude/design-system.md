# Apple Design System — Reference Tokens

## Colors
| Token | Hex | Use |
|---|---|---|
| `--color-primary` | #0066cc | All CTAs, links, focus signals |
| `--color-primary-focus` | #0071e3 | Focus ring |
| `--color-primary-on-dark` | #2997ff | Links on dark tiles |
| `--color-canvas` | #ffffff | Default page canvas |
| `--color-parchment` | #f5f5f7 | Alternating light tiles, footer |
| `--color-ink` | #1d1d1f | All text on light surfaces |
| `--color-body-muted` | #cccccc | Secondary text on dark |
| `--color-ink-muted-80` | #333333 | Body on pearl surface |
| `--color-ink-muted-48` | #7a7a7a | Disabled / fine print |
| `--color-tile-1` | #272729 | Primary dark tile |
| `--color-tile-2` | #2a2a2c | Dark tile variant (above/below tile-1) |
| `--color-tile-3` | #252527 | Bottom of stack / video frames |
| `--color-surface-black` | #000000 | Nav bar, video bg, closing slides |
| `--color-hairline` | #e0e0e0 | Card borders, table lines |
| `--color-divider-soft` | #f0f0f0 | Soft border on secondary buttons |

## Typography
| Token | Size | Weight | Line-height | Letter-spacing |
|---|---|---|---|---|
| hero-display | 56px | 600 | 1.07 | -0.28px |
| display-lg | 40px | 600 | 1.10 | 0 |
| display-md | 34px | 600 | 1.47 | -0.374px |
| lead | 28px | 400 | 1.14 | 0.196px |
| lead-airy | 24px | 300 | 1.5 | 0 |
| tagline | 21px | 600 | 1.19 | 0.231px |
| body-strong | 17px | 600 | 1.24 | -0.374px |
| body | 17px | 400 | 1.47 | -0.374px |
| caption | 14px | 400 | 1.43 | -0.224px |
| button-large | 18px | 300 | 1.0 | 0 |
| button-utility | 14px | 400 | 1.29 | -0.224px |
| fine-print | 12px | 400 | 1.0 | -0.12px |
| nav-link | 12px | 400 | 1.0 | -0.12px |

Font stack (display): `"SF Pro Display", -apple-system, BlinkMacSystemFont, "Inter", system-ui, sans-serif`
Font stack (body/UI): `"SF Pro Text", -apple-system, BlinkMacSystemFont, "Inter", system-ui, sans-serif`

## Spacing
| Token | Value |
|---|---|
| xxs | 4px |
| xs | 8px |
| sm | 12px |
| md | 17px |
| lg | 24px |
| xl | 32px |
| xxl | 48px |
| section | 80px |

## Border Radius
| Token | Value | Use |
|---|---|---|
| none | 0px | Full-bleed tiles |
| xs | 5px | Subtle chips |
| sm | 8px | Dark utility buttons |
| md | 11px | Pearl button capsules |
| lg | 18px | Store/accessory cards |
| pill | 9999px | Primary CTAs, search input |
| full | 9999px/50% | Circular controls |

## Elevation
- **Flat** — no shadow, no border (tiles, nav, footer)
- **Soft hairline** — `1px rgba(0,0,0,0.08)` border (utility cards)
- **Backdrop blur** — `backdrop-filter: saturate(180%) blur(20px)` (sticky bars)
- **Product shadow** — `rgba(0,0,0,0.22) 3px 5px 30px 0` (product renders ONLY)

## Slide Layout Catalog
| Layout key | Background | Title size | Notes |
|---|---|---|---|
| title-light | #ffffff | 40px/600 | Center-aligned, two pill CTAs |
| title-dark | #272729 | 40px/600 | Center-aligned, white text |
| section | #272729 | 28px/600 | Large section number in #0066cc top-left |
| content | #f5f5f7 | 24px/600 | 60% text left / 40% image right |
| bullets | #ffffff | 24px/600 | Blue square bullets |
| chart-bar | #f5f5f7 | 24px/600 | Bar chart, editable data |
| chart-line | #ffffff | 24px/600 | Line chart, editable data |
| chart-pie | #272729 | 24px/600 white | Pie chart on dark, editable data |
| table | #ffffff | 24px/600 | Dark header row, alternating rows |
| quote | #272729 | 24px/400 white | 72px ❝ in #0066cc |
| two-column | #ffffff | 24px/600 | Left parchment / right dark split |
| closing | #000000 | 40px/600 white | Pill CTA #0066cc |

## Do / Don't
- ✅ Single accent color #0066cc for ALL interactive elements
- ✅ Negative letter-spacing on headlines (–0.28 to –0.374px)
- ✅ Body copy at 17px, not 16px
- ✅ `transform: scale(0.95)` for button press state
- ❌ No second accent color
- ❌ No shadows on cards/buttons/text
- ❌ No gradients as decorative backgrounds
- ❌ No weight 500 (ladder: 300/400/600/700)
- ❌ No rounded corners on full-bleed tiles
