# Patching GUI Style Notes

Design rules and lessons learned while building the Patching GUI theme
system. Apply these when authoring or revising a theme in `Themes.psd1`,
or when writing a per-theme override in `ThemeOverrides/`. Covers
contrast, readability, color discipline, and gradient use.

## Architecture

- Canonical palettes live in `Themes.psd1` (single source of truth).
  Every theme provides the required 10 tokens + 3 meta fields, plus any
  optional richness fields listed in the schema header.
- Production GUI (`Invoke-PatchGUI.ps1`) builds its XAML dynamically
  from the resolved theme (preference file > env var > default).
- Heavily-customized themes that break the standard XAML structure
  (Consolas fonts, tactical brand bars, etc.) live in
  `ThemeOverrides/<ThemeKey>.ps1`. Production delegates to the override
  file if it exists.
- Gallery (`Invoke-PatchGUI-Gallery.ps1`) shows all themes as clickable
  cards; clicking one saves the selection and relaunches production.
- Shared logic used by more than one entry point lives in
  `Invoke-PatchGUI.Shared.ps1`. Every GUI that matters dot-sources it.

## The shared module rule

Any helper used by more than one GUI entry point lives in
`Invoke-PatchGUI.Shared.ps1`. XAML and named-control code stays local
to each entry point. This came out of the first CyberCommand override:
the override was a fork of the canonical GUI, and every fix to the
canonical had to be copy-pasted in or the override silently drifted.

What belongs in the shared file:

- Preferences I/O (`Get-PatchPreferences` / `Set-PatchPreferences`)
- Main-Switch discovery (`Get-MainSwitchNames` / `Get-MainSwitchListPaths`)
- DryRun mock data (`New-MockPatchResults`)
- Display-row flattening (`ConvertTo-DisplayRow`)
- Anything else that is pure logic and would be copy-pasted into a
  second GUI file

What does NOT belong:

- XAML strings (every entry point has its own visual identity)
- Event handlers tied to named WPF controls
- Functions that read or write `$script:mode` or other entry-point state
- Anything that imports the palette -- theme resolution is a
  production-GUI concern, not a shared one

### Why this works

The shared file lives at `GUI/`. Callers dot-source it from different
depths (`GUI/`, `GUI/ThemeOverrides/`). Because `$PSScriptRoot` inside a
function resolves to the file where the function was defined, helpers
in the shared file always see `GUI/` as their anchor. `Main-Switch.ps1`
is `..\..\Main-Switch.ps1` from there, regardless of which caller
dot-sourced. One canonical path, no per-caller path math.

Top-level code in the shared file runs in the caller's script scope on
dot-source, so setting `$script:PrefsDir` / `$script:PrefsPath` at the
top of the shared file populates those vars in every caller without
each caller having to repeat the two lines.

### When to add a new helper

If you are about to copy-paste a helper from the canonical GUI into a
second entry point, stop -- move it into the shared file instead.
Single-entry-point helpers stay local until the second caller appears.

## The SubText Trap

`SubText` was originally used for *three different roles* at *three
different required brightnesses*:

1. DataGrid column header text (needs to be readable from across the room)
2. Two-tone title secondary word (e.g. "Invoke" in "Invoke-Patch")
3. Footer / subtitle / hint text ("Patch Remediation", "min" labels)

A single `SubText` token cannot serve all three -- column headers need
to be noticeably brighter than true hint text. Split into two tokens:

- `HeaderText` - column headers and other near-body text. Target ~`#D4D4D4`
  on dark backgrounds (or the equivalent alpha on colored themes).
- `SubText` - true secondary / hint text. Target ~`#9C9C9C` minimum.
  Anything dimmer than that becomes unreadable at a glance.

Never let column headers fall below `#BEBEBE` on a dark theme. During
the contrast pass, Version / Compliant / New Version / Exit Code /
Comment were specifically flagged as hard to read when headers were at
`#787878`.

## Border Definition

Panel borders at `#333333` on a `#121212` background are *there* but do
not feel *defined* -- the panels look like they float into the void.

Minimum separation rule: **panel border lightness should be at least
0x18 (24 decimal) above the panel background's lightness.** On a
`#1C1C1C` surface, `#4A4A4A` works; `#333333` is too close.

This also applies to:
- ComboBox / TextBox borders (including the tiny timeout spinners)
- Checkbox empty-state box borders
- DataGrid column header bottom/side borders
- Button borders when the button uses a dark "ghost" style

### Exception: panel outlines can stay softer than input borders

Bumping *everything* to the same higher-contrast border made the big
card outlines fight the content inside them. The preferred split is:

- **Panel outlines** (the Header / Configuration / Progress / Results
  cards): can stay at the softer tone (~0x14 above Surface). They're
  big shapes; a faint outline is enough to define them. Themes can
  opt into this via the optional `PanelEdge` token.
- **Input / interior elements** (ComboBox, TextBox, CheckBox, scrollbar
  thumb, progress track, Cancel "ghost" button): must be at the higher
  contrast tone (~0x20-0x28 above Surface). These are small controls
  inside the cards; if their borders are soft, they visually disappear.

In the production XAML, the distinguishing pattern for panel outlines
is `BorderBrush="{StaticResource PanelEdge}"` on the four top-level
`<Border>` elements that wrap each section. The `Border` token itself
governs input ControlTemplates and should stay bright.

## Two-Tone Titles

When you split the title into a "light weight" + "bold" pair (e.g.
"Invoke" + "Patch" under `TitleStyle = 'split'`), the light side
should still be *readable*, not a whisper. `#787878` is too dim.
Target ~`#BEBEBE` so the eye sees a word, not noise. This is the
`TitleDimColor` optional token.

Same for the subtitle beneath -- push to `#9C9C9C` at minimum.

## AlternatingRowBackground

`#202020` against `#1C1C1C` is a 4-point difference and basically
invisible. Bump to at least `#232323` (7 points) so the stripes actually
register.

Hue-tinted panels make "neutral" grays shift toward the complementary
hue by simultaneous contrast -- a neutral gray alternating row on a
green-tinted panel can appear faintly purple. Tune row bg against the
theme's actual Surface, not against a neutral mental model.

## Progress Bar Track

Track background should match the panel border brightness, not the
panel background. A `#333333` track on a `#1C1C1C` panel looks like the
panel is empty; `#4A4A4A` reads as an actual track with something to
fill.

## Color Discipline and Gradients

Captured while reworking the Tokyo Night theme (cyan + pink + purple).
These apply to any theme with two or more accent colors.

### The core insight

Two-color themes have a discipline problem. Handing one color to some
elements and the other color to others ("Configuration" in cyan,
"Results" in pink) reads as tacky, not stylish. The fix is not "pick
better color assignments" -- it is "stop splitting structural duties
across colors." Pick ONE color for structural work and reserve the
second color for accent / alert. Let both colors coexist through
gradients on decorative surfaces.

### Gradient-friendly surfaces (these POP)

Use gradients on elements that are already "hero" surfaces -- the
user's eye is looking at them as individual objects, so weaving colors
there adds character.

- Accent pipe / title bar (`PipeStops`, `TitleStops`)
- Toggle track on-state (`ToggleOnStops`) -- also fixes the "solid red
  = alarm" problem
- Primary action button (`RunStops`)
- Progress bar fill (`ProgressStops`) -- pink -> cyan reads as "urgent
  start -> calm done"; matches the emotional journey of a long run
- Window background (`WindowBgGradient`) -- extremely subtle vertical
  darkening toward the bottom

### Gradient direction matters

- **Horizontal** (left-to-right) works for wide decorative bars: toggle
  tracks, button backgrounds, progress fills. Any time the element is
  much wider than it is tall.
- **Vertical** (top-to-bottom) is the right choice for text. Horizontal
  gradients on individual letters land each letter in a narrow hue band
  so the word reads like it's been stepped on. Vertical lets every
  letter traverse the full palette and feels intentional.
- **Elements adjacent to a title must match the title's gradient
  direction.** A pipe / mark / accent glyph next to the title should
  use vertical gradient (because the title does). Mixed directions
  side-by-side look chaotic even when each in isolation looks fine.
- Window-background gradients should be vertical (or a soft radial)
  and extremely subtle -- a barely-perceptible darkening toward the
  bottom is plenty.

### Gradient-hostile surfaces (these look tacky)

Don't use gradients on utility strokes or small-text surfaces. The
gradient becomes noise at small sizes.

- Column header underlines -- purple `#BB9AF7` was tested on TokyoNight
  and immediately flagged as fighting the design. Use the `Border`
  token instead.
- Panel borders / dividers
- Section headers and form labels (gradient text at 13-14pt is muddy)
- TextBox / ComboBox borders (reads as a broken focus state)
- Utility buttons (Default Sort, Export CSV) -- they should be quiet
- Checkbox tick marks (too small to benefit)

### Same-tier elements must share ONE color

The tackiness comes from mismatched peers, not from the presence of
two colors in the window.

- All **section headers** (Configuration, Results, Progress): same color
- All **DataGrid column headers**: same color as each other (often but
  not always the same as section headers)
- All **utility buttons** in a row: same color (neutral is usually best)
- All **form field labels**: same color

Breaking this rule even once makes the whole theme feel inconsistent.

### Button role -> color mapping

- **Primary action (Run)**: the celebration gradient via `RunStops`,
  or solid `Green` if the theme doesn't have a gradient identity.
  Ghost variant (deep-tinted bg + accent text + accent border) is
  available via `RunStyle='ghost'` + `RunGhostBg='#hex'`.
- **Destructive / stop (Cancel)**: default is the neutral ghost
  (`Overlay` bg, `SubText` fg, `Border` stroke). Themes with a
  strong alert color can opt into the accent ghost via `CancelBg`,
  `CancelFg`, `CancelBorder`.
- **Utility (Default Sort, Export CSV, Browse)**: neutral -- `Text`
  color on `Overlay` background with `Border` stroke. They should NOT
  compete with Run for attention.

In TokyoNight, Run gets the cyan->purple->pink gradient and Cancel
gets the pink ghost. That's the reference pattern for a two-color
theme with a saturated alert color.

### Toggle on-state colors are emotional

A solid red or pink "on" state reads as "I just turned on something
dangerous." It matters that the user feels okay about the default
position (Patch = on).

The fix: let the on-state track use a **gradient from the off-state
color to the on-state color** via `ToggleOnStops`. The eye reads it as
"slide between two modes" instead of "alarm." TokyoNight's toggle
track goes cyan -> purple -> pink left to right; with the knob on the
right, the user sees mostly gradient, not pink, and it feels like
transition rather than alert.

For themes with a distinct toggle identity (e.g. CobaltSlate's blue
toggle vs white Run button), use `ToggleOnBg` / `ToggleOnBorder` to
set a solid color that's different from the Run accent.

### Structural vs decorative discipline

Two mental buckets, keep them separate:

- **Structural**: title, section headers, column headers, labels,
  dividers, body text. Use ONE color family (usually the primary
  accent, e.g. cyan in TokyoNight). Monochrome within this bucket.
- **Decorative**: accent pipe, toggle, primary action button, progress
  bar, row-selection highlight, hero flourishes. This is where you can
  have fun with gradients and multi-color. Make it count.

The moment structural elements start picking up the decorative
palette's variety, the GUI feels cluttered.

### Adjacent gradients must not compete on the same endpoint color

When a gradient text element and a gradient pipe (or bar, or shape)
sit next to each other, the eye catches any misalignment in where a
strong color lands -- especially pink, red, or any saturated warm
end-stop. Two elements can't share a vivid endpoint color unless
their geometry forces that color to land at the exact same Y (or X)
coordinate. Since fonts and decorative bars rarely share bounds, the
easiest fix is to **drop the competing end-stop from one of them.**

Pattern used in TokyoNight: the pipe runs `cyan -> purple -> pink` top
to bottom; the title runs `cyan -> purple` only (no pink). Pink lives
in the decorative layer exclusively, and the title's gradient still
feels like "part of the palette" without fighting the pipe's pink
horizon.

Apply this whenever you put gradient text near a gradient shape.

### Reference themes

- **TokyoNight** is the canonical multi-color gradient example. Read
  its entry in `Themes.psd1` to see how `PipeStops` / `TitleStops` /
  `RunStops` / `ProgressStops` / `ToggleOnStops` / `WindowBgGradient`
  coexist without fighting each other.
- **CobaltSlate** is the monochromatic baseline -- no gradients, solid
  accent on everything, distinct blue toggle vs white Run button.
- **DarkForest** is the inspiration-driven palette example (see below).

## Palette Tuning: Start from Inspiration, Not from Hex Codes

When building or revising a theme's palette, iterating directly on hex
values leads to "palette decoherence" -- each tweak is made against an
abstract mental model rather than a concrete reference, so the color
family drifts with every change.

The pattern that works:

1. Find an inspiration image that captures the target mood. Art sites,
   dashboard gallery screenshots, photography all work. One image is
   enough; the important thing is it exists as a fixed reference.
2. Extract temperature cues from it before touching any hex: is the
   base warm or cool? Which greens (or blues, or reds) are present --
   true leaf or chartreuse, sage or mint, pine or forest? What's the
   warm accent doing -- gold, copper, orange, amber?
3. Propose replacements in each direction and render them side-by-side
   in the actual theme context (same panel background, next to the
   other palette colors they will coexist with). An in-context
   palette-picker WPF window works well for this -- render each
   candidate with the pipe + Run button + section-header text against
   the rest of the palette so harmony is visible at a glance.
4. Pick by eye, not by description. Descriptions lie; rendered
   swatches do not.

Reference: DarkForest was iterated blind on the action-green and
landed on `#8FBC4D` chartreuse that did not match any green in the
inspiration image. After comparing against a layered-dashboard /
tree-infographic photo, the palette shifted cool across the board
(Bg navy-ward, Green to true leaf, Amber to copper) and the whole
theme snapped into coherence.

## Light Companions

Pairing a dark theme with a light variant (e.g. CobaltSlateDay pairs
with CobaltSlate) is mostly mechanical token-flipping, but there are
gotchas that don't fall out of a naive bulk replace:

- **`AccentText` almost always needs setting on light themes.** Dark
  themes default `AccentText = Bg` (dark text on the bright Run
  button), which is fine. On a light theme, `Bg` is near-white, which
  fails contrast against any saturated accent -- set `AccentText =
  '#FFFFFF'` or the theme's own Text token.
- **Accent-typed optionals (`CancelBg`, `RunGhostBg`, `ToggleOnBg`,
  etc.) usually need light-theme-specific values.** A dark theme's
  `CancelBg = '#4A2538'` pale-rose tint looks invisible on white.
  Often you want a paler tint (e.g. `#F2E8EB`) instead.
- **SubText floor is different.** `#9C9C9C` on near-black reads fine;
  on near-white, you need something darker (~`#6A7590`) for the same
  perceived dimming.
- **Test in both Patch and Version mode** -- the subtitle, mode
  slider labels, and hidden-column visibility flips all surface token
  combinations the dark theme never exercises.
- **AccentText on the Run button is the #1 thing that breaks.** Run's
  text color defaults to `Bg`; if Bg is near-white, the label vanishes
  on any saturated accent. Always override `AccentText` explicitly
  for light themes.

## Theme Checklist

Before releasing a new theme:

- [ ] Column headers readable from 3 feet away
- [ ] Panel borders clearly separate panels from window background
- [ ] Input field borders (TextBox, ComboBox, timeout spinners) visible
- [ ] Checkbox empty-state boxes visible (not blending into panel)
- [ ] Two-tone title (if `TitleStyle='split'`): both halves legible,
      not just the bold half
- [ ] Subtitle text legible (`Patch Remediation` / `Version Audit`
      or the theme's `SubtitlePatch` / `SubtitleVersion` override)
- [ ] "min" unit labels and footer result count legible
- [ ] Scrollbar thumb visible against its track
- [ ] Alternating row stripes actually visible (>= 7 points separation
      from Surface)
- [ ] Cancel button (neutral-ghost by default) still has enough
      contrast that users can find it
- [ ] Run button text readable on its accent color (check `AccentText`
      on light themes)
- [ ] If the theme defines a `ThemeOverrides/<Key>.ps1` file, the
      Gallery preview approximation still reads as recognizably the
      same theme

## Recommended Palette Token Set

For any new theme, define at minimum:

```
Bg          - window background
Surface     - panel background
Overlay     - input background (textbox, combobox, button ghost)
Border      - panel/input borders  (>= Surface + 0x18)
Hover       - interactive hover state
Text        - body text  (>= #E0 equivalent)
HeaderText  - column headers, two-tone secondary (~#D4)
SubText     - hints, subtitles (>= #9C on dark themes)
Blue        - cool accent / info
Green       - primary action / success
Red         - destructive / alert
```

Themed accents layer on top of this grayscale ladder; the grayscale
ladder is what guarantees the GUI is readable when the accents are
sparse. See the schema header in `Themes.psd1` for the full list of
optional richness fields (gradients, ghost styles, alternate toggle
colors, custom subtitles, etc.).
