# Land It Well, Dewi!

An anime comic about the Cloud Adoption Framework and the Well-Architected Framework.

**Read it:** https://ibranibeny.github.io/contoso-waf-caf-comic/

Contoso is moving to Azure, and Dewi has to lead it. Over 16 animated episodes she learns the **Cloud Adoption Framework** (how a whole organization adopts Azure: Strategy, Plan, Ready, Adopt, Govern, Secure, Manage) and the **Well-Architected Framework** (how to build each workload well: Reliability, Security, Cost Optimization, Operational Excellence, Performance Efficiency).

- **Animated story:** panels zoom and pan, characters enter and react, speech bubbles type themselves out. Click, tap or press → to advance.
- **Auto-play:** press **A** (or open `e00.html?autoplay=1`, add `&speed=1.5` to go faster). Episodes chain automatically.
- **17 animated architecture diagrams** drawn with the official Azure Architecture Icons: glowing packets travel each connection hop by hop, with a caption for every step. See them all on [diagrams.html](https://ibranibeny.github.io/contoso-waf-caf-comic/diagrams.html).
- **Checkpoint quiz** at the end of every episode, a [cheat sheet](https://ibranibeny.github.io/contoso-waf-caf-comic/cheatsheet.html) and a [sources page](https://ibranibeny.github.io/contoso-waf-caf-comic/sources.html).
- **Export to PowerPoint:** the **⬇ PowerPoint** button (or **P** in an episode) builds an animated `.pptx` in the browser for one episode, the whole comic or all 17 diagrams: one slide per beat with Morph transitions and typed speech, lesson cards, quizzes, editable diagrams with official Azure icons and packets that fly along motion paths, and speaker notes on every slide. Nothing is uploaded anywhere.
- **Reduced motion:** if your device asks for reduced motion, the comic starts still and offers **Play animations**; your choice is remembered, and **M** toggles it. **T** opens a full text version of every episode.

## Episodes (one continuous story, about 133 minutes)

| # | Episode | Topic | Duration |
|---|---------|-------|----------|
| Prologue | [The Server Room Is Full](https://ibranibeny.github.io/contoso-waf-caf-comic/e00.html) | Prologue | 6 min |
| EP 1 | [Why Are We Flying?](https://ibranibeny.github.io/contoso-waf-caf-comic/e01.html) | CAF · Strategy | 7 min |
| EP 2 | [Count Every Crate](https://ibranibeny.github.io/contoso-waf-caf-comic/e02.html) | CAF · Plan | 8 min |
| EP 3 | [Build the Harbor](https://ibranibeny.github.io/contoso-waf-caf-comic/e03.html) | CAF · Ready | 9 min |
| EP 4 | [Roads, Gates and Docks](https://ibranibeny.github.io/contoso-waf-caf-comic/e04.html) | CAF · Ready | 9 min |
| EP 5 | [The Great Migration](https://ibranibeny.github.io/contoso-waf-caf-comic/e05.html) | CAF · Adopt | 8 min |
| EP 6 | [Rules of the Harbor](https://ibranibeny.github.io/contoso-waf-caf-comic/e06.html) | CAF · Govern | 9 min |
| EP 7 | [Guards at Every Gate](https://ibranibeny.github.io/contoso-waf-caf-comic/e07.html) | CAF · Secure | 8 min |
| EP 8 | [The Control Tower](https://ibranibeny.github.io/contoso-waf-caf-comic/e08.html) | CAF · Manage | 8 min |
| EP 9 | [The Five Pillars](https://ibranibeny.github.io/contoso-waf-caf-comic/e09.html) | WAF · Overview | 9 min |
| EP 10 | [When an Engine Fails](https://ibranibeny.github.io/contoso-waf-caf-comic/e10.html) | WAF · Reliability | 9 min |
| EP 11 | [Lock Every Hatch](https://ibranibeny.github.io/contoso-waf-caf-comic/e11.html) | WAF · Security | 9 min |
| EP 12 | [Fuel Wisely](https://ibranibeny.github.io/contoso-waf-caf-comic/e12.html) | WAF · Cost Optimization | 8 min |
| EP 13 | [Same Every Time](https://ibranibeny.github.io/contoso-waf-caf-comic/e13.html) | WAF · Operational Excellence | 9 min |
| EP 14 | [Rush Hour](https://ibranibeny.github.io/contoso-waf-caf-comic/e14.html) | WAF · Performance Efficiency | 8 min |
| Finale | [Land It Well](https://ibranibeny.github.io/contoso-waf-caf-comic/e15.html) | Finale | 9 min |

## Cast

Dewi (solution engineer, always asks "where does it say that?"), Bayu (platform team lead, CAF), Sora (workload architect, WAF), Pip (the platform team's cloud drone) and Ratna (Contoso's chief executive). Seven villains stand for real anti-patterns: Vague, Sprawl, Open Hatch, Lone Point, Coin Gobbler, Snowflake and Bottleneck.

## Sources and accuracy

Every technical statement links to the Microsoft Learn page it comes from (98 pages, retrieved 15 November 2026). The sky harbor and airships are an analogy to make the ideas memorable; the lesson cards and diagrams use the real Azure names. Learn changes often, so re-check the linked pages before you present.

Contoso and all characters are fictional. Artwork is AI-generated for this comic. Azure service icons are the official Azure Architecture Icons, used for training material. This is workshop material, not official Microsoft documentation.

## Files

Static site, no build step needed to host: `index.html`, `e00.html` to `e15.html`, `diagrams.html`, `cheatsheet.html`, `sources.html`, with images in `assets/`.
