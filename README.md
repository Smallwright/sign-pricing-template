# Sign shop pricing template

A free spreadsheet for pricing sign jobs from your own costs: material per square foot, waste, ink, labour at your shop rate, hardware, markup and a minimum per job. It covers vinyl banners, coroplast yard signs, ACM (aluminum composite) signs and printed adhesive vinyl, with worked examples you can check by hand.

- **`sign-pricing-template.xlsx`**: Excel, Google Sheets, LibreOffice or Numbers. Yellow cells are your inputs and everything else is a formula. It has four sheets: Read me, Costs, Products and Quote. The values calculate when you open the file, so a file preview may show blank cells.
- **`materials.csv`, `products.csv`, `examples.csv`**: the same numbers as plain CSV, for importing or for scripts.
- **`build.py`**: regenerates all of the above from the numbers at the top of the file (`pip install openpyxl`).

The example costs are real supplier prices checked in November 2026, with sources below. **They are not your prices.** Replace them with what you actually pay and charge.

## How do I price a banner?

Work out what the banner costs you, multiply by your markup, and never go below your minimum:

```
area (sq ft)  = width (in) x height (in) / 144 x quantity
material      = area x (board or media + any vinyl face, per sq ft) x (1 + waste %)
ink           = area x ink cost per sq ft x printed sides
labour        = (setup minutes + labour minutes per sq ft x area) / 60 x shop rate
cost          = material + ink + labour + hardware (grommets, stakes)
price         = max(minimum, cost x (1 + markup %))
```

**Worked example: a 3 x 6 ft, 13 oz vinyl banner, hemmed, with 8 grommets**

| Step | Working | $ |
|---|---|---|
| Area | 36 x 72 / 144 | 18 sq ft |
| Material | 18 x $0.32 (a $216 roll of 54 in x 150 ft = 675 sq ft) x 1.30 waste | 7.49 |
| Ink | 18 x $0.25 | 4.50 |
| Labour | (15 min setup + 1 min per sq ft x 18) / 60 x $75 an hour | 41.25 |
| Grommets | 8 x $0.10 | 0.80 |
| **Cost** | | **54.04** |
| **Price** | cost x (1 + 120% markup), $45 minimum, rounded to the cent | **118.88** ($6.60 per sq ft) |

Why 30% waste: a 36 in banner printed on a 54 in roll throws away a third of the width unless you can nest another job beside it.

Three things the example shows:
- **Labour is most of the cost of a small banner.** Material and ink are about $12, while the labour is $41. Shops that price banners only "per square foot" lose money on small ones. That's why every product here has a minimum and a setup time.
- **The price per square foot falls as the banner grows.** In this template a 2 x 4 ft banner comes to $9.48/sq ft and a 3 x 10 ft to $5.68/sq ft, which matches how shops price in practice.
- **On small jobs, your shop rate matters more than your material price.** On big runs it's the other way round: for 25 yard signs at 24 x 36 in, the coroplast costs more than the labour.

**What sign shops charge (for checking your result, not for copying):**
- Prices vary widely. In a 2022 Signs101 thread pricing a 3 x 8 ft full-colour banner, answers ranged from about $3.50–4/sq ft to $195–265 for the banner (about $8–11/sq ft) ([Signs101](https://www.signs101.com/threads/3-x-8-full-color-full-bleed-digitally-printed-13oz-banner-price.169033/)).
- Signs.com says "a standard pricing rule for local sign shops is to charge $8 per square foot for printed banners" ([Signs.com](https://www.signs.com/blog/banner-prices-the-basics/)).
- 4over4, an online printer, starts a 2 x 4 ft banner at $79 ([4over4](https://www.4over4.com/guide/banner-prices-explained)).

## All the worked examples

From `examples.csv`. Each one is the Quote sheet's formula applied to the example costs and products.

<!-- examples:start -->
| Job | Cost $ | Price $ | Each $ | Per sq ft $ |
|---|---|---|---|---|
| 2 x 4 ft banner, 4 grommets | 34.48 | 75.85 | 75.85 | 9.48 |
| 3 x 6 ft banner, 8 grommets | 54.04 | 118.88 | 118.88 | 6.60 |
| 3 x 10 ft banner, 12 grommets | 77.43 | 170.35 | 170.35 | 5.68 |
| 18 x 24 in yard sign with stake, 1 | 13.95 | 20.92 | 20.92 | 6.97 |
| 18 x 24 in yard signs with stakes, 10 | 83.24 | 124.86 | 12.49 | 4.16 |
| 24 x 36 in yard signs, 25 | 316.20 | 474.30 | 18.97 | 3.16 |
| 24 x 36 in ACM sign, 1 | 49.48 | 98.97 | 98.97 | 16.50 |
| 4 x 8 ft ACM sign, 1 | 209.75 | 419.50 | 419.50 | 13.11 |
| 3 x 5 ft window graphic, 1 | 59.69 | 149.23 | 149.23 | 9.95 |
| 12 x 12 in decals, 20 | 73.34 | 183.35 | 9.17 | 9.17 |
<!-- examples:end -->

The settings that produce these, from `products.csv`:

| Product | Materials | Waste % | Setup min | Labour min per sq ft | Hardware $ | Markup % | Minimum $ |
|---|---|---|---|---|---|---|---|
| Vinyl banner, 13 oz, hemmed | banner vinyl | 30 | 15 | 1.0 | grommets $0.10 each, as extras | 120 | 45 |
| Coroplast yard sign, 4 mm, printed direct (with or without an H-stake) | coroplast | 15 | 5 | 0.5 | $1.50 H-stake | 50 | 20 |
| ACM sign, 3 mm, printed vinyl face (unlaminated) | ACM + adhesive vinyl | 15 | 10 | 2.0 | – | 100 | 60 |
| Printed adhesive vinyl (window, wall, decal) | adhesive vinyl | 20 | 15 | 1.5 | – | 150 | 35 |

"Printed direct" means a flatbed (UV) printer. If you print vinyl and mount it on coroplast, set the yard sign's face material to the adhesive vinyl, as the ACM sign does.

## Where the example costs come from (November 2026)

Only the banner roll and ink figures were checked on the source page. The other three are typical ranges from search listings: check them with your supplier.

| Material | Example | Per sq ft | Source |
|---|---|---|---|
| 13 oz banner vinyl | $216 per 54 in x 150 ft roll | $0.320 | [FastPlot 13 oz scrim banner](https://macroenter.com/FastPlot-EcoSolvent-Scrim-Vinyl-Banner-13oz-54x150-Roll) |
| Printable adhesive vinyl, calendered | $340 per 54 in x 150 ft roll | $0.504 | **Typical range, check your supplier:** about $325–$346 per 54 in x 150 ft roll in Nov 2026 search listings for Roland and GF calendered vinyl (we couldn't open the product pages to confirm). |
| 4 mm coroplast | $33.15 per 4 x 8 ft sheet | $1.036 | **Typical range, check your supplier:** about $33–35 a sheet retail in 10-packs in Nov 2026 search listings (Home Depot); sign suppliers are usually cheaper (we couldn't open the product page to confirm). |
| 3 mm ACM | $78.88 per 4 x 8 ft sheet | $2.465 | **Typical range, check your supplier:** about $79 a sheet for matte white in a Nov 2026 search listing (Curbell Plastics); finish and quantity change it (we couldn't open the product page to confirm). |
| Ink | $0.25 per printed sq ft | $0.25 | Roland DGA puts it at $0.12 to over $0.32 per sq ft ([Roland DGA](https://www.rolanddga.com/blog/what-is-the-best-printer-for-your-sign-shop)). Owners in a 2010 Signs101 thread worked out about $0.24–$0.50 ([thread](https://www.signs101.com/threads/anyone-know-roughly-what-ink-costs-per-sq-ft.71085)). Ink cost varies by printer and coverage, so use your own figure. |

The $75/hour shop rate, the $0.10 grommet and $1.50 H-stake prices, the waste percentages, the minutes and the markups are illustrative settings, not survey results. Laminate, design time, installation and delivery aren't included. Add laminate as a material (and as the face material where you use it), and add the others as extras.

## Using it
1. **Costs:** enter what you paid for each material and how many square feet that purchase gave you. Then enter your shop rate.
2. **Products:** set the materials, waste, minutes, hardware, markup and minimum for each thing you sell. To add a product, add a row. To add a material, add a row on Costs. The formulas pick up rows up to 200.
3. **Quote:** pick a product, then type the width and height in inches and the quantity. The extras column covers per-piece add-ons such as grommets.

Corrections and better-sourced numbers are welcome as issues or pull requests.

---
Made by Smallwright. If you'd rather not keep a spreadsheet, [Grommetbook](https://grommetbook.smallwright.com/) does the same math in a browser price book, with customer approval on their phone.
