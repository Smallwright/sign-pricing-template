#!/usr/bin/env python3
"""Builds sign-pricing-template.xlsx (live formulas) and the CSV copies from the numbers below.
Edit the numbers here or in the spreadsheet itself. Needs openpyxl: pip install openpyxl"""
import csv, pathlib
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

HERE = pathlib.Path(__file__).resolve().parent
SHOP_RATE = 75.0  # $ per hour of shop labour: replace with your own

# name, how you buy it, price paid ($), square feet in that purchase, where the example price came from
MATERIALS = [
    ('13 oz banner vinyl', '54 in x 150 ft roll', 216.00, 675, 'FastPlot 13 oz scrim banner, 54 in x 150 ft roll: $216 (Nov 2026)'),
    ('Printable adhesive vinyl (calendered)', '54 in x 150 ft roll', 340.00, 675, 'Supplier listings, Nov 2026: $325-$346 per 54 in x 150 ft roll'),
    ('4 mm coroplast', '4 x 8 ft sheet', 33.15, 32, 'Home Depot 10-pack, Nov 2026: $33.15 a sheet (trade suppliers are usually cheaper)'),
    ('3 mm ACM (aluminum composite)', '4 x 8 ft sheet', 78.88, 32, 'Curbell Plastics, Nov 2026: $78.88 a sheet, matte white'),
    ('Ink', 'per printed sq ft', 0.25, 1, 'Roland DGA: $0.12 to over $0.32 per sq ft; a 2010 Signs101 thread: about $0.24-$0.50. Use your own figure'),
]
INK = 'Ink'
# product, material, face material ('' = printed direct), waste %, printed sides, setup minutes per line, labour minutes per sq ft,
# hardware $ per piece, markup %, minimum $ per line
PRODUCTS = [
    ('Vinyl banner, 13 oz, hemmed', '13 oz banner vinyl', '', 30, 1, 15, 1.0, 0.00, 120, 45),
    ('Coroplast yard sign, 4 mm, printed direct', '4 mm coroplast', '', 15, 1, 5, 0.5, 0.00, 50, 20),
    ('Coroplast yard sign, 4 mm, printed direct, with H-stake', '4 mm coroplast', '', 15, 1, 5, 0.5, 1.50, 50, 20),
    ('ACM sign, 3 mm, printed vinyl face (unlaminated)', '3 mm ACM (aluminum composite)', 'Printable adhesive vinyl (calendered)', 15, 1, 10, 2.0, 0.00, 100, 60),
    ('Printed adhesive vinyl (window, wall, decal)', 'Printable adhesive vinyl (calendered)', '', 20, 1, 15, 1.5, 0.00, 150, 35),
]
# label, product, width in, height in, quantity, extras $ per piece (e.g. grommets)
EXAMPLES = [
    ('2 x 4 ft banner, 4 grommets', 'Vinyl banner, 13 oz, hemmed', 24, 48, 1, 0.40),
    ('3 x 6 ft banner, 8 grommets', 'Vinyl banner, 13 oz, hemmed', 36, 72, 1, 0.80),
    ('3 x 10 ft banner, 12 grommets', 'Vinyl banner, 13 oz, hemmed', 36, 120, 1, 1.20),
    ('18 x 24 in yard sign with stake, 1', 'Coroplast yard sign, 4 mm, printed direct, with H-stake', 18, 24, 1, 0),
    ('18 x 24 in yard signs with stakes, 10', 'Coroplast yard sign, 4 mm, printed direct, with H-stake', 18, 24, 10, 0),
    ('24 x 36 in yard signs, 25', 'Coroplast yard sign, 4 mm, printed direct', 24, 36, 25, 0),
    ('24 x 36 in ACM sign, 1', 'ACM sign, 3 mm, printed vinyl face (unlaminated)', 24, 36, 1, 0),
    ('4 x 8 ft ACM sign, 1', 'ACM sign, 3 mm, printed vinyl face (unlaminated)', 48, 96, 1, 0),
    ('3 x 5 ft window graphic, 1', 'Printed adhesive vinyl (window, wall, decal)', 36, 60, 1, 0),
    ('12 x 12 in decals, 20', 'Printed adhesive vinyl (window, wall, decal)', 12, 12, 20, 0),
]

def price(prod, w, h, qty, extras):
    """The same arithmetic as the spreadsheet formulas, for the CSV copy and the README."""
    p = next(x for x in PRODUCTS if x[0] == prod)
    cps = lambda name: next(x[2] / x[3] for x in MATERIALS if x[0] == name)
    area = w * h / 144 * qty
    mat = area * (cps(p[1]) + (cps(p[2]) if p[2] else 0)) * (1 + p[3] / 100)
    inkc = area * cps(INK) * p[4]
    lab = (p[5] + p[6] * area) / 60 * SHOP_RATE
    ext = (p[7] + extras) * qty
    cost = mat + inkc + lab + ext
    pr = max(p[9], r2(cost * (1 + p[8] / 100)))
    return dict(sqft=r2(area), material=r2(mat), ink=r2(inkc), labour=r2(lab), extras=r2(ext),
                cost=r2(cost), price=r2(pr), each=r2(pr / qty), per_sqft=r2(pr / area))

def r2(x):
    """Round half away from zero, as Excel's ROUND and cell display do."""
    return float(Decimal(repr(x)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))

def main():
    head = Font(bold=True, color='FFFFFF'); fill = PatternFill('solid', fgColor='1F4E79'); inp = PatternFill('solid', fgColor='FFF2CC')
    wb = Workbook()
    rd = wb.active; rd.title = 'Read me'
    for i, line in enumerate([
        'Sign shop pricing template (free, MIT licence): https://github.com/Smallwright/sign-pricing-template',
        'Yellow cells are inputs: change them to your own costs. Everything else is a formula.',
        '1. Costs: what you pay for each material and how many sq ft you get, plus your shop rate per hour.',
        '2. Products: waste, setup and labour minutes, hardware, markup and minimum for each thing you sell.',
        '3. Quote: pick a product, type the size in inches and the quantity; the price is worked out with the math shown.',
        'Price = cost x (1 + markup), never below the product minimum. Cost = material incl. waste + ink + labour + hardware.',
        'Add materials on Costs and products on Products as new rows (rows up to 200 are picked up). Face material = a printed vinyl applied to the board; leave it blank for direct print.',
        'The example costs are real supplier prices from November 2026 (sources on the Costs sheet). They are not your prices.',
    ], 1):
        rd.cell(i, 1, line)
    rd.column_dimensions['A'].width = 120

    c = wb.create_sheet('Costs')
    c.append(['Material', 'You buy it as', 'Price paid ($)', 'Sq ft in that purchase', 'Cost per sq ft ($)', 'Example price source'])
    for i, m in enumerate(MATERIALS, 2):  # noqa
        c.append([m[0], m[1], m[2], m[3], f'=C{i}/D{i}', m[4]])
        for col in 'CD': c[f'{col}{i}'].fill = inp
        c[f'E{i}'].number_format = '0.000'
    nmat = 200  # formulas look up rows 2-200, so you can add materials and products
    c['H1'] = 'Shop rate ($ per hour of labour)'; c['H1'].font = Font(bold=True); c['H2'] = SHOP_RATE; c['H2'].fill = inp
    ink_row = [m[0] for m in MATERIALS].index(INK) + 2

    p = wb.create_sheet('Products')
    p.append(['Product', 'Material', 'Face material (blank = printed direct)', 'Waste %', 'Printed sides', 'Setup minutes per line', 'Labour minutes per sq ft', 'Hardware $ per piece', 'Markup %', 'Minimum $ per line'])
    for i, r in enumerate(PRODUCTS, 2):
        p.append(list(r))
        for col in 'CDEFGHIJ': p[f'{col}{i}'].fill = inp
    nprod = 200
    dv = DataValidation(type='list', formula1=f'=Costs!$A$2:$A${nmat}', allow_blank=True); p.add_data_validation(dv); dv.add('B2:C200')

    q = wb.create_sheet('Quote')
    q.append(['Job', 'Product', 'Width (in)', 'Height (in)', 'Quantity', 'Extras $ per piece', 'Total sq ft', 'Material $ (with waste)', 'Ink $', 'Labour $', 'Hardware + extras $', 'Cost $', 'Price $', 'Price each $', 'Price per sq ft $'])
    rows = EXAMPLES + [('', '', None, None, None, None)] * 10
    PR = f'Products!$A$2:$A${nprod}'
    def pcol(col, r): return f'INDEX(Products!${col}$2:${col}${nprod},MATCH($B{r},{PR},0))'
    for r, ex in enumerate(rows, 2):
        q.append(list(ex))
        for col in 'ABCDEF': q[f'{col}{r}'].fill = inp
        g = lambda f: f'=IF($B{r}="","",{f})'
        q[f'G{r}'] = g(f'C{r}*D{r}/144*E{r}')
        cost_of = lambda name: f"INDEX(Costs!$E$2:$E${nmat},MATCH({name},Costs!$A$2:$A${nmat},0))"
        face = pcol('C', r)
        q[f'H{r}'] = g(f"G{r}*({cost_of(pcol('B', r))}+IF({face}=\"\",0,{cost_of(face)}))*(1+{pcol('D', r)}/100)")
        q[f'I{r}'] = g(f"G{r}*Costs!$E${ink_row}*{pcol('E', r)}")
        q[f'J{r}'] = g(f"({pcol('F', r)}+{pcol('G', r)}*G{r})/60*Costs!$H$2")
        q[f'K{r}'] = g(f"({pcol('H', r)}+F{r})*E{r}")
        q[f'L{r}'] = g(f'H{r}+I{r}+J{r}+K{r}')
        q[f'M{r}'] = g(f"MAX({pcol('J', r)},ROUND(L{r}*(1+{pcol('I', r)}/100),2))")
        q[f'N{r}'] = g(f'M{r}/E{r}')
        q[f'O{r}'] = g(f'M{r}/G{r}')
        for col in 'GHIJKLMNO': q[f'{col}{r}'].number_format = '0.00'
    dv2 = DataValidation(type='list', formula1=f'=Products!$A$2:$A${nprod}', allow_blank=True); q.add_data_validation(dv2); dv2.add(f'B2:B{len(rows) + 1}')
    for ws in (c, p, q):
        for cell in ws[1]: cell.font = head; cell.fill = fill; cell.alignment = Alignment(wrap_text=True, vertical='top')
        ws.freeze_panes = 'A2'
        for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = max(12, min(48, max(len(str(x.value or '')) for x in col[1:]) + 2)) if col[0].column_letter not in 'F' or ws is not c else 70
    c.column_dimensions['H'].width = 34
    from openpyxl.workbook.properties import CalcProperties
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(HERE / 'sign-pricing-template.xlsx')

    with open(HERE / 'materials.csv', 'w', newline='') as f:
        w = csv.writer(f); w.writerow(['material', 'bought_as', 'price_paid_usd', 'sqft_in_purchase', 'cost_per_sqft_usd', 'source'])
        for m in MATERIALS: w.writerow([m[0], m[1], f'{m[2]:.2f}', m[3], f'{m[2] / m[3]:.4f}', m[4]])
    with open(HERE / 'products.csv', 'w', newline='') as f:
        w = csv.writer(f); w.writerow(['product', 'material', 'face_material', 'waste_percent', 'printed_sides', 'setup_minutes_per_line', 'labour_minutes_per_sqft', 'hardware_usd_per_piece', 'markup_percent', 'minimum_usd_per_line'])
        w.writerows(PRODUCTS)
    with open(HERE / 'examples.csv', 'w', newline='') as f:
        w = csv.writer(f); w.writerow(['job', 'product', 'width_in', 'height_in', 'quantity', 'extras_usd_per_piece', 'total_sqft', 'material_usd', 'ink_usd', 'labour_usd', 'hardware_extras_usd', 'cost_usd', 'price_usd', 'price_each_usd', 'price_per_sqft_usd'])
        for ex in EXAMPLES:
            r = price(ex[1], ex[2], ex[3], ex[4], ex[5])
            w.writerow(list(ex) + [r[k] for k in ('sqft', 'material', 'ink', 'labour', 'extras', 'cost', 'price', 'each', 'per_sqft')])
            print(f"{ex[0]:40} cost {r['cost']:8.2f} price {r['price']:8.2f} each {r['each']:7.2f} /sqft {r['per_sqft']:6.2f}")
    print('shop rate', SHOP_RATE)
    rd = HERE / 'README.md'
    if rd.exists():
        t = rd.read_text(); a, z = '<!-- examples:start -->', '<!-- examples:end -->'
        if a in t:
            rows = ['| Job | Cost $ | Price $ | Each $ | Per sq ft $ |', '|---|---|---|---|---|']
            for ex in EXAMPLES:
                r = price(ex[1], ex[2], ex[3], ex[4], ex[5])
                rows.append(f"| {ex[0]} | {r['cost']:.2f} | {r['price']:.2f} | {r['each']:.2f} | {r['per_sqft']:.2f} |")
            rd.write_text(t[:t.index(a) + len(a)] + '\n' + '\n'.join(rows) + '\n' + t[t.index(z):])

if __name__ == '__main__':
    main()
