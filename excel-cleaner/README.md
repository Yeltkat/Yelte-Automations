# Excel / CSV Cleaner

Merges messy spreadsheet exports into **one clean, formatted Excel report**, including a summary and a list of rows that need a human to look at them.

This demo uses two fake webshop order exports (one CSV, one Excel) that were exported from different tools, so they have different column names, date formats, price formats and plenty of duplicates.

## What it fixes

| Problem in the export | What the script does |
|---|---|
| Different headers per file (`E-mail` vs `Email Address`, `Qty` vs `Quantity`) | Maps them to one set of columns |
| `;` and `,` separated CSVs, plus `.xlsx` files | Detects the format automatically |
| `  FINN BAKKER `, `max jansen` | `Finn Bakker`, `Max Jansen` (keeps `de`/`van` lowercase) |
| `JULIA.VISSER@EXAMPLE.COM` | `julia.visser@example.com`, invalid emails get flagged |
| `06-1234 5678`, `0031654585178`, `+49 696728778` | `+31612345678` style international numbers |
| `26/01/2026`, `2026-01-16`, `10-1-2026`, `Jan 06 2026` | Real Excel dates |
| `€ 39,00`, `18,00 EUR`, `€9.99`, `18.00` | Numbers with € formatting |
| `Nederland`, `NL`, `Duitsland`, `België`, `BE` | `Netherlands`, `Germany`, `Belgium` |
| Empty rows, exact duplicates, duplicates that only differ in casing/spaces, the same order in two files | Removed |
| Rows with a missing quantity or broken email | Moved to a **Needs Review** sheet with the reason, instead of silently deleted |

## Before / after

**Before** (`input/orders_january.csv`):

```
 Order ID;customer name;E-mail;Phone;Order Date;Product;Qty;Unit Price;Country
ORD-1026;Julia Visser;JULIA.VISSER@EXAMPLE.COM;0031654585178;26/01/2026;Linen Tote Bag;4;18.00;Nederland
ORD-1038;Daan Visser;daan.visser@example.com;0049687267850;2026-01-16;Notebook A5;3;€9.99;Deutschland
ORD-1034;FINN BAKKER;finn.bakker@example.com;0032624305904;08/01/2026;Wool Scarf;4;€39.00;BE
;;;;;;;;
```

**After** (`output/clean_orders_report.xlsx`, sheet *Clean Orders*):

| Order ID | Customer | Email | Phone | Order Date | Product | Quantity | Unit Price | Country | Total |
|---|---|---|---|---|---|---|---|---|---|
| ORD-1034 | Finn Bakker | finn.bakker@example.com | +32624305904 | 08-01-2026 | Wool Scarf | 4 | €39.00 | Belgium | €156.00 |
| ORD-1038 | Daan Visser | daan.visser@example.com | +49687267850 | 16-01-2026 | Notebook A5 | 3 | €9.99 | Germany | €29.97 |
| ORD-1026 | Julia Visser | julia.visser@example.com | +31654585178 | 26-01-2026 | Linen Tote Bag | 4 | €18.00 | Netherlands | €72.00 |

The report has four sheets:

- **Clean Orders**: the cleaned data, with styled headers, filters, frozen top row and € formatting
- **Summary**: revenue per month, per country and best-selling products
- **Needs Review**: rows that couldn't be fixed automatically, shown with their original values and the reason
- **Cleaning Log**: what was changed, in numbers

```
Cleaning log:
  Rows read from all files                           101
  Empty rows removed                                   6
  Individual cells reformatted                       518
  Duplicate rows removed                              13
  Same order ID in multiple rows/files removed         2
  Rows moved to 'Needs Review'                         4
  Clean rows in final report                          76
```

## Run it

```bash
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt

python generate_sample_data.py    # (optional) recreate the messy sample files in input/
python main.py                    # writes output/clean_orders_report.xlsx
```

To use your own files, drop them in `input/` or point the script at another folder:

```bash
python main.py --input-dir path/to/exports --output path/to/report.xlsx
```

## Adapting it to other spreadsheets

Everything client-specific sits at the top of `main.py`:

- `COLUMN_ALIASES`: which header names map to which column
- `COUNTRY_MAP`: spellings to normalize
- `DATE_FORMATS`: accepted date formats (day-first by default; swap the order for US data)
- `DEFAULT_PHONE_PREFIX`: country code for local numbers starting with `0`

## Files

```
excel-cleaner/
├── main.py                  # the cleaner
├── generate_sample_data.py  # creates the fake messy input files
├── requirements.txt
├── input/                   # messy sample exports
└── output/                  # the generated report
```

All names, emails and phone numbers in the sample data are fake.
