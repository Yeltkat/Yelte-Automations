"""
Generates two deliberately messy order exports in input/.

This mimics what small webshops typically send: two monthly exports from
different tools, with inconsistent headers, duplicate rows, mixed date and
price formats, stray whitespace, and a few broken rows.

The data is fake and generated with a fixed seed, so the output is the same
every time you run it.
"""

import random
from pathlib import Path

import pandas as pd

SEED = 42
INPUT_DIR = Path(__file__).parent / "input"

FIRST_NAMES = ["Emma", "Daan", "Sophie", "Lucas", "Julia", "Sem", "Tess", "Finn",
               "Anna", "Lars", "Mila", "Noah", "Sara", "Jonas", "Lotte", "Max"]
LAST_NAMES = ["de Vries", "Jansen", "Bakker", "Visser", "Smit", "Meijer",
              "Müller", "Schmidt", "Peeters", "Janssens", "Mulder", "de Boer"]
PRODUCTS = {
    "Ceramic Mug": 12.50,
    "Linen Tote Bag": 18.00,
    "Scented Candle": 22.95,
    "Notebook A5": 9.99,
    "Desk Plant": 27.50,
    "Wool Scarf": 39.00,
}
# Each country is written several different ways, like in real exports.
COUNTRY_SPELLINGS = {
    "NL": ["NL", "Netherlands", "netherlands", "Nederland", "The Netherlands", " NL "],
    "DE": ["DE", "Germany", "germany", "Deutschland", "Duitsland"],
    "BE": ["BE", "Belgium", "belgie", "België", "Belgium "],
}
PHONE_PREFIX = {"NL": "+31", "DE": "+49", "BE": "+32"}


def messy_name(first, last, rng):
    name = f"{first} {last}"
    style = rng.random()
    if style < 0.2:
        return name.upper()
    if style < 0.4:
        return name.lower()
    if style < 0.55:
        return f"  {name} "
    return name


def messy_email(first, last, rng):
    local = f"{first}.{last}".lower().replace(" ", "").replace("ü", "u")
    email = f"{local}@example.com"
    if rng.random() < 0.25:
        email = email.upper()
    if rng.random() < 0.2:
        email = f" {email}  "
    return email


def messy_phone(country, rng):
    digits = f"6{rng.randint(10000000, 99999999)}"
    if country == "NL":
        return rng.choice([
            f"06{digits[1:]}",
            f"06-{digits[1:5]} {digits[5:]}",
            f"+31 6 {digits[1:]}",
            f"0031{digits}",
        ])
    return rng.choice([
        f"{PHONE_PREFIX[country]} {digits}",
        f"00{PHONE_PREFIX[country][1:]}{digits}",
    ])


def messy_date(day, month, rng):
    return rng.choice([
        f"2026-{month:02d}-{day:02d}",
        f"{day:02d}/{month:02d}/2026",
        f"{day}-{month}-2026",
        pd.Timestamp(2026, month, day).strftime("%b %d %Y"),
    ])


def messy_price(price, rng):
    return rng.choice([
        f"{price:.2f}",
        f"€ {price:.2f}".replace(".", ","),
        f"{price:.2f} EUR".replace(".", ","),
        f"€{price:.2f}",
    ])


def make_orders(month, start_id, count, rng):
    rows = []
    for i in range(count):
        first, last = rng.choice(FIRST_NAMES), rng.choice(LAST_NAMES)
        country = rng.choice(list(COUNTRY_SPELLINGS))
        product = rng.choice(list(PRODUCTS))
        rows.append({
            "order_id": f"ORD-{start_id + i}",
            "name": messy_name(first, last, rng),
            "email": messy_email(first, last, rng),
            "phone": messy_phone(country, rng),
            "date": messy_date(rng.randint(1, 28), month, rng),
            "product": product if rng.random() > 0.15 else f" {product.lower()} ",
            "qty": rng.choice(["1", "2", "3", " 2 ", "1.0", "4"]),
            "price": messy_price(PRODUCTS[product], rng),
            "country": rng.choice(COUNTRY_SPELLINGS[country]),
        })
    return rows


def add_problems(rows, rng):
    """Inject the kinds of issues that show up in real exports."""
    # Exact duplicates (e.g. the same order exported twice)
    for row in rng.sample(rows, 4):
        rows.append(dict(row))
    # "Near" duplicates: same order, different casing/whitespace
    for row in rng.sample(rows, 3):
        dup = dict(row)
        dup["name"] = f"  {dup['name'].strip().upper()}  "
        dup["email"] = dup["email"].strip().upper()
        rows.append(dup)
    # Rows that can't be fixed automatically and need a human to look at them
    broken = rng.sample(rows, 3)
    broken[0]["qty"] = ""
    broken[1]["price"] = "n/a"
    broken[2]["email"] = "not-an-email"
    # Completely empty rows
    rows.extend([{k: "" for k in rows[0]} for _ in range(3)])
    rng.shuffle(rows)
    return rows


def main():
    rng = random.Random(SEED)
    INPUT_DIR.mkdir(exist_ok=True)

    jan_rows = add_problems(make_orders(1, 1001, 40, rng), rng)
    feb_rows = add_problems(make_orders(2, 1041, 40, rng), rng)
    # A late January order that also shows up in the February export,
    # a classic cross-file duplicate
    feb_rows.insert(5, dict(next(r for r in jan_rows if r["order_id"])))

    # January: semicolon CSV export from "tool A"
    jan = pd.DataFrame(jan_rows)
    jan.columns = [" Order ID", "customer name", "E-mail", "Phone", "Order Date",
                   "Product", "Qty", "Unit Price", "Country"]
    jan.to_csv(INPUT_DIR / "orders_january.csv", index=False, sep=";")

    # February: Excel export from "tool B" with different header names
    feb = pd.DataFrame(feb_rows)
    feb.columns = ["OrderNumber", "Customer", "Email Address", "Telephone", "date",
                   "Item", "Quantity", "Price", "country "]
    feb.to_excel(INPUT_DIR / "orders_february.xlsx", index=False)

    print(f"Wrote {len(jan)} + {len(feb)} messy rows to {INPUT_DIR}/")


if __name__ == "__main__":
    main()
