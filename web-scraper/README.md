# Web Scraper

Collects structured data from a public website and delivers it as a **clean CSV and a formatted Excel file**.

This demo scrapes all 1,000 books from [books.toscrape.com](https://books.toscrape.com), a sandbox site made specifically for practicing web scraping. The same approach works for public product catalogs, price lists, listings and directories.

## What it does

1. Checks `robots.txt` before fetching anything and refuses URLs it disallows
2. Reads all 50 categories from the sidebar
3. Walks through every page of every category by following the "next" button
4. Optionally (`--details`) opens every book page for the UPC, exact stock count, number of reviews and description
5. Saves `output/books.csv` and `output/books.xlsx` (with a *Per Category* summary sheet, filters and clickable links)

It's built to be a good citizen and hard to break:

- Waits between requests (0.5s by default, configurable with `--delay`)
- Retries automatically on temporary errors (429/5xx) with backoff
- Identifies itself with a clear User-Agent
- Skips a failing category instead of crashing the whole run
- Removes duplicate results

## Example output

```
Found 50 categories
  Travel                 page 1: 11 books so far
  Mystery                page 1: 20 books so far
  Mystery                page 2: 32 books so far
  ...
Scraped 1000 books from 50 categories with 81 requests in 64s
```

| title | category | price_gbp | rating | in_stock |
|---|---|---|---|---|
| It's Only the Himalayas | Travel | 45.17 | 2 | True |
| Full Moon over Noah's Ark: An Odyssey to Mount Ararat and Beyond | Travel | 49.43 | 4 | True |
| See America: A Celebration of Our National Parks & Treasured Sites | Travel | 48.87 | 3 | True |
| Vagabonding: An Uncommon Guide to the Art of Long-Term World Travel | Travel | 36.94 | 2 | True |

Each row also has the book's `url`, `image_url` and a `scraped_at` timestamp. With `--details` you also get `upc`, `stock_count`, `reviews` and `description`.

## Run it

```bash
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt

python main.py                          # everything, listing pages only (about 1 minute)
python main.py --category Travel        # one category
python main.py --limit 3 --details      # first 3 categories, including each book's detail page
```

A full run with `--details` makes about 1,080 requests and takes roughly 10 minutes at the default delay.

## Responsible scraping

I only take scraping jobs for **publicly available data** where the website's terms allow it. That means no content behind logins, no collecting personal data, and no sites that prohibit scraping. The scraper checks `robots.txt` and rate-limits itself by default.

## Files

```
web-scraper/
├── main.py            # the scraper
├── requirements.txt
└── output/
    ├── books.csv      # sample result (1,000 books)
    └── books.xlsx     # same data, formatted, with a per-category summary
```
