# HNX Bond Auction Data Pipeline

Automated data pipeline that scrapes Vietnamese government bond auction
results from the Hanoi Stock Exchange (HNX) and maintains a single,
always-up-to-date Excel file — removing the need for manual lookups and
copy-pasting.

## What it does

- **Baseline build**: scrapes the full auction history (2016–present,
  ~2,400+ records) into a formatted Excel file.
- **Incremental update**: on each run, fetches only records published
  since the last update, merges them in without duplication, and re-saves
  the file with formatting intact.

## Tech stack

Python · Playwright (browser automation) · pandas (data cleaning) ·
openpyxl (Excel output)

## Key engineering challenges solved

- Corrected numeric parsing for Vietnamese-formatted decimals (comma as
  decimal separator), which pandas misreads by default.
- Replaced fixed-page-size pagination assumptions with a loop that tracks
  progress against the site's reported total record count.
- Worked around corporate IT restrictions by driving the pre-installed
  Microsoft Edge browser instead of downloading a separate Chromium build.
- Fixed intermittent automation failures caused by page-transition
  animations hiding form elements from the browser driver.

## Usage

```bash
pip install -r requirements.txt
python scrape_hnx.py     # one-time baseline build
python update_hnx.py     # run periodically to fetch new records
```

## Notes

Data source: [hnx.vn — Bond Auction Results](https://hnx.vn/trai-phieu/ket-qua-dau-thau.html).
Output Excel file is not tracked in this repository (see `.gitignore`) since
it's a generated data artifact rather than source code.
