# We Work Remotely Jobs Scraper

It scrapes every remote job listing from We Work Remotely and gives a formatted Excel file with one sheet per category, plus an "All Jobs" sheet sorted newest first.

## Output

print statement example: Saved 344 jobs across 10 sheets, skipped 18
and a file with name 'WeWorkRemotely.xlsx'

| Column | Example |
| ------ | ------- |
| Title | Entry-Level Account Manager |
| Company | NoGigiddy |
| Headquarter | Atlanta, Georgia |
| Date | 2026-09-17 |
| URL | https://weworkremotely.com/remote-jobs/nogigiddy-entry-level-account-manager-4 |
| Region | Anywhere in the World |
| Tags | Featured, Top 100, Contract |

Sheets come out sorted newest first, with bold headers, sized columns and a frozen top row.

## The problems I faced

The site sits behind Cloudflare bot protection. A normal automated browser gets
trapped in an endless verification loop — solving the checkbox doesn't help,
because Cloudflare is inspecting the browser itself, not the person clicking.

This uses undetected-chromedriver, which patches the automation signals
Cloudflare looks for, plus 5–10 second delays between page loads so the traffic
resembles someone reading each page.

## How it works

1. Collects every category URL from the homepage
2. Visits each category page and extracts every listing
3. Skips the adverts mixed into the results
4. Converts relative dates ("12d", "NEW") into real dates
5. Removes jobs appearing under more than one category
6. Writes the styled Excel workbook


## Requirements

Python 3.8+
install with:
pip install -r requirements.txt


## Usage

python weworkremotely-scraper.py 
