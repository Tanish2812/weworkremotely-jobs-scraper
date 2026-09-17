"""
We Work Remotely job Scraper
Collects every remote job listing, one sheet per category plus an "All Jobs" sheet.

OUTPUT:
WeWorkRemotely.xlsx
"""


import os
import time
import random
from datetime import date, timedelta
import re
import pandas as pd
from openpyxl.styles import Font
import undetected_chromedriver as uc
from selenium.common import TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


driver = uc.Chrome()
wait = WebDriverWait(driver, 10)
driver.maximize_window()


def safe_text(parent, by, value):
    found = parent.find_elements(by, value)
    return found[0].text.strip() if found else ""

def safe_attr(parent, by, value, attr):
    found = parent.find_elements(by, value)
    return found[0].get_attribute(attr) if found else ""

def real_date(text):
    if not text:
        return ""
    t = text.strip()
    if t.upper() == "NEW":
        return str(date.today())
    m = re.match(r"(\d+)d", t)
    if m:
        return str(date.today() - timedelta(days=int(m.group(1))))
    if re.match(r"(\d+)h", t) or re.match(r"(\d+)m", t):
        return str(date.today())
    return t



jobs_by_category = {}
skipped = []
NOT_A_PLACE = {"Featured", "Boosted", "Top 100", "Full-Time", "Contract", "Full-Time/Part-Time"}
BASE_URL = "https://weworkremotely.com/"
OUTPUT_FILE = "WeWorkRemotely.xlsx"


driver.get(BASE_URL)
time.sleep(random.uniform(5, 10))

try:
    sections = wait.until(
        EC.presence_of_all_elements_located((
            By.CSS_SELECTOR, "#search-results .jobs"
        ))
    )


except TimeoutException:
    print("Timeout loading homepage")
    driver.quit()
    raise SystemExit

category_links = [safe_attr(section, By.CSS_SELECTOR, "h2 a", "href") for section in sections]

for link in category_links:
    driver.get(link)
    time.sleep(random.uniform(5, 10))

    if 'just a moment' in driver.title.lower():
        print("Asking for verification !!")
        input("Press Enter, once the job list is visible: ")

    jobs = wait.until(
        EC.presence_of_all_elements_located((
            By.CSS_SELECTOR, "li.new-listing-container"
        ))
    )

    category = safe_text(driver, By.CSS_SELECTOR, "h2 a")

    for job in jobs:
        job_url = safe_attr(job, By.CSS_SELECTOR, ".listing-link--unlocked", 'href')

        if not job_url:
            skipped.append(category)
            continue

        try:
            title = safe_text(job, By.CSS_SELECTOR, ".new-listing__header__title__text")
            company = safe_text(job, By.CSS_SELECTOR, ".new-listing__company-name")
            headquarter = safe_text(job, By.CSS_SELECTOR, ".new-listing__company-headquarters")
            date_label = safe_text(job, By.CSS_SELECTOR, ".new-listing__header__icons__date")

            labels = [item.text.strip() for item in job.find_elements(By.CSS_SELECTOR, ".new-listing__categories__category")]
            regions = [label for label in labels if label not in NOT_A_PLACE and not label.startswith("$")]
            tags = [label for label in labels if label in NOT_A_PLACE or label.startswith("$")]



            jobs_by_category.setdefault(category, []).append({
                "Category": category,
                "Title": title,
                "Company": company,
                "Headquarter": headquarter,
                "Date": real_date(date_label),
                "URL": job_url,
                "Region": ', '.join(regions),
                "Tags": ', '.join(tags)

            })

        except StaleElementReferenceException:
            skipped.append(category)

seen = set()

for name in jobs_by_category:
    cleaned = []
    for row in jobs_by_category[name]:
        if row["URL"] in seen:
            continue
        seen.add(row["URL"])
        cleaned.append(row)
    jobs_by_category[name] = cleaned

all_jobs = [row for rows in jobs_by_category.values() for row in rows]

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    all_df = pd.DataFrame(all_jobs).sort_values("Date", ascending=False)

    all_df.to_excel(writer, sheet_name="All Jobs", index=False)

    for name, rows in sorted(jobs_by_category.items()):
        df = pd.DataFrame(rows).drop(columns=["Category"])
        df = df.sort_values("Date", ascending=False)
        df.to_excel(writer, sheet_name=name.replace(" Jobs", "")[:31], index=False)

    for sheet in writer.book.worksheets:
        for cell in sheet[1]:
            cell.font = Font(bold=True)

        for column in sheet.columns:
            letter = column[0].column_letter
            longest = max((len(str(c.value)) for c in column if c.value), default=0)
            sheet.column_dimensions[letter].width = min(max(longest + 2, 10), 60)
        sheet.freeze_panes = "A2"

print(f"Saved {len(all_jobs)} jobs across {len(jobs_by_category) + 1} sheets, skipped {len(skipped)}")

driver.quit()
os._exit(0)