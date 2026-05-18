import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import random
from tqdm import tqdm
import os


API_KEYS = [
    "71449a0c6a4bf546c700d8152502b01a9a1e2958747c1bd406bc0558d9ec3a1f",
    "56a4428921126bfb2ddb68f2de2e4923bce0a16e4d9a3bba1097e2f732c50250",
    "3e497dd5a127bb3aa38a831cd261b76f9b8f7637c6285677fde0c48c3058929d",
    "rfRRSHFqzpJyv396fwPWGVRK"
]  
INPUT_FILE = "Growth For Impact Data Assignment.xlsx"
OUTPUT_FILE = "Enriched_Companies_Wide.xlsx"
BATCH_SIZE = 50         
SLEEP_MIN = 1
SLEEP_MAX = 3
MAX_JOBS = 3
RETRIES = 3
TIMEOUT = 30             

def serp_search(query, company_index, retries=RETRIES):
    url = "https://serpapi.com/search.json"
    current_key = API_KEYS[company_index % len(API_KEYS)]
    params = {
        "q": query,
        "engine": "google",
        "api_key": current_key,
    }
    for attempt in range(retries):
        try:
            r = requests.get(url, params=params, timeout=TIMEOUT)
            data = r.json()
            if "organic_results" in data and len(data["organic_results"]) > 0:
                return data["organic_results"][0].get("link")
        except Exception as e:
            print(f"Attempt {attempt+1} failed for query '{query}' with key {current_key}: {e}")
            time.sleep(5)
    return None


def find_website(company, index):
    return serp_search(f"{company} official site", index)

def find_linkedin(company, index):
    return serp_search(f"{company} site:linkedin.com", index)

def find_careers(company_or_url, index):
    return serp_search(f"{company_or_url} careers site", index)

def identify_jobs_page(careers_url):
    if not careers_url:
        return None
    try:
        r = requests.get(careers_url, timeout=10)
        soup = BeautifulSoup(r.text, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"].lower()
            if any(ats in href for ats in ["greenhouse.io", "lever.co", "workday", "jobs", "careers"]):
                return href if href.startswith("http") else requests.compat.urljoin(careers_url, href)
    except:
        pass
    return careers_url


def scrape_jobs(jobs_url, max_jobs=MAX_JOBS):
    if not jobs_url:
        return []
    jobs = []
    try:
        r = requests.get(jobs_url, timeout=10)
        soup = BeautifulSoup(r.text, "html.parser")
        for a in soup.find_all("a", href=True):
            text = a.get_text(" ", strip=True)
            if text and len(jobs) < max_jobs:
                url = a["href"] if a["href"].startswith("http") else requests.compat.urljoin(jobs_url, a["href"])
                jobs.append({
                    "Job_Title": text[:80],
                    "Job_Link": url
                })
        return jobs
    except Exception as e:
        print("Job scrape error:", e)
        return jobs


def enrich_data(df):
    enriched_rows, jobs_data = [], []
    for i, row in tqdm(df.iterrows(), total=len(df), desc="Companies"):
        company = str(row["Company Name"])
        print(f"\n🔹 Processing company: {company}")

        website = find_website(company, i)
        print(f"Website: {website}")

        linkedin = find_linkedin(company, i)
        print(f"LinkedIn: {linkedin}")

        careers = find_careers(website or company, i)
        print(f"Careers: {careers}")

        jobs_page = identify_jobs_page(careers or website)
        print(f"Jobs Page: {jobs_page}")

        jobs = scrape_jobs(jobs_page)
        print(f"Jobs found: {len(jobs)}")

        enriched_rows.append({
            "Company Name": company,
            "Website URL": website,
            "Linkedin URL": linkedin,
            "Careers Page URL": careers,
            "Job listings page URL": jobs_page
        })

        for job in jobs:
            jobs_data.append({
                "Company Name": company,
                **job
            })

        time.sleep(random.uniform(SLEEP_MIN, SLEEP_MAX))  

    return pd.DataFrame(enriched_rows), pd.DataFrame(jobs_data)


def merge_jobs(enriched_df, jobs_df, max_jobs=MAX_JOBS):
    merged_rows = []
    for _, row in enriched_df.iterrows():
        company = row["Company Name"]
        company_jobs = jobs_df[jobs_df["Company Name"] == company].head(max_jobs)

        job_cols = {}
        for i, (_, job) in enumerate(company_jobs.iterrows(), start=1):
            job_cols[f"job post{i} URL"] = job["Job_Link"]
            job_cols[f"job post{i} title"] = job["Job_Title"]

        for i in range(len(company_jobs)+1, max_jobs+1):
            job_cols[f"job post{i} URL"] = ""
            job_cols[f"job post{i} title"] = ""

        merged_rows.append({**row.to_dict(), **job_cols})

    return pd.DataFrame(merged_rows)


if __name__ == "__main__":
    df = pd.read_excel(INPUT_FILE, sheet_name=0)
    all_enriched, all_jobs = pd.DataFrame(), pd.DataFrame()

    total_companies = len(df)
    for start in range(0, total_companies, BATCH_SIZE):
        batch_df = df.iloc[start:start+BATCH_SIZE]
        print(f"\n📦 Processing batch {start//BATCH_SIZE + 1} ({len(batch_df)} companies)")
        enriched_df, jobs_df = enrich_data(batch_df)

        all_enriched = pd.concat([all_enriched, enriched_df], ignore_index=True)
        all_jobs = pd.concat([all_jobs, jobs_df], ignore_index=True)

        
        final_df = merge_jobs(all_enriched, all_jobs, max_jobs=MAX_JOBS)
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            final_df.to_excel(writer, sheet_name="Companies", index=False)
            pd.DataFrame({
                "Step": [
                    "1. Loaded raw dataset",
                    "2. Used SerpAPI for Website, LinkedIn, Careers",
                    "3. Parsed jobs page for ATS",
                    "4. Scraped up to 3 jobs per company",
                    "5. Merged jobs into single row per company"
                ]
            }).to_excel(writer, sheet_name="Methodology", index=False)

        print(f"💾 Batch {start//BATCH_SIZE + 1} saved to {OUTPUT_FILE}")

    print(f"\n✅ Done! All batches processed. File saved as {OUTPUT_FILE}")
