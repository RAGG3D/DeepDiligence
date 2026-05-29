import requests
import time
from collections import defaultdict

URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK0001738021.json"
HEADERS = {"User-Agent": "financial-research contact@example.com"}

YEARS = [2020, 2021, 2022, 2023, 2024]

time.sleep(0.3)
print("Fetching XBRL data...")
resp = requests.get(URL, headers=HEADERS)
resp.raise_for_status()
data = resp.json()

print(f"Company: {data.get('entityName', 'N/A')}")
print(f"CIK: {data.get('cik', 'N/A')}")
print()

us_gaap = data.get("facts", {}).get("us-gaap", {})
print(f"Total us-gaap concepts available: {len(us_gaap)}")
print("=" * 120)

# Structure: year -> list of (concept, value, type)
by_year = defaultdict(list)

for concept_name, concept_data in sorted(us_gaap.items()):
    units = concept_data.get("units", {})
    # Check all unit types (USD, shares, USD/shares, pure, etc.)
    for unit_type, entries in units.items():
        for entry in entries:
            form = entry.get("form", "")
            if form != "10-K":
                continue

            end_date = entry.get("end", "")
            start_date = entry.get("start", "")
            fp = entry.get("fp", "")
            val = entry.get("val")

            # Determine year from end date
            if not end_date:
                continue

            # Parse year and month from end date
            parts = end_date.split("-")
            if len(parts) < 2:
                continue
            end_year = int(parts[0])
            end_month = int(parts[1])

            # We want fiscal year end in December of target years
            if end_month != 12:
                continue
            if end_year not in YEARS:
                continue

            # Determine if flow (has start date, fp=FY) or instant (balance sheet)
            if start_date:
                entry_type = f"FLOW (fp={fp}, {start_date} to {end_date})"
            else:
                entry_type = f"INSTANT (fp={fp}, as of {end_date})"

            by_year[end_year].append((concept_name, val, entry_type, unit_type))

# Print grouped by year
for year in YEARS:
    entries = by_year.get(year, [])
    print(f"\n{'#' * 120}")
    print(f"# YEAR {year} — {len(entries)} data points")
    print(f"{'#' * 120}")

    if not entries:
        print("  (no data)")
        continue

    # Sort by concept name for readability
    entries.sort(key=lambda x: (x[0], x[3]))

    # Group by concept to show all units together
    current_concept = None
    for concept_name, val, entry_type, unit_type in entries:
        if concept_name != current_concept:
            if current_concept is not None:
                print()
            current_concept = concept_name
            print(f"  {concept_name}")

        # Format value
        if isinstance(val, (int, float)):
            if unit_type == "USD" and abs(val) >= 1000:
                formatted = f"${val:,.0f}"
            elif unit_type == "USD":
                formatted = f"${val:,.2f}"
            elif unit_type == "shares":
                formatted = f"{val:,.0f} shares"
            elif unit_type == "USD/shares":
                formatted = f"${val:.4f}/share"
            elif unit_type == "pure":
                formatted = f"{val}"
            else:
                formatted = f"{val} [{unit_type}]"
        else:
            formatted = str(val)

        print(f"    [{unit_type:12s}] {formatted:>30s}  |  {entry_type}")

# Summary
print(f"\n{'=' * 120}")
print("SUMMARY")
print(f"{'=' * 120}")
for year in YEARS:
    entries = by_year.get(year, [])
    unique_concepts = len(set(e[0] for e in entries))
    print(f"  {year}: {len(entries)} data points across {unique_concepts} unique concepts")

