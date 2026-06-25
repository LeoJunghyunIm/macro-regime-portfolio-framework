import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.append(str(PROJECT_ROOT))


from fetch_data import load_config, fetch_all_series, align_to_weekly_friday

config = load_config("config.yaml")

raw_data = fetch_all_series(config)
weekly_data = align_to_weekly_friday(raw_data)

latest_raw_date = raw_data.index.max()
latest_weekly_date = weekly_data.index.max()

print()
print("Completed-week alignment test")
print("Latest raw date:", latest_raw_date.date())
print("Latest weekly AsOfDate:", latest_weekly_date.date())
print()
print("Latest weekly row:")
print(weekly_data.tail(1).to_string())

if latest_weekly_date > latest_raw_date:
    raise ValueError("ERROR: Weekly AsOfDate is later than latest raw data date.")

print()
print("PASS: Weekly AsOfDate is not later than latest raw data date.")
