import pandas as pd
from dash import Dash, dcc, html, Input, Output, State, ALL, ctx, no_update
import plotly.express as px
import plotly.graph_objects as go
import os

print("🚀 Starting dashboard...")

# --- LOAD DATA ---
# Supports local development and deployed environments
_BASE = os.path.dirname(os.path.abspath(__file__))
_CSV_LOCAL   = r"E:\Law School Application\Scholarship data\ALL_SCHOOLS_COMBINED.csv"
_EXCEL_LOCAL = r"E:\Law School Application\Scholarship data\ALL_SCHOOLS_COMBINED.xlsx"
_CSV_DEPLOY  = os.path.join(_BASE, "data", "ALL_SCHOOLS_COMBINED.csv")

if os.path.exists(_CSV_DEPLOY):
    df = pd.read_csv(_CSV_DEPLOY)
    print(f"  Loaded deployed CSV: {len(df):,} rows")
elif os.path.exists(_CSV_LOCAL):
    df = pd.read_csv(_CSV_LOCAL)
    print(f"  Loaded local CSV: {len(df):,} rows")
else:
    try:
        df = pd.read_excel(_EXCEL_LOCAL, engine="openpyxl")
        print(f"  Loaded local Excel: {len(df):,} rows")
    except Exception as e:
        raise FileNotFoundError(
            f"Could not find data file. Expected at {_CSV_DEPLOY} for deployment, "
            f"or {_CSV_LOCAL} locally."
        ) from e

# FIX 1: Lowercase columns FIRST, then create school_slug
df.columns = df.columns.str.lower()
df['school_slug'] = df['school']

# Remap old slugs in scatter data to match consolidated file slugs

# FIX 2: Use df['col'] directly instead of df.get('col')
df['lsat'] = pd.to_numeric(df['lsat'], errors='coerce')
df['gpa'] = pd.to_numeric(df['gpa'], errors='coerce')
df['scholarship'] = pd.to_numeric(df['scholarship'], errors='coerce')

df = df.dropna(subset=['lsat', 'gpa'])

# --- DISPLAY NAMES ---
slug_map = {
    "university-of-akron-school-of-law": "Akron",
    "albany-law-school": "Albany",
    "american-university-law-school": "American",
    "asu-law": "Arizona State",
    "boston-college-law-school": "Boston College",
    "boston-university-law-school": "Boston University",
    "byu-law-school": "BYU",
    "brooklyn-law-school": "Brooklyn",
    "uc-berkeley-law-school": "UC Berkeley",
    "uc-davis-law-school": "UC Davis",
    "uc-irvine-law": "UC Irvine",
    "ucla-law": "UC Los Angeles",
    "uc-law-san-francisco": "UCSF (Hastings)",
    "cardozo-school-of-law": "Cardozo",
    "case-western-reserve-university-school-of-law": "Case Western Reserve",
    "chapman-university-fowler-school-of-law": "Chapman",
    "chicago-law-school": "Uchicago",
    "colorado-law": "Colorado",
    "columbia-law-school": "Columbia",
    "uconn-law-school": "Connecticut",
    "cornell-law-school": "Cornell",
    "university-of-dayton-school-of-law": "Dayton",
    "denver-law-school": "Denver",
    "drexel-university-thomas-r-kline-school-of-law": "Drexel",
    "duke-law": "Duke",
    "emory-law-school": "Emory",
    "fordham-law-school": "Fordham",
    "antonin-scalia-law-school": "George Mason",
    "gw-law-school": "George Washington",
    "georgetown-law": "Georgetown",
    "uga-law-school": "Georgia",
    "harvard-law-school": "Harvard",
    "university-of-idaho-college-of-law": "Idaho",
    "iowa-law": "Iowa",
    "loyola-law-school-los-angeles": "LMU LA",
    "marquette-university-law-school": "Marquette",
    "michigan-law-school": "Michigan",
    "university-of-minnesota-law-school": "Minnesota",
    "nyu-law-school": "NYU",
    "unc-law-school": "North Carolina",
    "northeastern-law-school": "Northeastern",
    "northwestern-law-school": "Northwestern",
    "notre-dame-law-school": "Notre Dame",
    "penn-law": "Upenn",
    "university-of-richmond-school-of-law": "Richmond",
    "roger-williams-university-school-of-law": "Roger Williams",
    "usd-law-school": "San Diego",
    "santa-clara-university-school-of-law": "Santa Clara",
    "university-of-south-carolina-school-of-law": "South Carolina",
    "south-texas-college-of-law-houston": "South Texas (Houston)",
    "usc-gould-school-of-law": "USC",
    "smu-dedman-school-of-law": "SMU",
    "stanford-law-school": "Stanford",
    "texas-am-law-school": "Texas A&M",
    "ut-austin-law-school": "Texas",
    "moritz-college-of-law": "Ohio State",
    "vanderbilt-law-school": "Vanderbilt",
    "uva-law-school": "UVA",
    "wake-forest-law-school": "Wake Forest",
    "washburn-university-school-of-law": "Washburn",
    "washington-and-lee-law-school": "Washington and Lee",
    "washington-university-school-of-law": "WashU",
    "willamette-university-college-of-law": "Willamette University",
    "william-and-mary-law-school": "William & Mary",
    "wisconsin-law-school": "Wisconsin",
    "yale-law-school": "Yale",
}

medians = {
    "stanford-law-school":                  {"lsat":173,"gpa":3.96,"rank":1,  "lsat25":171,"lsat75":176,"gpa25":3.87,"gpa75":4.00},
    "yale-law-school":                       {"lsat":174,"gpa":3.96,"rank":2,  "lsat25":171,"lsat75":177,"gpa25":3.90,"gpa75":4.00},
    "chicago-law-school":                    {"lsat":174,"gpa":3.97,"rank":2,  "lsat25":171,"lsat75":176,"gpa25":3.87,"gpa75":4.00},
    "uva-law-school":                        {"lsat":173,"gpa":3.99,"rank":4,  "lsat25":168,"lsat75":175,"gpa25":3.83,"gpa75":4.04},
    "penn-law":                              {"lsat":173,"gpa":3.95,"rank":4,  "lsat25":167,"lsat75":174,"gpa25":3.77,"gpa75":4.00},
    "harvard-law-school":                    {"lsat":174,"gpa":3.96,"rank":6,  "lsat25":171,"lsat75":176,"gpa25":3.89,"gpa75":4.00},
    "duke-law":                              {"lsat":171,"gpa":3.91,"rank":7,  "lsat25":169,"lsat75":172,"gpa25":3.83,"gpa75":3.96},
    "nyu-law-school":                        {"lsat":172,"gpa":3.92,"rank":7,  "lsat25":169,"lsat75":174,"gpa25":3.81,"gpa75":3.97},
    "michigan-law-school":                   {"lsat":171,"gpa":3.88,"rank":9,  "lsat25":168,"lsat75":173,"gpa25":3.74,"gpa75":3.95},
    "columbia-law-school":                   {"lsat":173,"gpa":3.92,"rank":9,  "lsat25":169,"lsat75":175,"gpa25":3.85,"gpa75":3.98},
    "northwestern-law-school":               {"lsat":173,"gpa":3.96,"rank":9,  "lsat25":167,"lsat75":175,"gpa25":3.76,"gpa75":4.00},
    "vanderbilt-law-school":                 {"lsat":170,"gpa":3.91,"rank":12, "lsat25":167,"lsat75":171,"gpa25":3.77,"gpa75":3.97},
    "ucla-law":                              {"lsat":171,"gpa":3.95,"rank":13, "lsat25":166,"lsat75":172,"gpa25":3.73,"gpa75":4.00},
    "cornell-law-school":                    {"lsat":173,"gpa":3.92,"rank":13, "lsat25":168,"lsat75":175,"gpa25":3.75,"gpa75":3.97},
    "washington-university-school-of-law":   {"lsat":175,"gpa":3.96,"rank":13, "lsat25":165,"lsat75":176,"gpa25":3.58,"gpa75":4.00},
    "uc-berkeley-law-school":                {"lsat":170,"gpa":3.92,"rank":16, "lsat25":167,"lsat75":172,"gpa25":3.84,"gpa75":3.99},
    "ut-austin-law-school":                  {"lsat":172,"gpa":3.89,"rank":16, "lsat25":166,"lsat75":173,"gpa25":3.75,"gpa75":3.96},
    "georgetown-law":                        {"lsat":171,"gpa":3.93,"rank":18, "lsat25":166,"lsat75":173,"gpa25":3.75,"gpa75":3.98},
    "unc-law-school":                        {"lsat":168,"gpa":3.89,"rank":18, "lsat25":165,"lsat75":169,"gpa25":3.78,"gpa75":3.97},
    "notre-dame-law-school":                 {"lsat":170,"gpa":3.89,"rank":20, "lsat25":166,"lsat75":171,"gpa25":3.78,"gpa75":3.95},
    "boston-college-law-school":             {"lsat":168,"gpa":3.83,"rank":20, "lsat25":162,"lsat75":169,"gpa25":3.67,"gpa75":3.90},
    "texas-am-law-school":                   {"lsat":169,"gpa":4.00,"rank":22, "lsat25":161,"lsat75":170,"gpa25":3.67,"gpa75":4.00},
    "university-of-minnesota-law-school":    {"lsat":171,"gpa":3.88,"rank":22, "lsat25":166,"lsat75":173,"gpa25":3.62,"gpa75":3.95},
    "boston-university-law-school":          {"lsat":170,"gpa":3.88,"rank":24, "lsat25":164,"lsat75":171,"gpa25":3.71,"gpa75":3.93},
    "byu-law-school":                        {"lsat":170,"gpa":3.95,"rank":24, "lsat25":168,"lsat75":172,"gpa25":3.82,"gpa75":3.98},
    "usc-gould-school-of-law":               {"lsat":169,"gpa":3.91,"rank":26, "lsat25":165,"lsat75":170,"gpa25":3.73,"gpa75":3.97},
    "gw-law-school":                         {"lsat":168,"gpa":3.86,"rank":26, "lsat25":162,"lsat75":170,"gpa25":3.55,"gpa75":3.93},
    "uga-law-school":                        {"lsat":169,"gpa":3.92,"rank":26, "lsat25":159,"lsat75":171,"gpa25":3.70,"gpa75":3.97},
    "wisconsin-law-school":                  {"lsat":167,"gpa":3.81,"rank":26, "lsat25":161,"lsat75":169,"gpa25":3.61,"gpa75":3.93},
    "wake-forest-law-school":                {"lsat":166,"gpa":3.79,"rank":30, "lsat25":163,"lsat75":168,"gpa25":3.60,"gpa75":3.90},
    "moritz-college-of-law":                 {"lsat":168,"gpa":3.91,"rank":30, "lsat25":163,"lsat75":169,"gpa25":3.64,"gpa75":3.97},
    "antonin-scalia-law-school":             {"lsat":169,"gpa":3.93,"rank":32, "lsat25":162,"lsat75":171,"gpa25":3.55,"gpa75":3.98},
    "iowa-law":                              {"lsat":164,"gpa":3.78,"rank":32, "lsat25":161,"lsat75":165,"gpa25":3.62,"gpa75":3.92},
    "uc-irvine-law":                         {"lsat":169,"gpa":3.80,"rank":34, "lsat25":166,"lsat75":170,"gpa25":3.59,"gpa75":3.90},
    "william-and-mary-law-school":           {"lsat":166,"gpa":3.82,"rank":34, "lsat25":161,"lsat75":167,"gpa25":3.53,"gpa75":3.94},
    "washington-and-lee-law-school":         {"lsat":167,"gpa":3.75,"rank":34, "lsat25":161,"lsat75":168,"gpa25":3.51,"gpa75":3.85},
    "emory-law-school":                      {"lsat":166,"gpa":3.82,"rank":40, "lsat25":162,"lsat75":168,"gpa25":3.68,"gpa75":3.87},
    "asu-law":                               {"lsat":165,"gpa":3.91,"rank":44, "lsat25":158,"lsat75":167,"gpa25":3.58,"gpa75":3.98},
    "uc-davis-law-school":                   {"lsat":165,"gpa":3.70,"rank":52, "lsat25":160,"lsat75":167,"gpa25":3.49,"gpa75":3.87},
    "colorado-law":                          {"lsat":164,"gpa":3.81,"rank":54, "lsat25":161,"lsat75":167,"gpa25":3.62,"gpa75":3.91},
    "usd-law-school":                        {"lsat":163,"gpa":3.84,"rank":54, "lsat25":160,"lsat75":165,"gpa25":3.59,"gpa75":3.92},
    "uconn-law-school":                      {"lsat":162,"gpa":3.78,"rank":58, "lsat25":160,"lsat75":164,"gpa25":3.55,"gpa75":3.90},
    "cardozo-school-of-law":                 {"lsat":165,"gpa":3.78,"rank":59, "lsat25":162,"lsat75":167,"gpa25":3.56,"gpa75":3.87},
    "university-of-arizona-law-school":      {"lsat":163,"gpa":3.78,"rank":59, "lsat25":162,"lsat75":165,"gpa25":3.54,"gpa75":3.90},
    "northeastern-law-school":               {"lsat":164,"gpa":3.77,"rank":68, "lsat25":161,"lsat75":166,"gpa25":3.61,"gpa75":3.86},
    "loyola-law-school-los-angeles":         {"lsat":163,"gpa":3.74,"rank":70, "lsat25":160,"lsat75":165,"gpa25":3.52,"gpa75":3.89},
    "uc-law-san-francisco":                  {"lsat":161,"gpa":3.71,"rank":85, "lsat25":158,"lsat75":163,"gpa25":3.55,"gpa75":3.83},
    "denver-law-school":                     {"lsat":160,"gpa":3.66,"rank":91, "lsat25":156,"lsat75":162,"gpa25":3.36,"gpa75":3.80},
    "american-university-law-school":        {"lsat":162,"gpa":3.63,"rank":108,"lsat25":157,"lsat75":163,"gpa25":3.32,"gpa75":3.75},
    "chapman-university-fowler-school-of-law":{"lsat":163,"gpa":3.70,"rank":120,"lsat25":156,"lsat75":164,"gpa25":3.32,"gpa75":3.82},
}

# --- SCHOOL INFO (location + 509 tuition/fees/living, all annual figures) ---
# tuition/fees/living are annual; coa_3yr = (tuition+fees+living)*3
school_info = {
    "university-of-akron-school-of-law": {"city": "", "tuition": 22359, "fees": 4390, "tuition_res": 22359, "fees_res": 4290, "living": 5500},
    "university-of-alabama-school-of-law": {"city": "AL", "tuition": 48100, "fees": 899, "tuition_res": 24980, "fees_res": 899, "living": 22620},
    "albany-law-school": {"city": "", "tuition": 63994, "fees": 240, "tuition_res": 63994, "fees_res": 240, "living": 21844},
    "american-university-law-school": {"city": "", "tuition": 66990, "fees": 2182, "tuition_res": 66990, "fees_res": 2182, "living": 33832},
    "appalachian-school-of-law": {"city": "", "tuition": 41000, "fees": None, "tuition_res": 41000, "fees_res": None, "living": 41800},
    "asu-law": {"city": "AZ", "tuition": 52099, "fees": 1443, "tuition_res": 28390, "fees_res": 1443, "living": 33090},
    "university-of-arizona-law-school": {"city": "AZ", "tuition": 31360, "fees": 122, "tuition_res": 26490, "fees_res": 122, "living": 25600},
    "arkansas-little-rock-law": {"city": "", "tuition": None, "fees": None, "tuition_res": 13770, "fees_res": 3704, "living": 25459},
    "university-of-arkansas-school-of-law": {"city": "", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 22194, "credit_res": 541, "credit_oos": 1304},
    "john-marshall-atlanta": {"city": "", "tuition": 52536, "fees": 1050, "tuition_res": 52536, "fees_res": 1050, "living": 35426},
    "ave-maria-school-of-law": {"city": "", "tuition": 49500, "fees": 2850, "tuition_res": 49500, "fees_res": 2850, "living": 32765},
    "university-of-baltimore-school-of-law": {"city": "MD", "tuition": 50624, "fees": 2548, "tuition_res": 33944, "fees_res": 2474, "living": 25024},
    "barry-university-law": {"city": "", "tuition": 42000, "fees": 1150, "tuition_res": 42000, "fees_res": 1150, "living": 32008},
    "baylor-law-school": {"city": "TX", "tuition": 67452, "fees": 95, "tuition_res": None, "fees_res": None, "living": 25023, "credit_res": None, "credit_oos": None},
    "belmont-university-law": {"city": "", "tuition": 55500, "fees": 2205, "tuition_res": 55500, "fees_res": 2205, "living": 32026},
    "boston-college-law-school": {"city": "MA", "tuition": 72380, "fees": 375, "tuition_res": 72380, "fees_res": 375, "living": 25726},
    "boston-university-law-school": {"city": "MA", "tuition": 69870, "fees": 1512, "tuition_res": 69870, "fees_res": 1512, "living": 21532},
    "byu-law-school": {"city": "UT", "tuition": 31984, "fees": None, "tuition_res": 15992, "fees_res": None, "living": 22436},
    "brooklyn-law-school": {"city": "NY", "tuition": 75496, "fees": 1450, "tuition_res": 75496, "fees_res": 1450, "living": 36528},
    "university-of-buffalo-law": {"city": "NY", "tuition": 32100, "fees": 3320, "tuition_res": 26430, "fees_res": 3320, "living": 24814},
    "california-western-school-of-law": {"city": "CA", "tuition": 64720, "fees": 150, "tuition_res": 64720, "fees_res": 150, "living": 34342},
    "uc-berkeley-law-school": {"city": "CA", "tuition": 76149, "fees": 3842, "tuition_res": 62532, "fees_res": 3842, "living": 45866},
    "uc-davis-law-school": {"city": "CA", "tuition": 69705, "fees": 2410, "tuition_res": 57460, "fees_res": 2410, "living": 33048},
    "uc-irvine-law": {"city": "CA", "tuition": 71325, "fees": 2208, "tuition_res": 59080, "fees_res": 2208, "living": 37293},
    "ucla-law": {"city": "CA", "tuition": 71329, "fees": 2778, "tuition_res": 59084, "fees_res": 2778, "living": 40046},
    "uc-law-san-francisco": {"city": "CA", "tuition": 61773, "fees": 1242, "tuition_res": 53087, "fees_res": 1242, "living": 34466},
    "campbell-university-law": {"city": "", "tuition": 51840, "fees": 930, "tuition_res": 51840, "fees_res": 930, "living": 35281},
    "capital-university-law": {"city": "", "tuition": 45066, "fees": 390, "tuition_res": 45066, "fees_res": 390, "living": 21112},
    "cardozo-school-of-law": {"city": "NY", "tuition": 74438, "fees": 2020, "tuition_res": 74438, "fees_res": 2020, "living": 34114},
    "case-western-reserve-university-school-of-law": {"city": "", "tuition": 64600, "fees": 952, "tuition_res": 64600, "fees_res": 952, "living": 32008},
    "catholic-university-law": {"city": "DC", "tuition": 61580, "fees": 1700, "tuition_res": 61580, "fees_res": 1700, "living": 32374},
    "chapman-university-fowler-school-of-law": {"city": "CA", "tuition": 63126, "fees": 75, "tuition_res": 63126, "fees_res": 75, "living": 41778},
    "charleston-school-of-law": {"city": "SC", "tuition": 48600, "fees": 1084, "tuition_res": 48600, "fees_res": 1084, "living": 27752},
    "chicago-law-school": {"city": "IL", "tuition": 83316, "fees": 1590, "tuition_res": 83316, "fees_res": 1590, "living": 35418},
    "chicago-kent-college-of-law": {"city": "IL", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 27964, "credit_res": 1830, "credit_oos": 1830},
    "university-of-cincinnati-college-of-law": {"city": "OH", "tuition": 29010, "fees": 800, "tuition_res": 24010, "fees_res": 800, "living": 25956},
    "cuny-law-school": {"city": "NY", "tuition": 25640, "fees": 563, "tuition_res": 15450, "fees_res": 563, "living": 30805},
    "cleveland-state-law": {"city": "OH", "tuition": 29866, "fees": 1800, "tuition_res": 29766, "fees_res": 1800, "living": 21870},
    "colorado-law": {"city": "CO", "tuition": 44002, "fees": 1333, "tuition_res": 35992, "fees_res": 1333, "living": 27626},
    "columbia-law-school": {"city": "NY", "tuition": 85368, "fees": 3022, "tuition_res": 85368, "fees_res": 3022, "living": 31554},
    "uconn-law-school": {"city": "CT", "tuition": 61396, "fees": 1246, "tuition_res": 30354, "fees_res": 1246, "living": 25860},
    "cooley-law-school": {"city": "MI", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 29972, "credit_res": 1580, "credit_oos": 1580},
    "cornell-law-school": {"city": "NY", "tuition": 84722, "fees": 690, "tuition_res": 84722, "fees_res": 690, "living": 32952},
    "creighton-university-law": {"city": "", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 23400, "credit_res": 1542, "credit_oos": 1542},
    "university-of-dayton-school-of-law": {"city": "", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 19395, "credit_res": 1380, "credit_oos": 1380},
    "denver-law-school": {"city": "CO", "tuition": 63390, "fees": 422, "tuition_res": 63390, "fees_res": 422, "living": 27260},
    "depaul-university-law": {"city": "IL", "tuition": 55328, "fees": 300, "tuition_res": 55328, "fees_res": 300, "living": 29110},
    "university-of-detroit-mercy-school-of-law": {"city": "MI", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 23983, "credit_res": 1616, "credit_oos": 1616},
    "udc-law-school": {"city": "DC", "tuition": 24874, "fees": 1000, "tuition_res": 12438, "fees_res": 1000, "living": 34035},
    "drake-university-law": {"city": "", "tuition": 50920, "fees": 470, "tuition_res": 50920, "fees_res": 470, "living": 20923},
    "drexel-university-thomas-r-kline-school-of-law": {"city": "PA", "tuition": 57000, "fees": 1850, "tuition_res": 57000, "fees_res": 1850, "living": 28022},
    "duke-law": {"city": "NC", "tuition": 80100, "fees": 1729, "tuition_res": 80100, "fees_res": 1729, "living": 26116},
    "duquesne-university-law": {"city": "", "tuition": 58898, "fees": 1220, "tuition_res": 58898, "fees_res": 1220, "living": 24318},
    "elon-university-law": {"city": "", "tuition": 45333, "fees": None, "tuition_res": 45333, "fees_res": None, "living": 31000},
    "emory-law-school": {"city": "", "tuition": 69510, "fees": 962, "tuition_res": 69510, "fees_res": 962, "living": 38990},
    "faulkner-university-law": {"city": "", "tuition": 39900, "fees": 500, "tuition_res": 39900, "fees_res": 500, "living": 32600},
    "florida-am-law": {"city": "Orlando, FL", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 39030, "credit_res": 456, "credit_oos": 1098},
    "florida-international-university-college-of-law": {"city": "FL", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 33176, "credit_res": 691, "credit_oos": 1196},
    "florida-state-university-college-of-law": {"city": "FL", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 23180, "credit_res": 688, "credit_oos": 1422},
    "university-of-florida-levin-college-of-law": {"city": "FL", "tuition": 36148, "fees": 3514, "tuition_res": 19139, "fees_res": 2665, "living": 26646},
    "fordham-law-school": {"city": "NY", "tuition": 78078, "fees": 1170, "tuition_res": 78078, "fees_res": 1170, "living": 37354},
    "antonin-scalia-law-school": {"city": "VA", "tuition": 40978, "fees": 2968, "tuition_res": 24864, "fees_res": 2968, "living": 38342},
    "gw-law-school": {"city": "DC", "tuition": 75420, "fees": None, "tuition_res": 75420, "fees_res": None, "living": 25240},
    "georgetown-law": {"city": "DC", "tuition": 83576, "fees": None, "tuition_res": 83576, "fees_res": None, "living": 33905},
    "georgia-state-university-college-of-law": {"city": "GA", "tuition": 37886, "fees": 1434, "tuition_res": 16162, "fees_res": 1434, "living": 26693},
    "uga-law-school": {"city": "GA", "tuition": 38944, "fees": 1458, "tuition_res": 18044, "fees_res": 1458, "living": 25978},
    "gonzaga-university-law": {"city": "", "tuition": 57962, "fees": 825, "tuition_res": 57962, "fees_res": 825, "living": 17105},
    "harvard-law-school": {"city": "MA", "tuition": 77100, "fees": 1592, "tuition_res": 77100, "fees_res": 1592, "living": 38690},
    "university-of-hawaii-william-s-richardson-school-of-law": {"city": "HI", "tuition": 46728, "fees": 772, "tuition_res": 23304, "fees_res": 772, "living": 27322},
    "hofstra-university-school-of-law": {"city": "", "tuition": 73344, "fees": 1516, "tuition_res": 73344, "fees_res": 1516, "living": 36009},
    "university-of-houston-law-center": {"city": "TX", "tuition": 46051, "fees": 7896, "tuition_res": 31326, "fees_res": 7896, "living": 26086},
    "howard-university-school-of-law": {"city": "", "tuition": 39780, "fees": 1632, "tuition_res": 39780, "fees_res": 1632, "living": 38251},
    "university-of-idaho-college-of-law": {"city": "", "tuition": 46752, "fees": None, "tuition_res": 27516, "fees_res": None, "living": 20136},
    "university-of-illinois-chicago-school-of-law": {"city": "IL", "tuition": 46500, "fees": 3404, "tuition_res": 36500, "fees_res": 3404, "living": 20620},
    "uic-john-marshall-law-school": {"city": "IL", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 24976, "credit_res": 1248, "credit_oos": 1560},
    "indiana-university-maurer-school-of-law": {"city": "IN", "tuition": 58000, "fees": 1522, "tuition_res": 36950, "fees_res": 1522, "living": 25822},
    "indiana-university-indianapolis-law": {"city": "", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 23512, "credit_res": 1028, "credit_oos": 1655},
    "inter-american-university-pr-law": {"city": "PR", "tuition": 16585, "fees": 1175, "tuition_res": 16585, "fees_res": 1175, "living": 25286},
    "iowa-law": {"city": "IO", "tuition": 52191, "fees": 2545, "tuition_res": 30944, "fees_res": 2545, "living": 23920},
    "jacksonville-university-law": {"city": "", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 24911, "credit_res": 1200, "credit_oos": 1200},
    "kansas-law": {"city": "KS", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 20238, "credit_res": 857, "credit_oos": 1094},
    "university-of-kentucky-college-of-law": {"city": "", "tuition": 49836, "fees": 1616, "tuition_res": 25685, "fees_res": 1616, "living": 28425},
    "lewis-and-clark-law-school": {"city": "", "tuition": 61434, "fees": 50, "tuition_res": 61434, "fees_res": 50, "living": 28370},
    "liberty-university-law": {"city": "", "tuition": 43150, "fees": 2198, "tuition_res": 43150, "fees_res": 2198, "living": 26532},
    "lincoln-memorial-university-law": {"city": "", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 28325, "credit_res": 1490, "credit_oos": 1490},
    "louisiana-state-university-law": {"city": "LA", "tuition": 19750, "fees": 19355, "tuition_res": 19750, "fees_res": 3915, "living": 27034},
    "university-of-louisville-louis-d-brandeis-school-of-law": {"city": "KY", "tuition": 32000, "fees": 1396, "tuition_res": 27000, "fees_res": 1396, "living": 26966},
    "loyola-law-school-los-angeles": {"city": "CA", "tuition": 70360, "fees": 1026, "tuition_res": 70360, "fees_res": 1026, "living": 43030},
    "loyola-university-chicago-school-of-law": {"city": "IL", "tuition": 58670, "fees": 996, "tuition_res": 58670, "fees_res": 996, "living": 32658},
    "loyola-university-new-orleans-college-of-law": {"city": "LA", "tuition": 55024, "fees": 1870, "tuition_res": 55024, "fees_res": 1870, "living": 29205},
    "university-of-maine-school-of-law": {"city": "", "tuition": 38160, "fees": 4060, "tuition_res": 26100, "fees_res": 4060, "living": 26214},
    "marquette-university-law-school": {"city": "", "tuition": 51970, "fees": None, "tuition_res": 51970, "fees_res": None, "living": 22146},
    "university-of-maryland-carey-school-of-law": {"city": "MD", "tuition": 55844, "fees": 2179, "tuition_res": 37830, "fees_res": 2179, "living": 32842},
    "umass-dartmouth-law": {"city": "MA", "tuition": 41565, "fees": 1690, "tuition_res": 31766, "fees_res": 1690, "living": 34626},
    "university-of-memphis-school-of-law": {"city": "TN", "tuition": 24354, "fees": 2594, "tuition_res": 18634, "fees_res": 2594, "living": 28729},
    "mercer-university-school-of-law": {"city": "", "tuition": 44782, "fees": 300, "tuition_res": 44782, "fees_res": 300, "living": 23289},
    "university-of-miami-law": {"city": "FL", "tuition": 66720, "fees": 2498, "tuition_res": 66720, "fees_res": 2498, "living": 43924},
    "michigan-state-university-college-of-law": {"city": "MI", "tuition": 48222, "fees": 388, "tuition_res": 43480, "fees_res": 388, "living": 27152},
    "michigan-law-school": {"city": "MI", "tuition": 79108, "fees": 572, "tuition_res": 76108, "fees_res": 572, "living": 26886},
    "university-of-minnesota-law-school": {"city": "MN", "tuition": 64524, "fees": 0, "tuition_res": 54120, "fees_res": 0, "living": 19108},
    "mississippi-college-school-of-law": {"city": "MS", "tuition": 37980, "fees": 1662, "tuition_res": 37980, "fees_res": 1662, "living": 26100},
    "university-of-missouri-school-of-law": {"city": "MO", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 22002, "credit_res": 812, "credit_oos": 1028},
    "umkc-law-school": {"city": "MO", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 23924, "credit_res": 786, "credit_oos": 986},
    "william-mitchell-college-of-law": {"city": "", "tuition": 55300, "fees": 1000, "tuition_res": 55300, "fees_res": 1000, "living": 16848},
    "university-of-montana-school-of-law": {"city": "", "tuition": 51582, "fees": 3502, "tuition_res": 24582, "fees_res": 2930, "living": 35371},
    "university-of-nebraska-college-of-law": {"city": "NE", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 22092, "credit_res": 456, "credit_oos": 1229},
    "university-of-nevada-las-vegas-school-of-law": {"city": "NV", "tuition": 46040, "fees": 1610, "tuition_res": 31518, "fees_res": 1610, "living": 31263},
    "new-england-law-boston": {"city": "MA", "tuition": 62896, "fees": 320, "tuition_res": 62896, "fees_res": 320, "living": 48437},
    "university-of-new-hampshire-school-of-law": {"city": "NH", "tuition": 50000, "fees": 348, "tuition_res": 42000, "fees_res": 348, "living": 27094},
    "university-of-new-mexico-school-of-law": {"city": "NM", "tuition": 43192, "fees": 1546, "tuition_res": 19474, "fees_res": 1546, "living": 23114},
    "new-york-law-school": {"city": "NY", "tuition": 71052, "fees": 2472, "tuition_res": 71052, "fees_res": 2472, "living": 23759},
    "nyu-law-school": {"city": "NY", "tuition": 83952, "fees": 3262, "tuition_res": 83952, "fees_res": 3262, "living": 36094},
    "north-carolina-central-law": {"city": "NC", "tuition": 36116, "fees": 5746, "tuition_res": 13444, "fees_res": 5746, "living": 28292},
    "unc-law-school": {"city": "NC", "tuition": 51320, "fees": 3480, "tuition_res": 28082, "fees_res": 3480, "living": 25678},
    "university-of-north-dakota-school-of-law": {"city": "ND", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 20616, "credit_res": 591, "credit_oos": 1181},
    "unt-dallas-law": {"city": "TX", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 24638, "credit_res": 562, "credit_oos": 977},
    "northeastern-law-school": {"city": "MA", "tuition": 65652, "fees": 198, "tuition_res": 65652, "fees_res": 198, "living": 22900},
    "northern-illinois-university-law": {"city": "", "tuition": 23667, "fees": 400, "tuition_res": 23667, "fees_res": 400, "living": 22742},
    "northern-kentucky-university-law": {"city": "KY", "tuition": 38454, "fees": 1004, "tuition_res": 24388, "fees_res": 1004, "living": 20932},
    "northwestern-law-school": {"city": "IL", "tuition": 79722, "fees": 861, "tuition_res": 79772, "fees_res": 861, "living": 28220},
    "notre-dame-law-school": {"city": "PA", "tuition": 73430, "fees": 520, "tuition_res": 73430, "fees_res": 520, "living": 25990},
    "nova-southeastern-law": {"city": "", "tuition": 49350, "fees": 1360, "tuition_res": 49350, "fees_res": 1360, "living": 40413},
    "ohio-northern-university-law": {"city": "OH", "tuition": 33900, "fees": None, "tuition_res": 33900, "fees_res": None, "living": 22022},
    "oklahoma-city-university-law": {"city": "OK", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 28305, "credit_res": 1107, "credit_oos": 1107},
    "oregon-law": {"city": "OR", "tuition": 59495, "fees": 1372, "tuition_res": 47240, "fees_res": 1372, "living": 22020},
    "pace-university-school-of-law": {"city": "", "tuition": 56648, "fees": 1108, "tuition_res": 56648, "fees_res": 1108, "living": 33422},
    "university-of-pacific-law": {"city": "", "tuition": 62678, "fees": 640, "tuition_res": 62678, "fees_res": 640, "living": 35081},
    "penn-state-dickinson-law": {"city": "PA", "tuition": 60048, "fees": 431, "tuition_res": 60048, "fees_res": 431, "living": 28094},
    "penn-law": {"city": "PA", "tuition": 78348, "fees": 6144, "tuition_res": 78348, "fees_res": 6144, "living": 31640},
    "pepperdine-university-school-of-law": {"city": "", "tuition": 72920, "fees": 630, "tuition_res": 72920, "fees_res": 630, "living": 33700},
    "pittsburgh-law-school": {"city": "PA", "tuition": 53434, "fees": 1697, "tuition_res": 39936, "fees_res": 1697, "living": 20346},
    "university-of-puerto-rico-law": {"city": "PR", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 23501, "credit_res": 305, "credit_oos": 305},
    "quinnipiac-university-school-of-law": {"city": "", "tuition": 56610, "fees": 1150, "tuition_res": 56610, "fees_res": 1150, "living": 29274},
    "regent-university-law": {"city": "", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 32268, "credit_res": 1400, "credit_oos": 1400},
    "university-of-richmond-school-of-law": {"city": "VA", "tuition": None, "fees": None, "tuition_res": 57700, "fees_res": None, "living": 26510},
    "roger-williams-university-school-of-law": {"city": "", "tuition": 47400, "fees": 1690, "tuition_res": 47400, "fees_res": 1690, "living": 25616},
    "rutgers-law-school": {"city": "PA", "tuition": 47863, "fees": 4030, "tuition_res": 31577, "fees_res": 4030, "living": 33190},
    "saint-louis-university-law": {"city": "MO", "tuition": 55260, "fees": 930, "tuition_res": 55260, "fees_res": 930, "living": 26112},
    "samford-university-cumberland-school-of-law": {"city": "", "tuition": 46334, "fees": 1000, "tuition_res": 46334, "fees_res": 1000, "living": 31226},
    "usd-law-school": {"city": "CA", "tuition": 66950, "fees": 801, "tuition_res": 66950, "fees_res": 801, "living": 39197},
    "university-of-san-francisco-law": {"city": "CA", "tuition": 59800, "fees": 190, "tuition_res": 59800, "fees_res": 190, "living": 34100},
    "santa-clara-university-school-of-law": {"city": "CA", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 35618, "credit_res": 2260, "credit_oos": 2260},
    "seattle-university-school-of-law": {"city": "WA", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 34161, "credit_res": 2061, "credit_oos": 2061},
    "seton-hall-university-school-of-law": {"city": "", "tuition": 69660, "fees": 1890, "tuition_res": 69660, "fees_res": 1890, "living": 28548},
    "university-of-south-carolina-school-of-law": {"city": "SC", "tuition": 38100, "fees": 3700, "tuition_res": 20322, "fees_res": 3700, "living": 25392},
    "university-of-south-dakota-school-of-law": {"city": "SD", "tuition": 32184, "fees": 5037, "tuition_res": 12076, "fees_res": 5037, "living": 18034},
    "south-texas-college-of-law-houston": {"city": "TX", "tuition": 42960, "fees": 2000, "tuition_res": 42960, "fees_res": 2000, "living": 56958},
    "usc-gould-school-of-law": {"city": "CA", "tuition": 84034, "fees": 1914, "tuition_res": 84034, "fees_res": 1914, "living": 32298},
    "southern-illinois-university-law": {"city": "", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 18608, "credit_res": 547, "credit_oos": 547},
    "smu-dedman-school-of-law": {"city": "", "tuition": 60844, "fees": 7842, "tuition_res": 60844, "fees_res": 7842, "living": 17834},
    "southern-university-law": {"city": "", "tuition": 11338, "fees": 22994, "tuition_res": 11338, "fees_res": 7494, "living": 32262},
    "southwestern-law-school": {"city": "", "tuition": 63160, "fees": None, "tuition_res": 63160, "fees_res": None, "living": 38033},
    "st-john-s-university-school-of-law": {"city": "", "tuition": 75170, "fees": 1572, "tuition_res": 75170, "fees_res": 1572, "living": 32728},
    "st-mary-s-university-school-of-law": {"city": "", "tuition": 44910, "fees": 1400, "tuition_res": 44910, "fees_res": 1400, "living": 23878},
    "st-thomas-university-miami-law": {"city": "FL", "tuition": 44730, "fees": 2848, "tuition_res": 44730, "fees_res": 2848, "living": 42548},
    "st-thomas-minneapolis-law": {"city": "", "tuition": 50998, "fees": 634, "tuition_res": 50998, "fees_res": 634, "living": 27408},
    "stanford-law-school": {"city": "CA", "tuition": None, "fees": None, "tuition_res": 77454, "fees_res": 2253, "living": 51102},
    "stetson-university-law": {"city": "", "tuition": 56521, "fees": 496, "tuition_res": 56521, "fees_res": 496, "living": 27414},
    "suffolk-university-law-school": {"city": "", "tuition": 62610, "fees": 1436, "tuition_res": 62610, "fees_res": 1436, "living": 34706},
    "syracuse-university-law": {"city": "NY", "tuition": 64650, "fees": 2290, "tuition_res": 64650, "fees_res": 2290, "living": 24472},
    "temple-university-beasley-school-of-law": {"city": "NY", "tuition": 48996, "fees": 1549, "tuition_res": 32588, "fees_res": 1549, "living": 27812},
    "university-of-tennessee-college-of-law": {"city": "TN", "tuition": 35140, "fees": 4182, "tuition_res": 16696, "fees_res": 3866, "living": 26216},
    "texas-am-law-school": {"city": "TX", "tuition": 48522, "fees": None, "tuition_res": 32634, "fees_res": None, "living": 29696},
    "texas-southern-university-law": {"city": "TX", "tuition": 26331, "fees": 1836, "tuition_res": 19202, "fees_res": 1836, "living": 27784},
    "texas-tech-university-school-of-law": {"city": "TX", "tuition": 33840, "fees": 4942, "tuition_res": 22590, "fees_res": 4942, "living": 20107},
    "ut-austin-law-school": {"city": "TX", "tuition": 56822, "fees": None, "tuition_res": 38236, "fees_res": None, "living": 25398},
    "moritz-college-of-law": {"city": "OH", "tuition": 50902, "fees": 1613, "tuition_res": 35650, "fees_res": 1613, "living": 23692},
    "university-of-toledo-college-of-law": {"city": "", "tuition": 26446, "fees": 4152, "tuition_res": 26446, "fees_res": 3152, "living": 24247},
    "touro-law-center": {"city": "", "tuition": 61000, "fees": 1150, "tuition_res": 61000, "fees_res": 1150, "living": 52602},
    "tulane-university-law-school": {"city": "", "tuition": None, "fees": None, "tuition_res": 69578, "fees_res": 5178, "living": 29670},
    "tulsa-university-law": {"city": "OK", "tuition": 30840, "fees": 1460, "tuition_res": 30840, "fees_res": 1460, "living": 29488},
    "university-of-utah-sjd": {"city": "UT", "tuition": 45382, "fees": 1108, "tuition_res": 34721, "fees_res": 1108, "living": 38078},
    "vanderbilt-law-school": {"city": "TN", "tuition": 76440, "fees": 2359, "tuition_res": 76440, "fees_res": 2359, "living": 34934},
    "vermont-law-school": {"city": "", "tuition": None, "fees": None, "tuition_res": 55382, "fees_res": 1075, "living": 23952},
    "villanova-university-charles-widger-school-of-law": {"city": "", "tuition": 59800, "fees": 950, "tuition_res": 59800, "fees_res": 950, "living": 25492},
    "uva-law-school": {"city": "VA", "tuition": 76396, "fees": 4504, "tuition_res": 74078, "fees_res": 3822, "living": 28770},
    "wake-forest-law-school": {"city": "NC", "tuition": 57920, "fees": 1224, "tuition_res": 57920, "fees_res": 1224, "living": 30639},
    "washburn-university-school-of-law": {"city": "", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 17018, "credit_res": 932, "credit_oos": 1393},
    "washington-and-lee-law-school": {"city": "", "tuition": 57450, "fees": 2490, "tuition_res": 57450, "fees_res": 2490, "living": 21890},
    "washington-university-school-of-law": {"city": "MO", "tuition": 72792, "fees": 892, "tuition_res": 72792, "fees_res": 892, "living": 28603},
    "university-of-washington-school-of-law": {"city": "WA", "tuition": 57810, "fees": 1146, "tuition_res": 45927, "fees_res": 1146, "living": 29082},
    "wayne-state-university-law": {"city": "", "tuition": 42632, "fees": 3733, "tuition_res": 38861, "fees_res": 3733, "living": 25132},
    "west-virginia-university-college-of-law": {"city": "WV", "tuition": 45666, "fees": 1404, "tuition_res": 26370, "fees_res": 1404, "living": 26466},
    "western-new-england-university-school-of-law": {"city": "", "tuition": 50320, "fees": 2700, "tuition_res": 50320, "fees_res": 2700, "living": 20731},
    "western-state-law": {"city": "", "tuition": 54840, "fees": 808, "tuition_res": 54840, "fees_res": 808, "living": 32140},
    "widener-commonwealth-law": {"city": "", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 32862, "credit_res": 1875, "credit_oos": 1875},
    "widener-delaware-law": {"city": "", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 25720, "credit_res": 1950, "credit_oos": 1950},
    "willamette-university-college-of-law": {"city": "", "tuition": 56600, "fees": 80, "tuition_res": 56600, "fees_res": 80, "living": 22828},
    "william-and-mary-law-school": {"city": "VA", "tuition": 62900, "fees": 7939, "tuition_res": 38734, "fees_res": 7368, "living": 25184},
    "wilmington-university-law": {"city": "", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 24622, "credit_res": 838, "credit_oos": 838},
    "wisconsin-law-school": {"city": "WI", "tuition": 53640, "fees": 1660, "tuition_res": 36526, "fees_res": 1660, "living": 25056},
    "university-of-wyoming-college-of-law": {"city": "WY", "tuition": None, "fees": None, "tuition_res": None, "fees_res": None, "living": 17600, "credit_res": 576, "credit_oos": 1229},
    "yale-law-school": {"city": "CT", "tuition": 76636, "fees": 2325, "tuition_res": 76636, "fees_res": 2325, "living": 28202},
}

# --- EMPLOYMENT OUTCOMES (2024 Graduates, NALP/School data) ---
outcomes = {
    "stanford-law-school":                  {"grads":190, "biglaw_n":80,  "biglaw_pct":0.42, "mid_n":7,  "small_n":6,  "fed_clerk_n":35, "fed_clerk_pct":0.18, "state_clerk_n":3},
    "yale-law-school":                       {"grads":205, "biglaw_n":66,  "biglaw_pct":0.32, "mid_n":2,  "small_n":53, "fed_clerk_n":56, "fed_clerk_pct":0.27, "state_clerk_n":4},
    "uva-law-school":                        {"grads":295, "biglaw_n":180, "biglaw_pct":0.61, "mid_n":11, "small_n":4,  "fed_clerk_n":45, "fed_clerk_pct":0.15, "state_clerk_n":4},
    "penn-law":                              {"grads":303, "biglaw_n":195, "biglaw_pct":0.64, "mid_n":4,  "small_n":3,  "fed_clerk_n":25, "fed_clerk_pct":0.08, "state_clerk_n":5},
    "harvard-law-school":                    {"grads":598, "biglaw_n":320, "biglaw_pct":0.54, "mid_n":13, "small_n":6,  "fed_clerk_n":109,"fed_clerk_pct":0.18, "state_clerk_n":19},
    "duke-law":                              {"grads":277, "biglaw_n":188, "biglaw_pct":0.68, "mid_n":19, "small_n":7,  "fed_clerk_n":29, "fed_clerk_pct":0.10, "state_clerk_n":8},
    "nyu-law-school":                        {"grads":511, "biglaw_n":283, "biglaw_pct":0.55, "mid_n":27, "small_n":5,  "fed_clerk_n":26, "fed_clerk_pct":0.05, "state_clerk_n":10},
    "michigan-law-school":                   {"grads":311, "biglaw_n":162, "biglaw_pct":0.52, "mid_n":19, "small_n":10, "fed_clerk_n":33, "fed_clerk_pct":0.11, "state_clerk_n":21},
    "columbia-law-school":                   {"grads":512, "biglaw_n":341, "biglaw_pct":0.67, "mid_n":37, "small_n":4,  "fed_clerk_n":29, "fed_clerk_pct":0.06, "state_clerk_n":4},
    "northwestern-law-school":               {"grads":264, "biglaw_n":173, "biglaw_pct":0.66, "mid_n":17, "small_n":6,  "fed_clerk_n":14, "fed_clerk_pct":0.05, "state_clerk_n":6},
    "ucla-law":                              {"grads":362, "biglaw_n":188, "biglaw_pct":0.52, "mid_n":12, "small_n":15, "fed_clerk_n":18, "fed_clerk_pct":0.05, "state_clerk_n":4},
    "uc-berkeley-law-school":                {"grads":399, "biglaw_n":213, "biglaw_pct":0.53, "mid_n":20, "small_n":8,  "fed_clerk_n":36, "fed_clerk_pct":0.09, "state_clerk_n":14},
    "washington-university-school-of-law":   {"grads":281, "biglaw_n":112, "biglaw_pct":0.40, "mid_n":19, "small_n":20, "fed_clerk_n":27, "fed_clerk_pct":0.10, "state_clerk_n":11},
    "georgetown-law":                        {"grads":634, "biglaw_n":361, "biglaw_pct":0.57, "mid_n":23, "small_n":21, "fed_clerk_n":32, "fed_clerk_pct":0.05, "state_clerk_n":26},
    "ut-austin-law-school":                  {"grads":399, "biglaw_n":175, "biglaw_pct":0.44, "mid_n":20, "small_n":10, "fed_clerk_n":50, "fed_clerk_pct":0.13, "state_clerk_n":15},
    "vanderbilt-law-school":                 {"grads":163, "biglaw_n":79,  "biglaw_pct":0.48, "mid_n":10, "small_n":9,  "fed_clerk_n":15, "fed_clerk_pct":0.09, "state_clerk_n":4},
    "cornell-law-school":                    {"grads":192, "biglaw_n":141, "biglaw_pct":0.73, "mid_n":12, "small_n":3,  "fed_clerk_n":13, "fed_clerk_pct":0.07, "state_clerk_n":0},
    "unc-law-school":                        {"grads":204, "biglaw_n":52,  "biglaw_pct":0.25, "mid_n":10, "small_n":11, "fed_clerk_n":12, "fed_clerk_pct":0.06, "state_clerk_n":9},
    "notre-dame-law-school":                 {"grads":180, "biglaw_n":74,  "biglaw_pct":0.41, "mid_n":7,  "small_n":8,  "fed_clerk_n":31, "fed_clerk_pct":0.17, "state_clerk_n":9},
    "university-of-minnesota-law-school":    {"grads":224, "biglaw_n":37,  "biglaw_pct":0.17, "mid_n":18, "small_n":5,  "fed_clerk_n":12, "fed_clerk_pct":0.05, "state_clerk_n":43},
    "boston-university-law-school":          {"grads":286, "biglaw_n":106, "biglaw_pct":0.37, "mid_n":25, "small_n":12, "fed_clerk_n":11, "fed_clerk_pct":0.04, "state_clerk_n":16},
    "uga-law-school":                        {"grads":176, "biglaw_n":33,  "biglaw_pct":0.19, "mid_n":8,  "small_n":3,  "fed_clerk_n":17, "fed_clerk_pct":0.10, "state_clerk_n":10},
    "texas-am-law-school":                   {"grads":178, "biglaw_n":34,  "biglaw_pct":0.19, "mid_n":2,  "small_n":9,  "fed_clerk_n":3,  "fed_clerk_pct":0.02, "state_clerk_n":0},
    "boston-college-law-school":             {"grads":329, "biglaw_n":149, "biglaw_pct":0.45, "mid_n":23, "small_n":23, "fed_clerk_n":14, "fed_clerk_pct":0.04, "state_clerk_n":17},
    "usc-gould-school-of-law":               {"grads":209, "biglaw_n":122, "biglaw_pct":0.58, "mid_n":9,  "small_n":13, "fed_clerk_n":3,  "fed_clerk_pct":0.01, "state_clerk_n":3},
    "wake-forest-law-school":                {"grads":151, "biglaw_n":38,  "biglaw_pct":0.25, "mid_n":9,  "small_n":13, "fed_clerk_n":3,  "fed_clerk_pct":0.02, "state_clerk_n":8},
    "moritz-college-of-law":                {"grads":168, "biglaw_n":23,  "biglaw_pct":0.14, "mid_n":22, "small_n":11, "fed_clerk_n":8,  "fed_clerk_pct":0.05, "state_clerk_n":3},
    "wisconsin-law-school":                  {"grads":255, "biglaw_n":33,  "biglaw_pct":0.13, "mid_n":8,  "small_n":15, "fed_clerk_n":8,  "fed_clerk_pct":0.03, "state_clerk_n":7},
    "antonin-scalia-law-school":             {"grads":224, "biglaw_n":26,  "biglaw_pct":0.12, "mid_n":9,  "small_n":5,  "fed_clerk_n":15, "fed_clerk_pct":0.07, "state_clerk_n":33},
    "gw-law-school":                         {"grads":534, "biglaw_n":165, "biglaw_pct":0.31, "mid_n":28, "small_n":18, "fed_clerk_n":16, "fed_clerk_pct":0.03, "state_clerk_n":42},
    "william-and-mary-law-school":           {"grads":167, "biglaw_n":31,  "biglaw_pct":0.19, "mid_n":4,  "small_n":11, "fed_clerk_n":6,  "fed_clerk_pct":0.04, "state_clerk_n":18},
    "washington-and-lee-law-school":         {"grads":134, "biglaw_n":30,  "biglaw_pct":0.22, "mid_n":12, "small_n":12, "fed_clerk_n":7,  "fed_clerk_pct":0.05, "state_clerk_n":12},
    "iowa-law":                              {"grads":165, "biglaw_n":28,  "biglaw_pct":0.17, "mid_n":8,  "small_n":9,  "fed_clerk_n":9,  "fed_clerk_pct":0.05, "state_clerk_n":21},
    "uc-irvine-law":                         {"grads":152, "biglaw_n":48,  "biglaw_pct":0.32, "mid_n":11, "small_n":9,  "fed_clerk_n":3,  "fed_clerk_pct":0.02, "state_clerk_n":2},
    "emory-law-school":                      {"grads":262, "biglaw_n":96,  "biglaw_pct":0.37, "mid_n":18, "small_n":15, "fed_clerk_n":11, "fed_clerk_pct":0.04, "state_clerk_n":7},
    "asu-law":                               {"grads":272, "biglaw_n":32,  "biglaw_pct":0.12, "mid_n":15, "small_n":13, "fed_clerk_n":9,  "fed_clerk_pct":0.03, "state_clerk_n":36},
    "colorado-law":                          {"grads":154, "biglaw_n":31,  "biglaw_pct":0.20, "mid_n":10, "small_n":6,  "fed_clerk_n":3,  "fed_clerk_pct":0.02, "state_clerk_n":28},
    "uc-davis-law-school":                   {"grads":222, "biglaw_n":64,  "biglaw_pct":0.29, "mid_n":7,  "small_n":7,  "fed_clerk_n":3,  "fed_clerk_pct":0.01, "state_clerk_n":10},
    "uconn-law-school":                      {"grads":134, "biglaw_n":17,  "biglaw_pct":0.13, "mid_n":7,  "small_n":2,  "fed_clerk_n":0,  "fed_clerk_pct":0.00, "state_clerk_n":18},
    "usd-law-school":                        {"grads":244, "biglaw_n":33,  "biglaw_pct":0.14, "mid_n":11, "small_n":11, "fed_clerk_n":3,  "fed_clerk_pct":0.01, "state_clerk_n":0},
    "university-of-arizona-law-school":      {"grads":107, "biglaw_n":9,   "biglaw_pct":0.08, "mid_n":3,  "small_n":6,  "fed_clerk_n":6,  "fed_clerk_pct":0.06, "state_clerk_n":19},
    "cardozo-school-of-law":                {"grads":291, "biglaw_n":60,  "biglaw_pct":0.21, "mid_n":10, "small_n":13, "fed_clerk_n":8,  "fed_clerk_pct":0.03, "state_clerk_n":7},
    "northeastern-law-school":               {"grads":196, "biglaw_n":42,  "biglaw_pct":0.21, "mid_n":8,  "small_n":15, "fed_clerk_n":3,  "fed_clerk_pct":0.02, "state_clerk_n":10},
    "loyola-law-school-los-angeles":         {"grads":316, "biglaw_n":56,  "biglaw_pct":0.18, "mid_n":13, "small_n":19, "fed_clerk_n":5,  "fed_clerk_pct":0.02, "state_clerk_n":3},
    "uc-law-san-francisco":                  {"grads":316, "biglaw_n":90,  "biglaw_pct":0.28, "mid_n":11, "small_n":24, "fed_clerk_n":5,  "fed_clerk_pct":0.02, "state_clerk_n":4},
    "denver-law-school":                     {"grads":195, "biglaw_n":20,  "biglaw_pct":0.10, "mid_n":7,  "small_n":4,  "fed_clerk_n":0,  "fed_clerk_pct":0.00, "state_clerk_n":25},
    "chapman-university-fowler-school-of-law":{"grads":122,"biglaw_n":9,   "biglaw_pct":0.07, "mid_n":3,  "small_n":4,  "fed_clerk_n":2,  "fed_clerk_pct":0.02, "state_clerk_n":1},
    "american-university-law-school":        {"grads":315, "biglaw_n":51,  "biglaw_pct":0.16, "mid_n":7,  "small_n":6,  "fed_clerk_n":7,  "fed_clerk_pct":0.02, "state_clerk_n":46},
}



# --- ADMISSIONS DATA (from ABA 509 reports) — run parse_509_local.py to populate ---
admissions_data = {
    "university-of-akron-school-of-law": {
        "apps": 762, "offers": 446, "accept_rate": 58.53,
        "enrollees": 201, "yield_rate": 44.62, "enroll_rate": 26.12,
        "lsat25": 151, "lsat50": 154, "lsat75": 157,
        "gpa25": 3.22, "gpa50": 3.51, "gpa75": 3.76,
    },
    "university-of-alabama-school-of-law": {
        "apps": 1748, "offers": 448, "accept_rate": 25.63,
        "enrollees": 137, "yield_rate": 30.58, "enroll_rate": 7.84,
        "lsat25": 161, "lsat50": 167, "lsat75": 168,
        "gpa25": 3.76, "gpa50": 3.97, "gpa75": 4.05,
    },
    "albany-law-school": {
        "apps": 2638, "offers": 988, "accept_rate": 37.45,
        "enrollees": 236, "yield_rate": 22.77, "enroll_rate": 8.53,
        "lsat25": 154, "lsat50": 158, "lsat75": 161,
        "gpa25": 3.16, "gpa50": 3.54, "gpa75": 3.75,
    },
    "american-university-law-school": {
        "apps": 5952, "offers": 1981, "accept_rate": 33.28,
        "enrollees": 400, "yield_rate": 20.19, "enroll_rate": 6.72,
        "lsat25": 157, "lsat50": 162, "lsat75": 163,
        "gpa25": 3.32, "gpa50": 3.63, "gpa75": 3.75,
    },
    "appalachian-school-of-law": {
        "apps": 732, "offers": 446, "accept_rate": 60.93,
        "enrollees": 88, "yield_rate": 18.83, "enroll_rate": 11.48,
        "lsat25": 145, "lsat50": 146, "lsat75": 151,
        "gpa25": 2.72, "gpa50": 3.2, "gpa75": 3.53,
    },
    "asu-law": {
        "apps": 4610, "offers": 912, "accept_rate": 19.78,
        "enrollees": 264, "yield_rate": 28.95, "enroll_rate": 5.73,
        "lsat25": 158, "lsat50": 165, "lsat75": 167,
        "gpa25": 3.58, "gpa50": 3.91, "gpa75": 3.98,
    },
    "university-of-arizona-law-school": {
        "apps": 1781, "offers": 497, "accept_rate": 27.91,
        "enrollees": 109, "yield_rate": 21.93, "enroll_rate": 6.12,
        "lsat25": 162, "lsat50": 163, "lsat75": 165,
        "gpa25": 3.54, "gpa50": 3.78, "gpa75": 3.9,
    },
    "arkansas-little-rock-law": {
        "apps": 603, "offers": 331, "accept_rate": 54.89,
        "enrollees": 152, "yield_rate": 45.32, "enroll_rate": 24.88,
        "lsat25": 150, "lsat50": 153, "lsat75": 158,
        "gpa25": 2.98, "gpa50": 3.41, "gpa75": 3.72,
    },
    "university-of-arkansas-school-of-law": {
        "apps": 950, "offers": 253, "accept_rate": 26.63,
        "enrollees": 135, "yield_rate": 53.36, "enroll_rate": 14.21,
        "lsat25": 155, "lsat50": 158, "lsat75": 161,
        "gpa25": 3.54, "gpa50": 3.71, "gpa75": 3.87,
    },
    "john-marshall-atlanta": {
        "apps": 1274, "offers": 416, "accept_rate": 32.65,
        "enrollees": 135, "yield_rate": 31.97, "enroll_rate": 10.44,
        "lsat25": 150, "lsat50": 152, "lsat75": 155,
        "gpa25": 2.78, "gpa50": 3.13, "gpa75": 3.49,
    },
    "ave-maria-school-of-law": {
        "apps": 1212, "offers": 569, "accept_rate": 46.95,
        "enrollees": 164, "yield_rate": 26.71, "enroll_rate": 12.54,
        "lsat25": 152, "lsat50": 155, "lsat75": 158,
        "gpa25": 3.07, "gpa50": 3.36, "gpa75": 3.66,
    },
    "university-of-baltimore-school-of-law": {
        "apps": 1299, "offers": 643, "accept_rate": 49.5,
        "enrollees": 246, "yield_rate": 37.33, "enroll_rate": 18.48,
        "lsat25": 152, "lsat50": 155, "lsat75": 158,
        "gpa25": 3.12, "gpa50": 3.45, "gpa75": 3.7,
    },
    "barry-university-law": {
        "apps": 1582, "offers": 885, "accept_rate": 55.94,
        "enrollees": 257, "yield_rate": 29.04, "enroll_rate": 16.25,
        "lsat25": 148, "lsat50": 150, "lsat75": 152,
        "gpa25": 2.93, "gpa50": 3.23, "gpa75": 3.48,
    },
    "baylor-law-school": {
        "apps": 3658, "offers": 753, "accept_rate": 20.59,
        "enrollees": 165, "yield_rate": 21.91, "enroll_rate": 4.51,
        "lsat25": 159, "lsat50": 164, "lsat75": 165,
        "gpa25": 3.49, "gpa50": 3.77, "gpa75": 3.89,
    },
    "belmont-university-law": {
        "apps": 1337, "offers": 503, "accept_rate": 37.62,
        "enrollees": 136, "yield_rate": 27.04, "enroll_rate": 10.17,
        "lsat25": 158, "lsat50": 161, "lsat75": 163,
        "gpa25": 3.66, "gpa50": 3.84, "gpa75": 3.94,
    },
    "boston-college-law-school": {
        "apps": 7668, "offers": 650, "accept_rate": 8.48,
        "enrollees": 216, "yield_rate": 33.23, "enroll_rate": 2.82,
        "lsat25": 162, "lsat50": 168, "lsat75": 169,
        "gpa25": 3.67, "gpa50": 3.83, "gpa75": 3.9,
    },
    "boston-university-law-school": {
        "apps": 7892, "offers": 953, "accept_rate": 12.08,
        "enrollees": 234, "yield_rate": 24.55, "enroll_rate": 2.97,
        "lsat25": 164, "lsat50": 170, "lsat75": 171,
        "gpa25": 3.71, "gpa50": 3.88, "gpa75": 3.93,
    },
    "byu-law-school": {
        "apps": 649, "offers": 148, "accept_rate": 22.8,
        "enrollees": 120, "yield_rate": 81.08, "enroll_rate": 18.49,
        "lsat25": 168, "lsat50": 170, "lsat75": 172,
        "gpa25": 3.82, "gpa50": 3.95, "gpa75": 3.98,
    },
    "brooklyn-law-school": {
        "apps": 4600, "offers": 2024, "accept_rate": 44.0,
        "enrollees": 414, "yield_rate": 20.45, "enroll_rate": 9.0,
        "lsat25": 158, "lsat50": 161, "lsat75": 163,
        "gpa25": 3.37, "gpa50": 3.59, "gpa75": 3.75,
    },
    "university-of-buffalo-law": {
        "apps": 977, "offers": 414, "accept_rate": 42.37,
        "enrollees": 163, "yield_rate": 39.37, "enroll_rate": 16.68,
        "lsat25": 155, "lsat50": 157, "lsat75": 161,
        "gpa25": 3.51, "gpa50": 3.71, "gpa75": 3.88,
    },
    "california-western-school-of-law": {
        "apps": 2144, "offers": 1006, "accept_rate": 46.92,
        "enrollees": 242, "yield_rate": 23.36, "enroll_rate": 10.96,
        "lsat25": 152, "lsat50": 155, "lsat75": 158,
        "gpa25": 3.16, "gpa50": 3.5, "gpa75": 3.66,
    },
    "uc-berkeley-law-school": {
        "apps": 6562, "offers": 974, "accept_rate": 14.84,
        "enrollees": 348, "yield_rate": 35.73, "enroll_rate": 5.3,
        "lsat25": 167, "lsat50": 170, "lsat75": 172,
        "gpa25": 3.84, "gpa50": 3.92, "gpa75": 3.99,
    },
    "uc-davis-law-school": {
        "apps": 3912, "offers": 824, "accept_rate": 21.06,
        "enrollees": 205, "yield_rate": 24.88, "enroll_rate": 5.24,
        "lsat25": 160, "lsat50": 165, "lsat75": 167,
        "gpa25": 3.49, "gpa50": 3.7, "gpa75": 3.87,
    },
    "uc-irvine-law": {
        "apps": 3720, "offers": 555, "accept_rate": 14.92,
        "enrollees": 188, "yield_rate": 33.87, "enroll_rate": 5.05,
        "lsat25": 166, "lsat50": 169, "lsat75": 170,
        "gpa25": 3.59, "gpa50": 3.8, "gpa75": 3.9,
    },
    "ucla-law": {
        "apps": 8623, "offers": 1039, "accept_rate": 12.05,
        "enrollees": 322, "yield_rate": 30.99, "enroll_rate": 3.73,
        "lsat25": 166, "lsat50": 171, "lsat75": 172,
        "gpa25": 3.73, "gpa50": 3.95, "gpa75": 4.0,
    },
    "uc-law-san-francisco": {
        "apps": 4593, "offers": 1262, "accept_rate": 27.48,
        "enrollees": 374, "yield_rate": 29.64, "enroll_rate": 8.14,
        "lsat25": 158, "lsat50": 161, "lsat75": 163,
        "gpa25": 3.55, "gpa50": 3.71, "gpa75": 3.83,
    },
    "campbell-university-law": {
        "apps": 1168, "offers": 592, "accept_rate": 50.68,
        "enrollees": 199, "yield_rate": 33.28, "enroll_rate": 16.87,
        "lsat25": 152, "lsat50": 156, "lsat75": 159,
        "gpa25": 3.21, "gpa50": 3.51, "gpa75": 3.77,
    },
    "capital-university-law": {
        "apps": 874, "offers": 576, "accept_rate": 65.9,
        "enrollees": 178, "yield_rate": 30.21, "enroll_rate": 19.91,
        "lsat25": 148, "lsat50": 151, "lsat75": 157,
        "gpa25": 3.07, "gpa50": 3.44, "gpa75": 3.66,
    },
    "cardozo-school-of-law": {
        "apps": 3833, "offers": 1133, "accept_rate": 29.56,
        "enrollees": 319, "yield_rate": 28.16, "enroll_rate": 8.32,
        "lsat25": 162, "lsat50": 165, "lsat75": 167,
        "gpa25": 3.56, "gpa50": 3.78, "gpa75": 3.87,
    },
    "case-western-reserve-university-school-of-law": {
        "apps": 2267, "offers": 789, "accept_rate": 34.8,
        "enrollees": 178, "yield_rate": 22.56, "enroll_rate": 7.85,
        "lsat25": 156, "lsat50": 162, "lsat75": 164,
        "gpa25": 3.51, "gpa50": 3.76, "gpa75": 3.89,
    },
    "catholic-university-law": {
        "apps": 1867, "offers": 605, "accept_rate": 32.4,
        "enrollees": 179, "yield_rate": 29.59, "enroll_rate": 9.59,
        "lsat25": 158, "lsat50": 161, "lsat75": 163,
        "gpa25": 3.46, "gpa50": 3.68, "gpa75": 3.85,
    },
    "chapman-university-fowler-school-of-law": {
        "apps": 2514, "offers": 516, "accept_rate": 20.53,
        "enrollees": 152, "yield_rate": 29.46, "enroll_rate": 6.05,
        "lsat25": 156, "lsat50": 163, "lsat75": 164,
        "gpa25": 3.32, "gpa50": 3.7, "gpa75": 3.82,
    },
    "charleston-school-of-law": {
        "apps": 2215, "offers": 1112, "accept_rate": 50.2,
        "enrollees": 239, "yield_rate": 21.31, "enroll_rate": 10.7,
        "lsat25": 152, "lsat50": 154, "lsat75": 157,
        "gpa25": 3.26, "gpa50": 3.51, "gpa75": 3.75,
    },
    "chicago-law-school": {
        "apps": 6581, "offers": 641, "accept_rate": 9.74,
        "enrollees": 191, "yield_rate": 29.8, "enroll_rate": 2.9,
        "lsat25": 171, "lsat50": 174, "lsat75": 176,
        "gpa25": 3.87, "gpa50": 3.97, "gpa75": 4.0,
    },
    "chicago-kent-college-of-law": {
        "apps": 2212, "offers": 854, "accept_rate": 38.61,
        "enrollees": 244, "yield_rate": 28.57, "enroll_rate": 11.03,
        "lsat25": 156, "lsat50": 161, "lsat75": 164,
        "gpa25": 3.45, "gpa50": 3.72, "gpa75": 3.89,
    },
    "university-of-cincinnati-college-of-law": {
        "apps": 1353, "offers": 443, "accept_rate": 32.74,
        "enrollees": 135, "yield_rate": 30.47, "enroll_rate": 9.98,
        "lsat25": 156, "lsat50": 159, "lsat75": 161,
        "gpa25": 3.64, "gpa50": 3.8, "gpa75": 3.91,
    },
    "cuny-law-school": {
        "apps": 2300, "offers": 659, "accept_rate": 28.65,
        "enrollees": 212, "yield_rate": 31.87, "enroll_rate": 9.13,
        "lsat25": 152, "lsat50": 155, "lsat75": 161,
        "gpa25": 3.26, "gpa50": 3.54, "gpa75": 3.77,
    },
    "cleveland-state-law": {
        "apps": 1733, "offers": 597, "accept_rate": 34.45,
        "enrollees": 195, "yield_rate": 32.66, "enroll_rate": 11.25,
        "lsat25": 152, "lsat50": 155, "lsat75": 158,
        "gpa25": 3.11, "gpa50": 3.5, "gpa75": 3.76,
    },
    "colorado-law": {
        "apps": 3458, "offers": 953, "accept_rate": 27.56,
        "enrollees": 193, "yield_rate": 20.15, "enroll_rate": 5.55,
        "lsat25": 161, "lsat50": 164, "lsat75": 167,
        "gpa25": 3.62, "gpa50": 3.81, "gpa75": 3.91,
    },
    "columbia-law-school": {
        "apps": 9463, "offers": 1120, "accept_rate": 11.84,
        "enrollees": 412, "yield_rate": 36.7, "enroll_rate": 4.34,
        "lsat25": 169, "lsat50": 173, "lsat75": 175,
        "gpa25": 3.85, "gpa50": 3.92, "gpa75": 3.98,
    },
    "uconn-law-school": {
        "apps": 2299, "offers": 461, "accept_rate": 20.05,
        "enrollees": 149, "yield_rate": 32.32, "enroll_rate": 6.48,
        "lsat25": 160, "lsat50": 162, "lsat75": 164,
        "gpa25": 3.55, "gpa50": 3.78, "gpa75": 3.9,
    },
    "cooley-law-school": {
        "apps": 1323, "offers": 700, "accept_rate": 52.91,
        "enrollees": 237, "yield_rate": 33.57, "enroll_rate": 17.76,
        "lsat25": 145, "lsat50": 147, "lsat75": 151,
        "gpa25": 2.67, "gpa50": 3.09, "gpa75": 3.45,
    },
    "cornell-law-school": {
        "apps": 4608, "offers": 838, "accept_rate": 18.19,
        "enrollees": 210, "yield_rate": 25.06, "enroll_rate": 4.56,
        "lsat25": 168, "lsat50": 173, "lsat75": 175,
        "gpa25": 3.75, "gpa50": 3.92, "gpa75": 3.97,
    },
    "creighton-university-law": {
        "apps": 815, "offers": 526, "accept_rate": 64.54,
        "enrollees": 160, "yield_rate": 30.42, "enroll_rate": 19.63,
        "lsat25": 150, "lsat50": 153, "lsat75": 156,
        "gpa25": 3.14, "gpa50": 3.51, "gpa75": 3.74,
    },
    "university-of-dayton-school-of-law": {
        "apps": 1611, "offers": 494, "accept_rate": 30.66,
        "enrollees": 167, "yield_rate": 33.81, "enroll_rate": 10.37,
        "lsat25": 154, "lsat50": 157, "lsat75": 160,
        "gpa25": 3.31, "gpa50": 3.6, "gpa75": 3.87,
    },
    "denver-law-school": {
        "apps": 2214, "offers": 915, "accept_rate": 41.33,
        "enrollees": 283, "yield_rate": 30.93, "enroll_rate": 12.78,
        "lsat25": 156, "lsat50": 160, "lsat75": 162,
        "gpa25": 3.36, "gpa50": 3.66, "gpa75": 3.8,
    },
    "depaul-university-law": {
        "apps": 2643, "offers": 920, "accept_rate": 34.81,
        "enrollees": 185, "yield_rate": 20.11, "enroll_rate": 7.0,
        "lsat25": 154, "lsat50": 158, "lsat75": 160,
        "gpa25": 3.17, "gpa50": 3.61, "gpa75": 3.8,
    },
    "university-of-detroit-mercy-school-of-law": {
        "apps": 1478, "offers": 612, "accept_rate": 41.41,
        "enrollees": 247, "yield_rate": 40.2, "enroll_rate": 16.64,
        "lsat25": 154, "lsat50": 156, "lsat75": 159,
        "gpa25": 3.16, "gpa50": 3.52, "gpa75": 3.73,
    },
    "udc-law-school": {
        "apps": 749, "offers": 277, "accept_rate": 36.98,
        "enrollees": 82, "yield_rate": 25.99, "enroll_rate": 9.61,
        "lsat25": 148, "lsat50": 151, "lsat75": 153,
        "gpa25": 3.05, "gpa50": 3.28, "gpa75": 3.55,
    },
    "drake-university-law": {
        "apps": 788, "offers": 432, "accept_rate": 54.82,
        "enrollees": 141, "yield_rate": 32.64, "enroll_rate": 17.89,
        "lsat25": 153, "lsat50": 156, "lsat75": 159,
        "gpa25": 3.38, "gpa50": 3.67, "gpa75": 3.89,
    },
    "drexel-university-thomas-r-kline-school-of-law": {
        "apps": 2689, "offers": 797, "accept_rate": 29.64,
        "enrollees": 177, "yield_rate": 22.21, "enroll_rate": 6.58,
        "lsat25": 154, "lsat50": 160, "lsat75": 162,
        "gpa25": 3.35, "gpa50": 3.76, "gpa75": 3.87,
    },
    "duke-law": {
        "apps": 6240, "offers": 804, "accept_rate": 12.88,
        "enrollees": 222, "yield_rate": 27.61, "enroll_rate": 3.56,
        "lsat25": 169, "lsat50": 171, "lsat75": 172,
        "gpa25": 3.83, "gpa50": 3.91, "gpa75": 3.96,
    },
    "duquesne-university-law": {
        "apps": 1432, "offers": 619, "accept_rate": 43.23,
        "enrollees": 183, "yield_rate": 29.56, "enroll_rate": 12.78,
        "lsat25": 155, "lsat50": 157, "lsat75": 160,
        "gpa25": 3.35, "gpa50": 3.62, "gpa75": 3.85,
    },
    "elon-university-law": {
        "apps": 1107, "offers": 589, "accept_rate": 53.21,
        "enrollees": 228, "yield_rate": 38.54, "enroll_rate": 20.51,
        "lsat25": 151, "lsat50": 154, "lsat75": 158,
        "gpa25": 3.11, "gpa50": 3.4, "gpa75": 3.7,
    },
    "emory-law-school": {
        "apps": 3891, "offers": 1173, "accept_rate": 30.15,
        "enrollees": 391, "yield_rate": 33.16, "enroll_rate": 10.0,
        "lsat25": 162, "lsat50": 166, "lsat75": 168,
        "gpa25": 3.68, "gpa50": 3.82, "gpa75": 3.87,
    },
    "faulkner-university-law": {
        "apps": 620, "offers": 334, "accept_rate": 53.87,
        "enrollees": 167, "yield_rate": 47.6, "enroll_rate": 25.65,
        "lsat25": 148, "lsat50": 150, "lsat75": 154,
        "gpa25": 2.99, "gpa50": 3.33, "gpa75": 3.66,
    },
    "florida-am-law": {
        "apps": 1546, "offers": 536, "accept_rate": 34.67,
        "enrollees": 177, "yield_rate": 33.02, "enroll_rate": 11.45,
        "lsat25": 149, "lsat50": 152, "lsat75": 155,
        "gpa25": 3.18, "gpa50": 3.48, "gpa75": 3.74,
    },
    "florida-international-university-college-of-law": {
        "apps": 3100, "offers": 582, "accept_rate": 18.77,
        "enrollees": 175, "yield_rate": 30.07, "enroll_rate": 5.65,
        "lsat25": 159, "lsat50": 161, "lsat75": 162,
        "gpa25": 3.66, "gpa50": 3.8, "gpa75": 3.89,
    },
    "florida-state-university-college-of-law": {
        "apps": 2711, "offers": 433, "accept_rate": 15.97,
        "enrollees": 153, "yield_rate": 35.33, "enroll_rate": 5.64,
        "lsat25": 158, "lsat50": 166, "lsat75": 167,
        "gpa25": 3.6, "gpa50": 3.93, "gpa75": 3.99,
    },
    "university-of-florida-levin-college-of-law": {
        "apps": 3887, "offers": 642, "accept_rate": 16.52,
        "enrollees": 228, "yield_rate": 35.51, "enroll_rate": 5.87,
        "lsat25": 165, "lsat50": 169, "lsat75": 170,
        "gpa25": 3.7, "gpa50": 3.91, "gpa75": 3.97,
    },
    "fordham-law-school": {
        "apps": 8811, "offers": 1429, "accept_rate": 16.22,
        "enrollees": 433, "yield_rate": 30.3, "enroll_rate": 4.91,
        "lsat25": 165, "lsat50": 168, "lsat75": 170,
        "gpa25": 3.6, "gpa50": 3.79, "gpa75": 3.88,
    },
    "antonin-scalia-law-school": {
        "apps": 2792, "offers": 443, "accept_rate": 15.87,
        "enrollees": 159, "yield_rate": 35.89, "enroll_rate": 5.69,
        "lsat25": 162, "lsat50": 169, "lsat75": 171,
        "gpa25": 3.55, "gpa50": 3.93, "gpa75": 3.98,
    },
    "gw-law-school": {
        "apps": 9718, "offers": 2644, "accept_rate": 27.21,
        "enrollees": 595, "yield_rate": 22.5, "enroll_rate": 6.12,
        "lsat25": 162, "lsat50": 168, "lsat75": 170,
        "gpa25": 3.55, "gpa50": 3.86, "gpa75": 3.93,
    },
    "georgetown-law": {
        "apps": 13924, "offers": 2193, "accept_rate": 15.75,
        "enrollees": 629, "yield_rate": 28.68, "enroll_rate": 4.52,
        "lsat25": 166, "lsat50": 171, "lsat75": 173,
        "gpa25": 3.75, "gpa50": 3.93, "gpa75": 3.98,
    },
    "georgia-state-university-college-of-law": {
        "apps": 2450, "offers": 749, "accept_rate": 30.57,
        "enrollees": 237, "yield_rate": 31.64, "enroll_rate": 9.67,
        "lsat25": 158, "lsat50": 160, "lsat75": 163,
        "gpa25": 3.45, "gpa50": 3.64, "gpa75": 3.82,
    },
    "uga-law-school": {
        "apps": 4695, "offers": 597, "accept_rate": 12.72,
        "enrollees": 186, "yield_rate": 31.16, "enroll_rate": 3.96,
        "lsat25": 159, "lsat50": 169, "lsat75": 171,
        "gpa25": 3.7, "gpa50": 3.92, "gpa75": 3.97,
    },
    "gonzaga-university-law": {
        "apps": 1273, "offers": 651, "accept_rate": 51.14,
        "enrollees": 201, "yield_rate": 30.88, "enroll_rate": 15.79,
        "lsat25": 153, "lsat50": 155, "lsat75": 158,
        "gpa25": 3.24, "gpa50": 3.55, "gpa75": 3.76,
    },
    "harvard-law-school": {
        "apps": 8872, "offers": 816, "accept_rate": 9.2,
        "enrollees": 483, "yield_rate": 59.19, "enroll_rate": 5.44,
        "lsat25": 171, "lsat50": 174, "lsat75": 176,
        "gpa25": 3.89, "gpa50": 3.96, "gpa75": 4.0,
    },
    "university-of-hawaii-william-s-richardson-school-of-law": {
        "apps": 1654, "offers": 434, "accept_rate": 26.24,
        "enrollees": 138, "yield_rate": 31.57, "enroll_rate": 8.28,
        "lsat25": 155, "lsat50": 158, "lsat75": 162,
        "gpa25": 3.4, "gpa50": 3.62, "gpa75": 3.8,
    },
    "hofstra-university-school-of-law": {
        "apps": 2448, "offers": 993, "accept_rate": 40.56,
        "enrollees": 260, "yield_rate": 26.08, "enroll_rate": 10.58,
        "lsat25": 154, "lsat50": 157, "lsat75": 160,
        "gpa25": 3.35, "gpa50": 3.65, "gpa75": 3.83,
    },
    "university-of-houston-law-center": {
        "apps": 3510, "offers": 807, "accept_rate": 22.99,
        "enrollees": 246, "yield_rate": 30.48, "enroll_rate": 7.01,
        "lsat25": 160, "lsat50": 163, "lsat75": 165,
        "gpa25": 3.56, "gpa50": 3.79, "gpa75": 3.9,
    },
    "howard-university-school-of-law": {
        "apps": 2428, "offers": 584, "accept_rate": 24.05,
        "enrollees": 172, "yield_rate": 29.11, "enroll_rate": 7.0,
        "lsat25": 153, "lsat50": 156, "lsat75": 160,
        "gpa25": 3.4, "gpa50": 3.63, "gpa75": 3.8,
    },
    "university-of-idaho-college-of-law": {
        "apps": 740, "offers": 494, "accept_rate": 66.76,
        "enrollees": 194, "yield_rate": 39.27, "enroll_rate": 26.22,
        "lsat25": 149, "lsat50": 153, "lsat75": 157,
        "gpa25": 3.07, "gpa50": 3.54, "gpa75": 3.81,
    },
    "university-of-illinois-chicago-school-of-law": {
        "apps": 2153, "offers": 709, "accept_rate": 32.93,
        "enrollees": 188, "yield_rate": 26.38, "enroll_rate": 8.69,
        "lsat25": 160, "lsat50": 166, "lsat75": 168,
        "gpa25": 3.55, "gpa50": 3.81, "gpa75": 3.92,
    },
    "uic-john-marshall-law-school": {
        "apps": 2488, "offers": 1308, "accept_rate": 52.57,
        "enrollees": 372, "yield_rate": 28.21, "enroll_rate": 14.83,
        "lsat25": 150, "lsat50": 153, "lsat75": 156,
        "gpa25": 3.13, "gpa50": 3.4, "gpa75": 3.64,
    },
    "indiana-university-maurer-school-of-law": {
        "apps": 1831, "offers": 655, "accept_rate": 35.77,
        "enrollees": 186, "yield_rate": 28.4, "enroll_rate": 10.16,
        "lsat25": 158, "lsat50": 164, "lsat75": 166,
        "gpa25": 3.63, "gpa50": 3.91, "gpa75": 3.98,
    },
    "indiana-university-indianapolis-law": {
        "apps": 904, "offers": 505, "accept_rate": 55.86,
        "enrollees": 273, "yield_rate": 54.06, "enroll_rate": 30.2,
        "lsat25": 152, "lsat50": 155, "lsat75": 159,
        "gpa25": 3.33, "gpa50": 3.58, "gpa75": 3.82,
    },
    "inter-american-university-pr-law": {
        "apps": 542, "offers": 276, "accept_rate": 50.92,
        "enrollees": 205, "yield_rate": 74.28, "enroll_rate": 37.82,
        "lsat25": 140, "lsat50": 144, "lsat75": 148,
        "gpa25": 3.18, "gpa50": 3.39, "gpa75": 3.65,
    },
    "iowa-law": {
        "apps": 1247, "offers": 554, "accept_rate": 44.43,
        "enrollees": 156, "yield_rate": 28.16, "enroll_rate": 12.51,
        "lsat25": 161, "lsat50": 164, "lsat75": 165,
        "gpa25": 3.62, "gpa50": 3.78, "gpa75": 3.92,
    },
    "jacksonville-university-law": {
        "apps": 572, "offers": 201, "accept_rate": 35.14,
        "enrollees": 67, "yield_rate": 33.33, "enroll_rate": 11.71,
        "lsat25": 154, "lsat50": 156, "lsat75": 159,
        "gpa25": 3.28, "gpa50": 3.48, "gpa75": 3.68,
    },
    "kansas-law": {
        "apps": 1172, "offers": 458, "accept_rate": 39.08,
        "enrollees": 115, "yield_rate": 25.11, "enroll_rate": 9.81,
        "lsat25": 158, "lsat50": 162, "lsat75": 164,
        "gpa25": 3.58, "gpa50": 3.85, "gpa75": 3.95,
    },
    "university-of-kentucky-college-of-law": {
        "apps": 1236, "offers": 504, "accept_rate": 40.78,
        "enrollees": 143, "yield_rate": 28.37, "enroll_rate": 11.57,
        "lsat25": 156, "lsat50": 159, "lsat75": 162,
        "gpa25": 3.52, "gpa50": 3.72, "gpa75": 3.9,
    },
    "lewis-and-clark-law-school": {
        "apps": 1133, "offers": 620, "accept_rate": 54.72,
        "enrollees": 171, "yield_rate": 27.58, "enroll_rate": 15.09,
        "lsat25": 158, "lsat50": 161, "lsat75": 166,
        "gpa25": 3.24, "gpa50": 3.59, "gpa75": 3.79,
    },
    "liberty-university-law": {
        "apps": 628, "offers": 369, "accept_rate": 58.76,
        "enrollees": 140, "yield_rate": 37.4, "enroll_rate": 21.97,
        "lsat25": 151, "lsat50": 154, "lsat75": 158,
        "gpa25": 3.36, "gpa50": 3.54, "gpa75": 3.81,
    },
    "lincoln-memorial-university-law": {
        "apps": 828, "offers": 432, "accept_rate": 52.17,
        "enrollees": 163, "yield_rate": 37.5, "enroll_rate": 19.57,
        "lsat25": 150, "lsat50": 153, "lsat75": 156,
        "gpa25": 3.08, "gpa50": 3.44, "gpa75": 3.68,
    },
    "louisiana-state-university-law": {
        "apps": 1253, "offers": 483, "accept_rate": 38.55,
        "enrollees": 230, "yield_rate": 47.2, "enroll_rate": 18.2,
        "lsat25": 155, "lsat50": 157, "lsat75": 160,
        "gpa25": 3.48, "gpa50": 3.74, "gpa75": 3.93,
    },
    "university-of-louisville-louis-d-brandeis-school-of-law": {
        "apps": 579, "offers": 367, "accept_rate": 63.39,
        "enrollees": 145, "yield_rate": 39.24, "enroll_rate": 24.87,
        "lsat25": 154, "lsat50": 157, "lsat75": 161,
        "gpa25": 3.32, "gpa50": 3.67, "gpa75": 3.9,
    },
    "loyola-law-school-los-angeles": {
        "apps": 4549, "offers": 1185, "accept_rate": 26.05,
        "enrollees": 337, "yield_rate": 28.1, "enroll_rate": 7.32,
        "lsat25": 160, "lsat50": 163, "lsat75": 165,
        "gpa25": 3.52, "gpa50": 3.74, "gpa75": 3.89,
    },
    "loyola-university-chicago-school-of-law": {
        "apps": 4244, "offers": 1196, "accept_rate": 28.18,
        "enrollees": 291, "yield_rate": 24.33, "enroll_rate": 6.86,
        "lsat25": 159, "lsat50": 161, "lsat75": 163,
        "gpa25": 3.55, "gpa50": 3.7, "gpa75": 3.83,
    },
    "loyola-university-new-orleans-college-of-law": {
        "apps": 1241, "offers": 632, "accept_rate": 50.93,
        "enrollees": 223, "yield_rate": 34.81, "enroll_rate": 17.73,
        "lsat25": 151, "lsat50": 154, "lsat75": 157,
        "gpa25": 3.21, "gpa50": 3.47, "gpa75": 3.71,
    },
    "university-of-maine-school-of-law": {
        "apps": 915, "offers": 341, "accept_rate": 37.27,
        "enrollees": 93, "yield_rate": 27.27, "enroll_rate": 10.16,
        "lsat25": 154, "lsat50": 158, "lsat75": 162,
        "gpa25": 3.41, "gpa50": 3.7, "gpa75": 3.8,
    },
    "marquette-university-law-school": {
        "apps": 2112, "offers": 467, "accept_rate": 22.11,
        "enrollees": 218, "yield_rate": 46.68, "enroll_rate": 10.32,
        "lsat25": 156, "lsat50": 158, "lsat75": 160,
        "gpa25": 3.5, "gpa50": 3.7, "gpa75": 3.83,
    },
    "university-of-maryland-carey-school-of-law": {
        "apps": 2876, "offers": 750, "accept_rate": 26.08,
        "enrollees": 229, "yield_rate": 30.53, "enroll_rate": 7.96,
        "lsat25": 156, "lsat50": 164, "lsat75": 166,
        "gpa25": 3.37, "gpa50": 3.72, "gpa75": 3.88,
    },
    "umass-dartmouth-law": {
        "apps": 1311, "offers": 684, "accept_rate": 52.17,
        "enrollees": 140, "yield_rate": 20.18, "enroll_rate": 10.53,
        "lsat25": 149, "lsat50": 152, "lsat75": 155,
        "gpa25": 3.1, "gpa50": 3.36, "gpa75": 3.69,
    },
    "university-of-memphis-school-of-law": {
        "apps": 582, "offers": 239, "accept_rate": 41.07,
        "enrollees": 96, "yield_rate": 40.17, "enroll_rate": 16.49,
        "lsat25": 152, "lsat50": 155, "lsat75": 159,
        "gpa25": 3.22, "gpa50": 3.53, "gpa75": 3.77,
    },
    "mercer-university-school-of-law": {
        "apps": 1184, "offers": 447, "accept_rate": 37.75,
        "enrollees": 152, "yield_rate": 34.0, "enroll_rate": 12.84,
        "lsat25": 155, "lsat50": 157, "lsat75": 160,
        "gpa25": 3.3, "gpa50": 3.62, "gpa75": 3.8,
    },
    "university-of-miami-law": {
        "apps": 4593, "offers": 1132, "accept_rate": 24.65,
        "enrollees": 306, "yield_rate": 27.03, "enroll_rate": 6.66,
        "lsat25": 160, "lsat50": 164, "lsat75": 165,
        "gpa25": 3.58, "gpa50": 3.79, "gpa75": 3.88,
    },
    "michigan-state-university-college-of-law": {
        "apps": 1526, "offers": 397, "accept_rate": 26.02,
        "enrollees": 150, "yield_rate": 37.78, "enroll_rate": 9.83,
        "lsat25": 158, "lsat50": 162, "lsat75": 164,
        "gpa25": 3.47, "gpa50": 3.73, "gpa75": 3.87,
    },
    "michigan-law-school": {
        "apps": 8982, "offers": 770, "accept_rate": 8.57,
        "enrollees": 329, "yield_rate": 42.73, "enroll_rate": 3.66,
        "lsat25": 168, "lsat50": 171, "lsat75": 173,
        "gpa25": 3.74, "gpa50": 3.88, "gpa75": 3.95,
    },
    "university-of-minnesota-law-school": {
        "apps": 3405, "offers": 909, "accept_rate": 26.7,
        "enrollees": 225, "yield_rate": 24.75, "enroll_rate": 6.61,
        "lsat25": 166, "lsat50": 171, "lsat75": 173,
        "gpa25": 3.62, "gpa50": 3.88, "gpa75": 3.95,
    },
    "mississippi-college-school-of-law": {
        "apps": 806, "offers": 482, "accept_rate": 59.8,
        "enrollees": 178, "yield_rate": 36.31, "enroll_rate": 21.71,
        "lsat25": 148, "lsat50": 152, "lsat75": 155,
        "gpa25": 3.03, "gpa50": 3.37, "gpa75": 3.73,
    },
    "university-of-missouri-school-of-law": {
        "apps": 798, "offers": 335, "accept_rate": 41.98,
        "enrollees": 125, "yield_rate": 37.31, "enroll_rate": 15.66,
        "lsat25": 156, "lsat50": 161, "lsat75": 162,
        "gpa25": 3.47, "gpa50": 3.72, "gpa75": 3.91,
    },
    "umkc-law-school": {
        "apps": 714, "offers": 328, "accept_rate": 45.94,
        "enrollees": 142, "yield_rate": 43.29, "enroll_rate": 19.89,
        "lsat25": 154, "lsat50": 156, "lsat75": 159,
        "gpa25": 3.33, "gpa50": 3.65, "gpa75": 3.8,
    },
    "william-mitchell-college-of-law": {
        "apps": 1603, "offers": 824, "accept_rate": 51.4,
        "enrollees": 363, "yield_rate": 43.69, "enroll_rate": 22.46,
        "lsat25": 151, "lsat50": 154, "lsat75": 158,
        "gpa25": 2.96, "gpa50": 3.29, "gpa75": 3.6,
    },
    "university-of-montana-school-of-law": {
        "apps": 568, "offers": 292, "accept_rate": 51.41,
        "enrollees": 95, "yield_rate": 32.53, "enroll_rate": 16.73,
        "lsat25": 153, "lsat50": 156, "lsat75": 159,
        "gpa25": 3.33, "gpa50": 3.64, "gpa75": 3.8,
    },
    "university-of-nebraska-college-of-law": {
        "apps": 1096, "offers": 440, "accept_rate": 40.15,
        "enrollees": 152, "yield_rate": 34.55, "enroll_rate": 13.87,
        "lsat25": 156, "lsat50": 160, "lsat75": 162,
        "gpa25": 3.6, "gpa50": 3.81, "gpa75": 3.98,
    },
    "university-of-nevada-las-vegas-school-of-law": {
        "apps": 1056, "offers": 325, "accept_rate": 30.78,
        "enrollees": 156, "yield_rate": 48.0, "enroll_rate": 14.77,
        "lsat25": 156, "lsat50": 160, "lsat75": 162,
        "gpa25": 3.39, "gpa50": 3.76, "gpa75": 3.88,
    },
    "new-england-law-boston": {
        "apps": 3997, "offers": 2450, "accept_rate": 61.3,
        "enrollees": 446, "yield_rate": 18.16, "enroll_rate": 11.13,
        "lsat25": 149, "lsat50": 153, "lsat75": 157,
        "gpa25": 3.06, "gpa50": 3.37, "gpa75": 3.68,
    },
    "university-of-new-hampshire-school-of-law": {
        "apps": 1455, "offers": 756, "accept_rate": 51.96,
        "enrollees": 245, "yield_rate": 32.41, "enroll_rate": 16.84,
        "lsat25": 151, "lsat50": 155, "lsat75": 159,
        "gpa25": 3.1, "gpa50": 3.5, "gpa75": 3.72,
    },
    "university-of-new-mexico-school-of-law": {
        "apps": 640, "offers": 265, "accept_rate": 41.41,
        "enrollees": 103, "yield_rate": 38.49, "enroll_rate": 15.94,
        "lsat25": 154, "lsat50": 157, "lsat75": 161,
        "gpa25": 3.12, "gpa50": 3.56, "gpa75": 3.82,
    },
    "new-york-law-school": {
        "apps": 4114, "offers": 1632, "accept_rate": 39.67,
        "enrollees": 405, "yield_rate": 24.51, "enroll_rate": 9.72,
        "lsat25": 154, "lsat50": 157, "lsat75": 160,
        "gpa25": 3.41, "gpa50": 3.63, "gpa75": 3.78,
    },
    "nyu-law-school": {
        "apps": 10546, "offers": 1412, "accept_rate": 13.39,
        "enrollees": 430, "yield_rate": 30.45, "enroll_rate": 4.08,
        "lsat25": 169, "lsat50": 172, "lsat75": 174,
        "gpa25": 3.81, "gpa50": 3.92, "gpa75": 3.97,
    },
    "north-carolina-central-law": {
        "apps": 1303, "offers": 509, "accept_rate": 39.06,
        "enrollees": 133, "yield_rate": 25.54, "enroll_rate": 9.98,
        "lsat25": 149, "lsat50": 151, "lsat75": 154,
        "gpa25": 3.05, "gpa50": 3.31, "gpa75": 3.56,
    },
    "unc-law-school": {
        "apps": 3459, "offers": 387, "accept_rate": 11.19,
        "enrollees": 176, "yield_rate": 45.48, "enroll_rate": 5.09,
        "lsat25": 165, "lsat50": 168, "lsat75": 169,
        "gpa25": 3.78, "gpa50": 3.89, "gpa75": 3.97,
    },
    "university-of-north-dakota-school-of-law": {
        "apps": 401, "offers": 253, "accept_rate": 63.09,
        "enrollees": 97, "yield_rate": 38.34, "enroll_rate": 24.19,
        "lsat25": 148, "lsat50": 151, "lsat75": 155,
        "gpa25": 3.24, "gpa50": 3.48, "gpa75": 3.75,
    },
    "unt-dallas-law": {
        "apps": 1883, "offers": 456, "accept_rate": 24.22,
        "enrollees": 127, "yield_rate": 27.85, "enroll_rate": 6.74,
        "lsat25": 151, "lsat50": 153, "lsat75": 156,
        "gpa25": 3.18, "gpa50": 3.41, "gpa75": 3.66,
    },
    "northeastern-law-school": {
        "apps": 5652, "offers": 1369, "accept_rate": 24.22,
        "enrollees": 280, "yield_rate": 20.38, "enroll_rate": 4.94,
        "lsat25": 161, "lsat50": 164, "lsat75": 166,
        "gpa25": 3.61, "gpa50": 3.77, "gpa75": 3.86,
    },
    "northern-illinois-university-law": {
        "apps": 951, "offers": 421, "accept_rate": 44.27,
        "enrollees": 133, "yield_rate": 31.12, "enroll_rate": 13.77,
        "lsat25": 148, "lsat50": 150, "lsat75": 154,
        "gpa25": 3.17, "gpa50": 3.41, "gpa75": 3.64,
    },
    "northern-kentucky-university-law": {
        "apps": 661, "offers": 388, "accept_rate": 58.7,
        "enrollees": 150, "yield_rate": 37.89, "enroll_rate": 22.24,
        "lsat25": 152, "lsat50": 154, "lsat75": 157,
        "gpa25": 3.15, "gpa50": 3.41, "gpa75": 3.68,
    },
    "northwestern-law-school": {
        "apps": 7976, "offers": 981, "accept_rate": 12.3,
        "enrollees": 245, "yield_rate": 24.97, "enroll_rate": 3.07,
        "lsat25": 167, "lsat50": 173, "lsat75": 175,
        "gpa25": 3.76, "gpa50": 3.96, "gpa75": 4.0,
    },
    "notre-dame-law-school": {
        "apps": 3579, "offers": 575, "accept_rate": 16.07,
        "enrollees": 182, "yield_rate": 31.65, "enroll_rate": 5.09,
        "lsat25": 166, "lsat50": 170, "lsat75": 171,
        "gpa25": 3.78, "gpa50": 3.89, "gpa75": 3.95,
    },
    "nova-southeastern-law": {
        "apps": 1458, "offers": 690, "accept_rate": 47.33,
        "enrollees": 224, "yield_rate": 32.17, "enroll_rate": 15.23,
        "lsat25": 152, "lsat50": 155, "lsat75": 158,
        "gpa25": 3.34, "gpa50": 3.52, "gpa75": 3.71,
    },
    "ohio-northern-university-law": {
        "apps": 1355, "offers": 433, "accept_rate": 31.96,
        "enrollees": 129, "yield_rate": 29.79, "enroll_rate": 9.52,
        "lsat25": 146, "lsat50": 149, "lsat75": 151,
        "gpa25": 3.06, "gpa50": 3.34, "gpa75": 3.67,
    },
    "oklahoma-city-university-law": {
        "apps": 715, "offers": 369, "accept_rate": 51.61,
        "enrollees": 156, "yield_rate": 42.01, "enroll_rate": 21.68,
        "lsat25": 149, "lsat50": 151, "lsat75": 155,
        "gpa25": 3.21, "gpa50": 3.56, "gpa75": 3.78,
    },
    "oregon-law": {
        "apps": 1850, "offers": 1017, "accept_rate": 54.97,
        "enrollees": 185, "yield_rate": 18.19, "enroll_rate": 10.0,
        "lsat25": 157, "lsat50": 160, "lsat75": 162,
        "gpa25": 3.32, "gpa50": 3.62, "gpa75": 3.79,
    },
    "pace-university-school-of-law": {
        "apps": 2293, "offers": 1012, "accept_rate": 44.13,
        "enrollees": 323, "yield_rate": 31.82, "enroll_rate": 14.04,
        "lsat25": 151, "lsat50": 154, "lsat75": 157,
        "gpa25": 3.31, "gpa50": 3.54, "gpa75": 3.74,
    },
    "university-of-pacific-law": {
        "apps": 1248, "offers": 679, "accept_rate": 54.41,
        "enrollees": 212, "yield_rate": 30.49, "enroll_rate": 16.59,
        "lsat25": 153, "lsat50": 156, "lsat75": 160,
        "gpa25": 3.17, "gpa50": 3.52, "gpa75": 3.73,
    },
    "penn-state-dickinson-law": {
        "apps": 1912, "offers": 822, "accept_rate": 42.99,
        "enrollees": 208, "yield_rate": 25.3, "enroll_rate": 10.88,
        "lsat25": 157, "lsat50": 159, "lsat75": 161,
        "gpa25": 3.5, "gpa50": 3.71, "gpa75": 3.86,
    },
    "penn-law": {
        "apps": 8074, "offers": 650, "accept_rate": 8.05,
        "enrollees": 254, "yield_rate": 39.08, "enroll_rate": 3.15,
        "lsat25": 167, "lsat50": 173, "lsat75": 174,
        "gpa25": 3.77, "gpa50": 3.95, "gpa75": 4.0,
    },
    "pepperdine-university-school-of-law": {
        "apps": 4272, "offers": 998, "accept_rate": 23.36,
        "enrollees": 223, "yield_rate": 22.34, "enroll_rate": 5.22,
        "lsat25": 161, "lsat50": 164, "lsat75": 167,
        "gpa25": 3.61, "gpa50": 3.85, "gpa75": 3.94,
    },
    "pittsburgh-law-school": {
        "apps": 1606, "offers": 560, "accept_rate": 34.87,
        "enrollees": 155, "yield_rate": 27.68, "enroll_rate": 9.65,
        "lsat25": 158, "lsat50": 160, "lsat75": 163,
        "gpa25": 3.26, "gpa50": 3.54, "gpa75": 3.74,
    },
    "university-of-puerto-rico-law": {
        "apps": 383, "offers": 152, "accept_rate": 39.69,
        "enrollees": 122, "yield_rate": 79.61, "enroll_rate": 31.59,
        "lsat25": 146, "lsat50": 152, "lsat75": 156,
        "gpa25": 3.57, "gpa50": 3.77, "gpa75": 3.9,
    },
    "quinnipiac-university-school-of-law": {
        "apps": 1485, "offers": 640, "accept_rate": 43.1,
        "enrollees": 133, "yield_rate": 20.78, "enroll_rate": 8.96,
        "lsat25": 153, "lsat50": 155, "lsat75": 158,
        "gpa25": 3.41, "gpa50": 3.65, "gpa75": 3.84,
    },
    "regent-university-law": {
        "apps": 537, "offers": 238, "accept_rate": 44.32,
        "enrollees": 126, "yield_rate": 52.94, "enroll_rate": 23.46,
        "lsat25": 153, "lsat50": 158, "lsat75": 161,
        "gpa25": 3.37, "gpa50": 3.68, "gpa75": 3.89,
    },
    "university-of-richmond-school-of-law": {
        "apps": 1738, "offers": 689, "accept_rate": 39.64,
        "enrollees": 146, "yield_rate": 21.19, "enroll_rate": 8.4,
        "lsat25": 159, "lsat50": 164, "lsat75": 167,
        "gpa25": 3.66, "gpa50": 3.81, "gpa75": 3.92,
    },
    "roger-williams-university-school-of-law": {
        "apps": 1003, "offers": 615, "accept_rate": 61.32,
        "enrollees": 188, "yield_rate": 30.24, "enroll_rate": 18.54,
        "lsat25": 147, "lsat50": 150, "lsat75": 154,
        "gpa25": 3.02, "gpa50": 3.43, "gpa75": 3.73,
    },
    "rutgers-law-school": {
        "apps": 4305, "offers": 1444, "accept_rate": 33.54,
        "enrollees": 482, "yield_rate": 33.24, "enroll_rate": 11.15,
        "lsat25": 155, "lsat50": 158, "lsat75": 161,
        "gpa25": 3.35, "gpa50": 3.66, "gpa75": 3.82,
    },
    "saint-louis-university-law": {
        "apps": 941, "offers": 531, "accept_rate": 56.43,
        "enrollees": 221, "yield_rate": 41.43, "enroll_rate": 23.38,
        "lsat25": 154, "lsat50": 157, "lsat75": 161,
        "gpa25": 3.37, "gpa50": 3.61, "gpa75": 3.83,
    },
    "samford-university-cumberland-school-of-law": {
        "apps": 925, "offers": 453, "accept_rate": 48.97,
        "enrollees": 168, "yield_rate": 37.09, "enroll_rate": 18.16,
        "lsat25": 154, "lsat50": 156, "lsat75": 160,
        "gpa25": 3.42, "gpa50": 3.63, "gpa75": 3.84,
    },
    "usd-law-school": {
        "apps": 4241, "offers": 997, "accept_rate": 23.51,
        "enrollees": 239, "yield_rate": 23.97, "enroll_rate": 5.64,
        "lsat25": 160, "lsat50": 163, "lsat75": 165,
        "gpa25": 3.59, "gpa50": 3.84, "gpa75": 3.92,
    },
    "university-of-san-francisco-law": {
        "apps": 2100, "offers": 890, "accept_rate": 42.38,
        "enrollees": 198, "yield_rate": 22.25, "enroll_rate": 9.43,
        "lsat25": 152, "lsat50": 155, "lsat75": 157,
        "gpa25": 3.16, "gpa50": 3.47, "gpa75": 3.68,
    },
    "santa-clara-university-school-of-law": {
        "apps": 2575, "offers": 1354, "accept_rate": 52.58,
        "enrollees": 253, "yield_rate": 18.69, "enroll_rate": 9.83,
        "lsat25": 157, "lsat50": 159, "lsat75": 162,
        "gpa25": 3.21, "gpa50": 3.5, "gpa75": 3.74,
    },
    "seattle-university-school-of-law": {
        "apps": 2304, "offers": 897, "accept_rate": 38.93,
        "enrollees": 209, "yield_rate": 23.3, "enroll_rate": 9.07,
        "lsat25": 155, "lsat50": 159, "lsat75": 162,
        "gpa25": 3.22, "gpa50": 3.6, "gpa75": 3.8,
    },
    "seton-hall-university-school-of-law": {
        "apps": 2714, "offers": 930, "accept_rate": 34.27,
        "enrollees": 300, "yield_rate": 32.26, "enroll_rate": 11.05,
        "lsat25": 158, "lsat50": 161, "lsat75": 164,
        "gpa25": 3.51, "gpa50": 3.71, "gpa75": 3.86,
    },
    "university-of-south-carolina-school-of-law": {
        "apps": 2114, "offers": 632, "accept_rate": 29.9,
        "enrollees": 215, "yield_rate": 34.02, "enroll_rate": 10.17,
        "lsat25": 160, "lsat50": 162, "lsat75": 164,
        "gpa25": 3.43, "gpa50": 3.67, "gpa75": 3.86,
    },
    "university-of-south-dakota-school-of-law": {
        "apps": 405, "offers": 253, "accept_rate": 62.47,
        "enrollees": 85, "yield_rate": 32.81, "enroll_rate": 20.49,
        "lsat25": 149, "lsat50": 152, "lsat75": 156,
        "gpa25": 3.16, "gpa50": 3.47, "gpa75": 3.81,
    },
    "south-texas-college-of-law-houston": {
        "apps": 2770, "offers": 932, "accept_rate": 33.65,
        "enrollees": 437, "yield_rate": 46.57, "enroll_rate": 15.67,
        "lsat25": 153, "lsat50": 155, "lsat75": 158,
        "gpa25": 3.08, "gpa50": 3.41, "gpa75": 3.65,
    },
    "usc-gould-school-of-law": {
        "apps": 7204, "offers": 805, "accept_rate": 11.17,
        "enrollees": 227, "yield_rate": 28.2, "enroll_rate": 3.15,
        "lsat25": 165, "lsat50": 169, "lsat75": 170,
        "gpa25": 3.73, "gpa50": 3.91, "gpa75": 3.97,
    },
    "southern-illinois-university-law": {
        "apps": 724, "offers": 407, "accept_rate": 56.22,
        "enrollees": 132, "yield_rate": 31.7, "enroll_rate": 17.82,
        "lsat25": 146, "lsat50": 149, "lsat75": 155,
        "gpa25": 2.93, "gpa50": 3.38, "gpa75": 3.78,
    },
    "smu-dedman-school-of-law": {
        "apps": 3038, "offers": 747, "accept_rate": 24.59,
        "enrollees": 254, "yield_rate": 34.0, "enroll_rate": 8.36,
        "lsat25": 163, "lsat50": 167, "lsat75": 168,
        "gpa25": 3.52, "gpa50": 3.81, "gpa75": 3.92,
    },
    "southern-university-law": {
        "apps": 1153, "offers": 652, "accept_rate": 56.55,
        "enrollees": 297, "yield_rate": 43.1, "enroll_rate": 24.37,
        "lsat25": 145, "lsat50": 147, "lsat75": 150,
        "gpa25": 2.86, "gpa50": 3.15, "gpa75": 3.48,
    },
    "southwestern-law-school": {
        "apps": 3584, "offers": 1307, "accept_rate": 36.47,
        "enrollees": 458, "yield_rate": 34.74, "enroll_rate": 12.67,
        "lsat25": 154, "lsat50": 157, "lsat75": 160,
        "gpa25": 3.23, "gpa50": 3.51, "gpa75": 3.72,
    },
    "st-john-s-university-school-of-law": {
        "apps": 4417, "offers": 1018, "accept_rate": 23.05,
        "enrollees": 238, "yield_rate": 23.38, "enroll_rate": 5.39,
        "lsat25": 156, "lsat50": 164, "lsat75": 166,
        "gpa25": 3.44, "gpa50": 3.81, "gpa75": 3.91,
    },
    "st-mary-s-university-school-of-law": {
        "apps": 2906, "offers": 755, "accept_rate": 25.98,
        "enrollees": 279, "yield_rate": 36.16, "enroll_rate": 9.39,
        "lsat25": 151, "lsat50": 153, "lsat75": 156,
        "gpa25": 3.14, "gpa50": 3.49, "gpa75": 3.69,
    },
    "st-thomas-university-miami-law": {
        "apps": 1771, "offers": 1067, "accept_rate": 60.25,
        "enrollees": 424, "yield_rate": 38.33, "enroll_rate": 23.09,
        "lsat25": 150, "lsat50": 152, "lsat75": 156,
        "gpa25": 3.12, "gpa50": 3.42, "gpa75": 3.65,
    },
    "st-thomas-minneapolis-law": {
        "apps": 847, "offers": 488, "accept_rate": 57.62,
        "enrollees": 164, "yield_rate": 33.61, "enroll_rate": 19.36,
        "lsat25": 155, "lsat50": 159, "lsat75": 162,
        "gpa25": 3.46, "gpa50": 3.72, "gpa75": 3.88,
    },
    "stanford-law-school": {
        "apps": 5526, "offers": 337, "accept_rate": 6.1,
        "enrollees": 173, "yield_rate": 51.34, "enroll_rate": 3.13,
        "lsat25": 171, "lsat50": 173, "lsat75": 176,
        "gpa25": 3.87, "gpa50": 3.96, "gpa75": 4.0,
    },
    "stetson-university-law": {
        "apps": 2791, "offers": 1017, "accept_rate": 36.44,
        "enrollees": 333, "yield_rate": 32.74, "enroll_rate": 11.93,
        "lsat25": 157, "lsat50": 159, "lsat75": 162,
        "gpa25": 3.42, "gpa50": 3.63, "gpa75": 3.81,
    },
    "suffolk-university-law-school": {
        "apps": 3481, "offers": 2153, "accept_rate": 61.85,
        "enrollees": 541, "yield_rate": 25.03, "enroll_rate": 15.48,
        "lsat25": 152, "lsat50": 155, "lsat75": 159,
        "gpa25": 3.37, "gpa50": 3.58, "gpa75": 3.79,
    },
    "syracuse-university-law": {
        "apps": 2752, "offers": 952, "accept_rate": 34.59,
        "enrollees": 268, "yield_rate": 27.94, "enroll_rate": 9.67,
        "lsat25": 154, "lsat50": 158, "lsat75": 160,
        "gpa25": 3.18, "gpa50": 3.6, "gpa75": 3.77,
    },
    "temple-university-beasley-school-of-law": {
        "apps": 3043, "offers": 742, "accept_rate": 24.38,
        "enrollees": 232, "yield_rate": 31.27, "enroll_rate": 7.62,
        "lsat25": 162, "lsat50": 165, "lsat75": 168,
        "gpa25": 3.59, "gpa50": 3.76, "gpa75": 3.9,
    },
    "university-of-tennessee-college-of-law": {
        "apps": 2176, "offers": 539, "accept_rate": 24.77,
        "enrollees": 139, "yield_rate": 25.79, "enroll_rate": 6.39,
        "lsat25": 160, "lsat50": 164, "lsat75": 166,
        "gpa25": 3.67, "gpa50": 3.89, "gpa75": 3.95,
    },
    "texas-am-law-school": {
        "apps": 5240, "offers": 634, "accept_rate": 12.1,
        "enrollees": 121, "yield_rate": 19.09, "enroll_rate": 2.31,
        "lsat25": 161, "lsat50": 169, "lsat75": 170,
        "gpa25": 3.67, "gpa50": 4.0, "gpa75": 4.0,
    },
    "texas-southern-university-law": {
        "apps": 2032, "offers": 673, "accept_rate": 33.12,
        "enrollees": 219, "yield_rate": 30.76, "enroll_rate": 10.19,
        "lsat25": 148, "lsat50": 150, "lsat75": 152,
        "gpa25": 2.99, "gpa50": 3.33, "gpa75": 3.62,
    },
    "texas-tech-university-school-of-law": {
        "apps": 2011, "offers": 515, "accept_rate": 25.61,
        "enrollees": 168, "yield_rate": 31.65, "enroll_rate": 8.11,
        "lsat25": 157, "lsat50": 159, "lsat75": 160,
        "gpa25": 3.41, "gpa50": 3.69, "gpa75": 3.85,
    },
    "ut-austin-law-school": {
        "apps": 6297, "offers": 881, "accept_rate": 13.99,
        "enrollees": 361, "yield_rate": 40.98, "enroll_rate": 5.73,
        "lsat25": 166, "lsat50": 172, "lsat75": 173,
        "gpa25": 3.75, "gpa50": 3.89, "gpa75": 3.96,
    },
    "moritz-college-of-law": {
        "apps": 2291, "offers": 564, "accept_rate": 24.62,
        "enrollees": 158, "yield_rate": 28.01, "enroll_rate": 6.9,
        "lsat25": 163, "lsat50": 168, "lsat75": 169,
        "gpa25": 3.64, "gpa50": 3.91, "gpa75": 3.97,
    },
    "university-of-toledo-college-of-law": {
        "apps": 732, "offers": 384, "accept_rate": 52.46,
        "enrollees": 125, "yield_rate": 32.55, "enroll_rate": 17.08,
        "lsat25": 150, "lsat50": 153, "lsat75": 156,
        "gpa25": 3.37, "gpa50": 3.53, "gpa75": 3.79,
    },
    "touro-law-center": {
        "apps": 1462, "offers": 667, "accept_rate": 45.62,
        "enrollees": 222, "yield_rate": 33.28, "enroll_rate": 15.18,
        "lsat25": 151, "lsat50": 153, "lsat75": 156,
        "gpa25": 3.05, "gpa50": 3.35, "gpa75": 3.66,
    },
    "tulane-university-law-school": {
        "apps": 2069, "offers": 827, "accept_rate": 39.97,
        "enrollees": 240, "yield_rate": 28.9, "enroll_rate": 11.55,
        "lsat25": 159, "lsat50": 161, "lsat75": 164,
        "gpa25": 3.47, "gpa50": 3.67, "gpa75": 3.81,
    },
    "tulsa-university-law": {
        "apps": 509, "offers": 288, "accept_rate": 56.58,
        "enrollees": 118, "yield_rate": 40.97, "enroll_rate": 23.18,
        "lsat25": 155, "lsat50": 157, "lsat75": 160,
        "gpa25": 3.27, "gpa50": 3.45, "gpa75": 3.72,
    },
    "university-of-utah-sjd": {
        "apps": 1495, "offers": 378, "accept_rate": 25.28,
        "enrollees": 106, "yield_rate": 28.04, "enroll_rate": 7.09,
        "lsat25": 162, "lsat50": 166, "lsat75": 168,
        "gpa25": 3.62, "gpa50": 3.87, "gpa75": 3.94,
    },
    "vanderbilt-law-school": {
        "apps": 6124, "offers": 812, "accept_rate": 13.26,
        "enrollees": 170, "yield_rate": 20.94, "enroll_rate": 2.78,
        "lsat25": 167, "lsat50": 170, "lsat75": 171,
        "gpa25": 3.77, "gpa50": 3.91, "gpa75": 3.97,
    },
    "vermont-law-school": {
        "apps": 1553, "offers": 909, "accept_rate": 58.53,
        "enrollees": 289, "yield_rate": 31.79, "enroll_rate": 18.61,
        "lsat25": 148, "lsat50": 152, "lsat75": 157,
        "gpa25": 2.81, "gpa50": 3.25, "gpa75": 3.61,
    },
    "villanova-university-charles-widger-school-of-law": {
        "apps": 2889, "offers": 749, "accept_rate": 25.93,
        "enrollees": 231, "yield_rate": 30.84, "enroll_rate": 8.0,
        "lsat25": 159, "lsat50": 164, "lsat75": 166,
        "gpa25": 3.64, "gpa50": 3.8, "gpa75": 3.9,
    },
    "uva-law-school": {
        "apps": 6734, "offers": 685, "accept_rate": 10.17,
        "enrollees": 302, "yield_rate": 44.09, "enroll_rate": 4.48,
        "lsat25": 168, "lsat50": 173, "lsat75": 175,
        "gpa25": 3.83, "gpa50": 3.99, "gpa75": 4.04,
    },
    "wake-forest-law-school": {
        "apps": 2883, "offers": 710, "accept_rate": 24.63,
        "enrollees": 201, "yield_rate": 28.31, "enroll_rate": 6.97,
        "lsat25": 163, "lsat50": 166, "lsat75": 168,
        "gpa25": 3.6, "gpa50": 3.79, "gpa75": 3.9,
    },
    "washburn-university-school-of-law": {
        "apps": 549, "offers": 255, "accept_rate": 46.45,
        "enrollees": 136, "yield_rate": 53.33, "enroll_rate": 24.77,
        "lsat25": 151, "lsat50": 155, "lsat75": 159,
        "gpa25": 3.37, "gpa50": 3.68, "gpa75": 3.88,
    },
    "washington-and-lee-law-school": {
        "apps": 2341, "offers": 532, "accept_rate": 22.73,
        "enrollees": 128, "yield_rate": 24.06, "enroll_rate": 5.47,
        "lsat25": 161, "lsat50": 167, "lsat75": 168,
        "gpa25": 3.51, "gpa50": 3.75, "gpa75": 3.85,
    },
    "washington-university-school-of-law": {
        "apps": 5616, "offers": 1065, "accept_rate": 18.96,
        "enrollees": 258, "yield_rate": 24.23, "enroll_rate": 4.59,
        "lsat25": 165, "lsat50": 175, "lsat75": 176,
        "gpa25": 3.58, "gpa50": 3.96, "gpa75": 4.0,
    },
    "university-of-washington-school-of-law": {
        "apps": 3412, "offers": 669, "accept_rate": 19.61,
        "enrollees": 202, "yield_rate": 30.19, "enroll_rate": 5.92,
        "lsat25": 162, "lsat50": 165, "lsat75": 167,
        "gpa25": 3.6, "gpa50": 3.76, "gpa75": 3.86,
    },
    "wayne-state-university-law": {
        "apps": 1290, "offers": 320, "accept_rate": 24.81,
        "enrollees": 128, "yield_rate": 40.0, "enroll_rate": 9.92,
        "lsat25": 157, "lsat50": 164, "lsat75": 166,
        "gpa25": 3.56, "gpa50": 3.89, "gpa75": 3.95,
    },
    "west-virginia-university-college-of-law": {
        "apps": 590, "offers": 238, "accept_rate": 40.34,
        "enrollees": 114, "yield_rate": 47.9, "enroll_rate": 19.32,
        "lsat25": 151, "lsat50": 155, "lsat75": 159,
        "gpa25": 3.34, "gpa50": 3.65, "gpa75": 3.89,
    },
    "western-new-england-university-school-of-law": {
        "apps": 871, "offers": 452, "accept_rate": 51.89,
        "enrollees": 112, "yield_rate": 24.78, "enroll_rate": 12.86,
        "lsat25": 149, "lsat50": 153, "lsat75": 156,
        "gpa25": 2.84, "gpa50": 3.27, "gpa75": 3.57,
    },
    "western-state-law": {
        "apps": 916, "offers": 400, "accept_rate": 43.67,
        "enrollees": 113, "yield_rate": 28.25, "enroll_rate": 12.34,
        "lsat25": 150, "lsat50": 152, "lsat75": 156,
        "gpa25": 3.03, "gpa50": 3.36, "gpa75": 3.62,
    },
    "widener-commonwealth-law": {
        "apps": 909, "offers": 549, "accept_rate": 60.4,
        "enrollees": 156, "yield_rate": 28.23, "enroll_rate": 17.05,
        "lsat25": 147, "lsat50": 150, "lsat75": 153,
        "gpa25": 2.98, "gpa50": 3.33, "gpa75": 3.64,
    },
    "widener-delaware-law": {
        "apps": 1317, "offers": 786, "accept_rate": 59.68,
        "enrollees": 276, "yield_rate": 35.11, "enroll_rate": 20.96,
        "lsat25": 149, "lsat50": 152, "lsat75": 154,
        "gpa25": 2.95, "gpa50": 3.33, "gpa75": 3.6,
    },
    "willamette-university-college-of-law": {
        "apps": 858, "offers": 527, "accept_rate": 61.42,
        "enrollees": 139, "yield_rate": 26.38, "enroll_rate": 16.2,
        "lsat25": 151, "lsat50": 154, "lsat75": 157,
        "gpa25": 3.22, "gpa50": 3.48, "gpa75": 3.7,
    },
    "william-and-mary-law-school": {
        "apps": 2176, "offers": 478, "accept_rate": 21.97,
        "enrollees": 181, "yield_rate": 37.87, "enroll_rate": 8.32,
        "lsat25": 161, "lsat50": 166, "lsat75": 167,
        "gpa25": 3.53, "gpa50": 3.82, "gpa75": 3.94,
    },
    "wilmington-university-law": {
        "apps": 273, "offers": 142, "accept_rate": 52.01,
        "enrollees": 73, "yield_rate": 48.59, "enroll_rate": 25.27,
        "lsat25": 146, "lsat50": 149, "lsat75": 152,
        "gpa25": 2.67, "gpa50": 3.13, "gpa75": 3.49,
    },
    "wisconsin-law-school": {
        "apps": 2479, "offers": 474, "accept_rate": 19.12,
        "enrollees": 220, "yield_rate": 46.41, "enroll_rate": 8.87,
        "lsat25": 161, "lsat50": 167, "lsat75": 169,
        "gpa25": 3.61, "gpa50": 3.81, "gpa75": 3.93,
    },
    "university-of-wyoming-college-of-law": {
        "apps": 483, "offers": 173, "accept_rate": 35.82,
        "enrollees": 71, "yield_rate": 41.04, "enroll_rate": 14.7,
        "lsat25": 153, "lsat50": 156, "lsat75": 160,
        "gpa25": 3.32, "gpa50": 3.55, "gpa75": 3.8,
    },
    "yale-law-school": {
        "apps": 5562, "offers": 226, "accept_rate": 4.06,
        "enrollees": 175, "yield_rate": 77.43, "enroll_rate": 3.15,
        "lsat25": 171, "lsat50": 174, "lsat75": 177,
        "gpa25": 3.9, "gpa50": 3.96, "gpa75": 4.0,
    },
}

# --- GRANT DATA (from ABA 509 reports) ---
grant_data = {
    "university-of-akron-school-of-law": {
        "total_students": 437, "total_receiving": 411, "pct_receiving": 94.0,
        "lt_half_n": 258, "lt_half_pct": 59.0,
        "half_to_full_n": 144, "half_to_full_pct": 33.0,
        "full_n": 6, "full_pct": 1.0,
        "gt_full_n": 3, "gt_full_pct": 1.0,
        "p25": 3000, "p50": 9000, "p75": 18300,
    },
    "university-of-alabama-school-of-law": {
        "total_students": 395, "total_receiving": 380, "pct_receiving": 96.0,
        "lt_half_n": 119, "lt_half_pct": 30.0,
        "half_to_full_n": 161, "half_to_full_pct": 41.0,
        "full_n": 24, "full_pct": 6.0,
        "gt_full_n": 76, "gt_full_pct": 19.0,
        "p25": 12000, "p50": 21500, "p75": 30000,
    },
    "albany-law-school": {
        "total_students": 621, "total_receiving": 584, "pct_receiving": 94.0,
        "lt_half_n": 232, "lt_half_pct": 37.0,
        "half_to_full_n": 294, "half_to_full_pct": 47.0,
        "full_n": 40, "full_pct": 6.0,
        "gt_full_n": 18, "gt_full_pct": 3.0,
        "p25": 20000, "p50": 38000, "p75": 46000,
    },
    "american-university-law-school": {
        "total_students": 1232, "total_receiving": 732, "pct_receiving": 59.0,
        "lt_half_n": 161, "lt_half_pct": 13.0,
        "half_to_full_n": 493, "half_to_full_pct": 40.0,
        "full_n": 68, "full_pct": 6.0,
        "gt_full_n": 10, "gt_full_pct": 1.0,
        "p25": 40000, "p50": 40000, "p75": 40000,
    },
    "appalachian-school-of-law": {
        "total_students": 163, "total_receiving": 141, "pct_receiving": 87.0,
        "lt_half_n": 89, "lt_half_pct": 55.0,
        "half_to_full_n": 31, "half_to_full_pct": 19.0,
        "full_n": 11, "full_pct": 7.0,
        "gt_full_n": 10, "gt_full_pct": 6.0,
        "p25": 5500, "p50": 10250, "p75": 24600,
    },
    "asu-law": {
        "total_students": 749, "total_receiving": 722, "pct_receiving": 96.0,
        "lt_half_n": 332, "lt_half_pct": 44.0,
        "half_to_full_n": 293, "half_to_full_pct": 39.0,
        "full_n": 79, "full_pct": 11.0,
        "gt_full_n": 18, "gt_full_pct": 2.0,
        "p25": 5000, "p50": 20000, "p75": 33336,
    },
    "university-of-arizona-law-school": {
        "total_students": 355, "total_receiving": 311, "pct_receiving": 88.0,
        "lt_half_n": 39, "lt_half_pct": 11.0,
        "half_to_full_n": 115, "half_to_full_pct": 32.0,
        "full_n": 68, "full_pct": 19.0,
        "gt_full_n": 89, "gt_full_pct": 25.0,
        "p25": 15018, "p50": 25235, "p75": 29870,
    },
    "arkansas-little-rock-law": {
        "total_students": 420, "total_receiving": 251, "pct_receiving": 60.0,
        "lt_half_n": 117, "lt_half_pct": 28.0,
        "half_to_full_n": 97, "half_to_full_pct": 23.0,
        "full_n": 24, "full_pct": 6.0,
        "gt_full_n": 13, "gt_full_pct": 3.0,
        "p25": 3443, "p50": 7461, "p75": 13770,
    },
    "university-of-arkansas-school-of-law": {
        "total_students": 368, "total_receiving": 320, "pct_receiving": 87.0,
        "lt_half_n": 177, "lt_half_pct": 48.0,
        "half_to_full_n": 135, "half_to_full_pct": 37.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 8, "gt_full_pct": 2.0,
        "p25": 5500, "p50": 9000, "p75": 11000,
    },
    "john-marshall-atlanta": {
        "total_students": 373, "total_receiving": 314, "pct_receiving": 84.0,
        "lt_half_n": 297, "lt_half_pct": 80.0,
        "half_to_full_n": 1, "half_to_full_pct": None,
        "full_n": 16, "full_pct": 4.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 1125, "p50": 4500, "p75": 9000,
    },
    "ave-maria-school-of-law": {
        "total_students": 337, "total_receiving": 247, "pct_receiving": 73.0,
        "lt_half_n": 73, "lt_half_pct": 22.0,
        "half_to_full_n": 162, "half_to_full_pct": 48.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 12, "gt_full_pct": 4.0,
        "p25": 7365, "p50": 30000, "p75": 47400,
    },
    "university-of-baltimore-school-of-law": {
        "total_students": 700, "total_receiving": 471, "pct_receiving": 67.0,
        "lt_half_n": 318, "lt_half_pct": 45.0,
        "half_to_full_n": 151, "half_to_full_pct": 22.0,
        "full_n": 5, "full_pct": 1.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 7648, "p50": 15000, "p75": 22000,
    },
    "barry-university-law": {
        "total_students": 775, "total_receiving": 609, "pct_receiving": 79.0,
        "lt_half_n": 448, "lt_half_pct": 58.0,
        "half_to_full_n": 150, "half_to_full_pct": 19.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 11, "gt_full_pct": 1.0,
        "p25": 6500, "p50": 16000, "p75": 22000,
    },
    "baylor-law-school": {
        "total_students": 451, "total_receiving": 437, "pct_receiving": 97.0,
        "lt_half_n": 186, "lt_half_pct": 41.0,
        "half_to_full_n": 131, "half_to_full_pct": 29.0,
        "full_n": 89, "full_pct": 20.0,
        "gt_full_n": 31, "gt_full_pct": 7.0,
        "p25": 18000, "p50": 33518, "p75": 57873,
    },
    "belmont-university-law": {
        "total_students": 367, "total_receiving": 227, "pct_receiving": 62.0,
        "lt_half_n": 34, "lt_half_pct": 9.0,
        "half_to_full_n": 89, "half_to_full_pct": 24.0,
        "full_n": 104, "full_pct": 28.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 27140, "p50": 40710, "p75": 54280,
    },
    "boston-college-law-school": {
        "total_students": 635, "total_receiving": 566, "pct_receiving": 89.0,
        "lt_half_n": 473, "lt_half_pct": 74.0,
        "half_to_full_n": 85, "half_to_full_pct": 13.0,
        "full_n": 8, "full_pct": 1.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 20000, "p50": 24000, "p75": 28000,
    },
    "boston-university-law-school": {
        "total_students": 662, "total_receiving": 604, "pct_receiving": 91.0,
        "lt_half_n": 382, "lt_half_pct": 58.0,
        "half_to_full_n": 179, "half_to_full_pct": 27.0,
        "full_n": 42, "full_pct": 6.0,
        "gt_full_n": 1, "gt_full_pct": None,
        "p25": 20000, "p50": 29000, "p75": 43000,
    },
    "byu-law-school": {
        "total_students": 350, "total_receiving": 344, "pct_receiving": 98.0,
        "lt_half_n": 138, "lt_half_pct": 39.0,
        "half_to_full_n": 48, "half_to_full_pct": 14.0,
        "full_n": 72, "full_pct": 21.0,
        "gt_full_n": 86, "gt_full_pct": 25.0,
        "p25": 3882, "p50": 11646, "p75": 15528,
    },
    "brooklyn-law-school": {
        "total_students": 1114, "total_receiving": 1017, "pct_receiving": 91.0,
        "lt_half_n": 503, "lt_half_pct": 45.0,
        "half_to_full_n": 466, "half_to_full_pct": 42.0,
        "full_n": 41, "full_pct": 4.0,
        "gt_full_n": 7, "gt_full_pct": 1.0,
        "p25": 25000, "p50": 34000, "p75": 48000,
    },
    "university-of-buffalo-law": {
        "total_students": 412, "total_receiving": 212, "pct_receiving": 51.0,
        "lt_half_n": 168, "lt_half_pct": 41.0,
        "half_to_full_n": 44, "half_to_full_pct": 11.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 3500, "p50": 7500, "p75": 11500,
    },
    "california-western-school-of-law": {
        "total_students": 650, "total_receiving": 392, "pct_receiving": 60.0,
        "lt_half_n": 157, "lt_half_pct": 24.0,
        "half_to_full_n": 158, "half_to_full_pct": 24.0,
        "full_n": 77, "full_pct": 12.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 18720, "p50": 37440, "p75": 53040,
    },
    "uc-berkeley-law-school": {
        "total_students": 1012, "total_receiving": 869, "pct_receiving": 86.0,
        "lt_half_n": 491, "lt_half_pct": 49.0,
        "half_to_full_n": 283, "half_to_full_pct": 28.0,
        "full_n": 4, "full_pct": None,
        "gt_full_n": 91, "gt_full_pct": 9.0,
        "p25": 20000, "p50": 31158, "p75": 42245,
    },
    "uc-davis-law-school": {
        "total_students": 605, "total_receiving": 598, "pct_receiving": 99.0,
        "lt_half_n": 350, "lt_half_pct": 58.0,
        "half_to_full_n": 243, "half_to_full_pct": 40.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 5, "gt_full_pct": 1.0,
        "p25": 14500, "p50": 25000, "p75": 35000,
    },
    "uc-irvine-law": {
        "total_students": 474, "total_receiving": 451, "pct_receiving": 95.0,
        "lt_half_n": 371, "lt_half_pct": 78.0,
        "half_to_full_n": 80, "half_to_full_pct": 17.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 17500, "p50": 25000, "p75": 35000,
    },
    "ucla-law": {
        "total_students": 994, "total_receiving": 825, "pct_receiving": 83.0,
        "lt_half_n": 498, "lt_half_pct": 50.0,
        "half_to_full_n": 240, "half_to_full_pct": 24.0,
        "full_n": 62, "full_pct": 6.0,
        "gt_full_n": 25, "gt_full_pct": 3.0,
        "p25": 15000, "p50": 24456, "p75": 43333,
    },
    "uc-law-san-francisco": {
        "total_students": 1087, "total_receiving": 936, "pct_receiving": 86.0,
        "lt_half_n": 722, "lt_half_pct": 66.0,
        "half_to_full_n": 202, "half_to_full_pct": 19.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 12, "gt_full_pct": 1.0,
        "p25": 10000, "p50": 20000, "p75": 25000,
    },
    "campbell-university-law": {
        "total_students": 545, "total_receiving": 539, "pct_receiving": 99.0,
        "lt_half_n": 348, "lt_half_pct": 64.0,
        "half_to_full_n": 188, "half_to_full_pct": 34.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 2, "gt_full_pct": None,
        "p25": 15000, "p50": 21771, "p75": 29850,
    },
    "capital-university-law": {
        "total_students": 461, "total_receiving": 444, "pct_receiving": 96.0,
        "lt_half_n": 263, "lt_half_pct": 57.0,
        "half_to_full_n": 181, "half_to_full_pct": 39.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 12000, "p50": 21000, "p75": 32000,
    },
    "cardozo-school-of-law": {
        "total_students": 933, "total_receiving": 917, "pct_receiving": 98.0,
        "lt_half_n": 423, "lt_half_pct": 45.0,
        "half_to_full_n": 470, "half_to_full_pct": 50.0,
        "full_n": 23, "full_pct": 2.0,
        "gt_full_n": 1, "gt_full_pct": None,
        "p25": 30000, "p50": 40000, "p75": 50000,
    },
    "case-western-reserve-university-school-of-law": {
        "total_students": 508, "total_receiving": 496, "pct_receiving": 98.0,
        "lt_half_n": 108, "lt_half_pct": 21.0,
        "half_to_full_n": 388, "half_to_full_pct": 76.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 31000, "p50": 42500, "p75": 48500,
    },
    "catholic-university-law": {
        "total_students": 374, "total_receiving": 369, "pct_receiving": 99.0,
        "lt_half_n": 185, "lt_half_pct": 49.0,
        "half_to_full_n": 183, "half_to_full_pct": 49.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 1, "gt_full_pct": None,
        "p25": 15727, "p50": 26000, "p75": 45000,
    },
    "chapman-university-fowler-school-of-law": {
        "total_students": 404, "total_receiving": 237, "pct_receiving": 59.0,
        "lt_half_n": 49, "lt_half_pct": 12.0,
        "half_to_full_n": 128, "half_to_full_pct": 32.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 60, "gt_full_pct": 15.0,
        "p25": 35000, "p50": 59500, "p75": 64500,
    },
    "charleston-school-of-law": {
        "total_students": 697, "total_receiving": 661, "pct_receiving": 95.0,
        "lt_half_n": 405, "lt_half_pct": 58.0,
        "half_to_full_n": 236, "half_to_full_pct": 34.0,
        "full_n": 19, "full_pct": 3.0,
        "gt_full_n": 1, "gt_full_pct": None,
        "p25": 13000, "p50": 20000, "p75": 28500,
    },
    "chicago-law-school": {
        "total_students": 620, "total_receiving": 485, "pct_receiving": 78.0,
        "lt_half_n": 396, "lt_half_pct": 64.0,
        "half_to_full_n": 31, "half_to_full_pct": 5.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 58, "gt_full_pct": 9.0,
        "p25": 10000, "p50": 15000, "p75": 30000,
    },
    "chicago-kent-college-of-law": {
        "total_students": 728, "total_receiving": 669, "pct_receiving": 92.0,
        "lt_half_n": 192, "lt_half_pct": 26.0,
        "half_to_full_n": 448, "half_to_full_pct": 62.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 29, "gt_full_pct": 4.0,
        "p25": 25000, "p50": 35000, "p75": 43862,
    },
    "university-of-cincinnati-college-of-law": {
        "total_students": 374, "total_receiving": 375, "pct_receiving": 100.0,
        "lt_half_n": 119, "lt_half_pct": 32.0,
        "half_to_full_n": 215, "half_to_full_pct": 57.0,
        "full_n": 13, "full_pct": 3.0,
        "gt_full_n": 28, "gt_full_pct": 7.0,
        "p25": 11700, "p50": 19000, "p75": 23400,
    },
    "cuny-law-school": {
        "total_students": 670, "total_receiving": 112, "pct_receiving": 17.0,
        "lt_half_n": 73, "lt_half_pct": 11.0,
        "half_to_full_n": 38, "half_to_full_pct": 6.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 1, "gt_full_pct": None,
        "p25": 8000, "p50": 8000, "p75": 8000,
    },
    "cleveland-state-law": {
        "total_students": 467, "total_receiving": 456, "pct_receiving": 98.0,
        "lt_half_n": 321, "lt_half_pct": 69.0,
        "half_to_full_n": 134, "half_to_full_pct": 29.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 6000, "p50": 11000, "p75": 18000,
    },
    "colorado-law": {
        "total_students": 511, "total_receiving": 444, "pct_receiving": 87.0,
        "lt_half_n": 293, "lt_half_pct": 57.0,
        "half_to_full_n": 134, "half_to_full_pct": 26.0,
        "full_n": 11, "full_pct": 2.0,
        "gt_full_n": 6, "gt_full_pct": 1.0,
        "p25": 10300, "p50": 15000, "p75": 27440,
    },
    "columbia-law-school": {
        "total_students": 1353, "total_receiving": 728, "pct_receiving": 54.0,
        "lt_half_n": 486, "lt_half_pct": 36.0,
        "half_to_full_n": 173, "half_to_full_pct": 13.0,
        "full_n": 67, "full_pct": 5.0,
        "gt_full_n": 2, "gt_full_pct": None,
        "p25": 20000, "p50": 32000, "p75": 48000,
    },
    "uconn-law-school": {
        "total_students": 471, "total_receiving": 428, "pct_receiving": 91.0,
        "lt_half_n": 291, "lt_half_pct": 62.0,
        "half_to_full_n": 125, "half_to_full_pct": 27.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 12, "gt_full_pct": 3.0,
        "p25": 10000, "p50": 15177, "p75": 22034,
    },
    "cooley-law-school": {
        "total_students": 457, "total_receiving": 260, "pct_receiving": 57.0,
        "lt_half_n": 195, "lt_half_pct": 43.0,
        "half_to_full_n": 59, "half_to_full_pct": 13.0,
        "full_n": 6, "full_pct": 1.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 9279, "p50": 13746, "p75": 23463,
    },
    "cornell-law-school": {
        "total_students": 600, "total_receiving": 544, "pct_receiving": 91.0,
        "lt_half_n": 402, "lt_half_pct": 67.0,
        "half_to_full_n": 133, "half_to_full_pct": 22.0,
        "full_n": 9, "full_pct": 2.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 28000, "p50": 30000, "p75": 45000,
    },
    "creighton-university-law": {
        "total_students": 384, "total_receiving": 343, "pct_receiving": 89.0,
        "lt_half_n": 173, "lt_half_pct": 45.0,
        "half_to_full_n": 109, "half_to_full_pct": 28.0,
        "full_n": 49, "full_pct": 13.0,
        "gt_full_n": 12, "gt_full_pct": 3.0,
        "p25": 15200, "p50": 25600, "p75": 37458,
    },
    "university-of-dayton-school-of-law": {
        "total_students": 377, "total_receiving": 384, "pct_receiving": 102.0,
        "lt_half_n": 62, "lt_half_pct": 16.0,
        "half_to_full_n": 301, "half_to_full_pct": 80.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 21, "gt_full_pct": 6.0,
        "p25": 28000, "p50": 33000, "p75": 37000,
    },
    "denver-law-school": {
        "total_students": 866, "total_receiving": 679, "pct_receiving": 78.0,
        "lt_half_n": 349, "lt_half_pct": 40.0,
        "half_to_full_n": 202, "half_to_full_pct": 23.0,
        "full_n": 98, "full_pct": 11.0,
        "gt_full_n": 30, "gt_full_pct": 3.0,
        "p25": 23000, "p50": 30000, "p75": 41000,
    },
    "depaul-university-law": {
        "total_students": 495, "total_receiving": 491, "pct_receiving": 99.0,
        "lt_half_n": 190, "lt_half_pct": 38.0,
        "half_to_full_n": 294, "half_to_full_pct": 59.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 6, "gt_full_pct": 1.0,
        "p25": 20000, "p50": 34950, "p75": 35000,
    },
    "university-of-detroit-mercy-school-of-law": {
        "total_students": 628, "total_receiving": 364, "pct_receiving": 58.0,
        "lt_half_n": 178, "lt_half_pct": 28.0,
        "half_to_full_n": 180, "half_to_full_pct": 29.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 5, "gt_full_pct": 1.0,
        "p25": 10504, "p50": 18636, "p75": 32614,
    },
    "udc-law-school": {
        "total_students": 247, "total_receiving": 172, "pct_receiving": 70.0,
        "lt_half_n": 91, "lt_half_pct": 37.0,
        "half_to_full_n": 75, "half_to_full_pct": 30.0,
        "full_n": 6, "full_pct": 2.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 5458, "p50": 8292, "p75": 12436,
    },
    "drake-university-law": {
        "total_students": 349, "total_receiving": 286, "pct_receiving": 82.0,
        "lt_half_n": 76, "lt_half_pct": 22.0,
        "half_to_full_n": 179, "half_to_full_pct": 51.0,
        "full_n": 22, "full_pct": 6.0,
        "gt_full_n": 9, "gt_full_pct": 3.0,
        "p25": 24000, "p50": 31000, "p75": 37080,
    },
    "drexel-university-thomas-r-kline-school-of-law": {
        "total_students": 462, "total_receiving": 349, "pct_receiving": 76.0,
        "lt_half_n": 168, "lt_half_pct": 36.0,
        "half_to_full_n": 164, "half_to_full_pct": 35.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 16, "gt_full_pct": 3.0,
        "p25": 13000, "p50": 28667, "p75": 45000,
    },
    "duke-law": {
        "total_students": 726, "total_receiving": 685, "pct_receiving": 94.0,
        "lt_half_n": 487, "lt_half_pct": 67.0,
        "half_to_full_n": 196, "half_to_full_pct": 27.0,
        "full_n": 2, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 27000, "p50": 35000, "p75": 40000,
    },
    "duquesne-university-law": {
        "total_students": 509, "total_receiving": 488, "pct_receiving": 96.0,
        "lt_half_n": 268, "lt_half_pct": 53.0,
        "half_to_full_n": 136, "half_to_full_pct": 27.0,
        "full_n": 81, "full_pct": 16.0,
        "gt_full_n": 3, "gt_full_pct": 1.0,
        "p25": 20000, "p50": 30000, "p75": 43000,
    },
    "elon-university-law": {
        "total_students": 479, "total_receiving": 430, "pct_receiving": 90.0,
        "lt_half_n": 194, "lt_half_pct": 41.0,
        "half_to_full_n": 202, "half_to_full_pct": 42.0,
        "full_n": 23, "full_pct": 5.0,
        "gt_full_n": 11, "gt_full_pct": 2.0,
        "p25": 11429, "p50": 23571, "p75": 32543,
    },
    "emory-law-school": {
        "total_students": 758, "total_receiving": 744, "pct_receiving": 98.0,
        "lt_half_n": 340, "lt_half_pct": 45.0,
        "half_to_full_n": 394, "half_to_full_pct": 52.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 10, "gt_full_pct": 1.0,
        "p25": 28000, "p50": 36000, "p75": 45000,
    },
    "faulkner-university-law": {
        "total_students": 342, "total_receiving": 246, "pct_receiving": 72.0,
        "lt_half_n": 144, "lt_half_pct": 42.0,
        "half_to_full_n": 102, "half_to_full_pct": 30.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 3990, "p50": 15960, "p75": 25935,
    },
    "florida-am-law": {
        "total_students": 336, "total_receiving": 150, "pct_receiving": 45.0,
        "lt_half_n": 78, "lt_half_pct": 23.0,
        "half_to_full_n": 52, "half_to_full_pct": 15.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 20, "gt_full_pct": 6.0,
        "p25": 5000, "p50": 7500, "p75": 10000,
    },
    "florida-international-university-college-of-law": {
        "total_students": 555, "total_receiving": 306, "pct_receiving": 55.0,
        "lt_half_n": 158, "lt_half_pct": 28.0,
        "half_to_full_n": 63, "half_to_full_pct": 11.0,
        "full_n": 52, "full_pct": 9.0,
        "gt_full_n": 33, "gt_full_pct": 6.0,
        "p25": 7000, "p50": 10365, "p75": 20716,
    },
    "florida-state-university-college-of-law": {
        "total_students": 421, "total_receiving": 328, "pct_receiving": 78.0,
        "lt_half_n": 107, "lt_half_pct": 25.0,
        "half_to_full_n": 105, "half_to_full_pct": 25.0,
        "full_n": 50, "full_pct": 12.0,
        "gt_full_n": 66, "gt_full_pct": 16.0,
        "p25": 5737, "p50": 15464, "p75": 20640,
    },
    "university-of-florida-levin-college-of-law": {
        "total_students": 636, "total_receiving": 558, "pct_receiving": 88.0,
        "lt_half_n": 215, "lt_half_pct": 34.0,
        "half_to_full_n": 119, "half_to_full_pct": 19.0,
        "full_n": 85, "full_pct": 13.0,
        "gt_full_n": 139, "gt_full_pct": 22.0,
        "p25": 6000, "p50": 19000, "p75": 30804,
    },
    "fordham-law-school": {
        "total_students": 1335, "total_receiving": 1161, "pct_receiving": 87.0,
        "lt_half_n": 693, "lt_half_pct": 52.0,
        "half_to_full_n": 468, "half_to_full_pct": 35.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 25000, "p50": 35000, "p75": 43000,
    },
    "antonin-scalia-law-school": {
        "total_students": 437, "total_receiving": 365, "pct_receiving": 84.0,
        "lt_half_n": 138, "lt_half_pct": 32.0,
        "half_to_full_n": 201, "half_to_full_pct": 46.0,
        "full_n": 23, "full_pct": 5.0,
        "gt_full_n": 3, "gt_full_pct": 1.0,
        "p25": 14000, "p50": 21000, "p75": 32000,
    },
    "gw-law-school": {
        "total_students": 1679, "total_receiving": 1296, "pct_receiving": 77.0,
        "lt_half_n": 879, "lt_half_pct": 52.0,
        "half_to_full_n": 389, "half_to_full_pct": 23.0,
        "full_n": 27, "full_pct": 2.0,
        "gt_full_n": 1, "gt_full_pct": None,
        "p25": 20000, "p50": 28000, "p75": 40000,
    },
    "georgetown-law": {
        "total_students": 2228, "total_receiving": 1348, "pct_receiving": 61.0,
        "lt_half_n": 774, "lt_half_pct": 35.0,
        "half_to_full_n": 515, "half_to_full_pct": 23.0,
        "full_n": 43, "full_pct": 2.0,
        "gt_full_n": 16, "gt_full_pct": 1.0,
        "p25": 25000, "p50": 35000, "p75": 47500,
    },
    "georgia-state-university-college-of-law": {
        "total_students": 647, "total_receiving": 315, "pct_receiving": 49.0,
        "lt_half_n": 123, "lt_half_pct": 19.0,
        "half_to_full_n": 110, "half_to_full_pct": 17.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 82, "gt_full_pct": 13.0,
        "p25": 2336, "p50": 7776, "p75": 10627,
    },
    "uga-law-school": {
        "total_students": 527, "total_receiving": 443, "pct_receiving": 84.0,
        "lt_half_n": 152, "lt_half_pct": 29.0,
        "half_to_full_n": 68, "half_to_full_pct": 13.0,
        "full_n": 107, "full_pct": 20.0,
        "gt_full_n": 116, "gt_full_pct": 22.0,
        "p25": 5200, "p50": 19460, "p75": 22059,
    },
    "gonzaga-university-law": {
        "total_students": 573, "total_receiving": 561, "pct_receiving": 98.0,
        "lt_half_n": 181, "lt_half_pct": 32.0,
        "half_to_full_n": 368, "half_to_full_pct": 64.0,
        "full_n": 12, "full_pct": 2.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 27000, "p50": 32000, "p75": 34140,
    },
    "harvard-law-school": {
        "total_students": 1769, "total_receiving": 670, "pct_receiving": 38.0,
        "lt_half_n": 458, "lt_half_pct": 26.0,
        "half_to_full_n": 197, "half_to_full_pct": 11.0,
        "full_n": 8, "full_pct": None,
        "gt_full_n": 7, "gt_full_pct": None,
        "p25": 12106, "p50": 27510, "p75": 43338,
    },
    "university-of-hawaii-william-s-richardson-school-of-law": {
        "total_students": 306, "total_receiving": 211, "pct_receiving": 69.0,
        "lt_half_n": 166, "lt_half_pct": 54.0,
        "half_to_full_n": 42, "half_to_full_pct": 14.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 3, "gt_full_pct": 1.0,
        "p25": 6000, "p50": 7000, "p75": 13000,
    },
    "hofstra-university-school-of-law": {
        "total_students": 812, "total_receiving": 681, "pct_receiving": 84.0,
        "lt_half_n": 333, "lt_half_pct": 41.0,
        "half_to_full_n": 164, "half_to_full_pct": 20.0,
        "full_n": 184, "full_pct": 23.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 17500, "p50": 35793, "p75": 71586,
    },
    "university-of-houston-law-center": {
        "total_students": 811, "total_receiving": 600, "pct_receiving": 74.0,
        "lt_half_n": 488, "lt_half_pct": 60.0,
        "half_to_full_n": 88, "half_to_full_pct": 11.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 23, "gt_full_pct": 3.0,
        "p25": 4000, "p50": 11500, "p75": 16900,
    },
    "howard-university-school-of-law": {
        "total_students": 539, "total_receiving": 261, "pct_receiving": 48.0,
        "lt_half_n": 99, "lt_half_pct": 18.0,
        "half_to_full_n": 135, "half_to_full_pct": 25.0,
        "full_n": 25, "full_pct": 5.0,
        "gt_full_n": 2, "gt_full_pct": None,
        "p25": 10037, "p50": 23657, "p75": 32364,
    },
    "university-of-idaho-college-of-law": {
        "total_students": 502, "total_receiving": 231, "pct_receiving": 46.0,
        "lt_half_n": 214, "lt_half_pct": 43.0,
        "half_to_full_n": 17, "half_to_full_pct": 3.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 5000, "p50": 8000, "p75": 11000,
    },
    "university-of-illinois-chicago-school-of-law": {
        "total_students": 472, "total_receiving": 466, "pct_receiving": 99.0,
        "lt_half_n": 183, "lt_half_pct": 39.0,
        "half_to_full_n": 273, "half_to_full_pct": 58.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 10, "gt_full_pct": 2.0,
        "p25": 13125, "p50": 32000, "p75": 46500,
    },
    "uic-john-marshall-law-school": {
        "total_students": 988, "total_receiving": 599, "pct_receiving": 61.0,
        "lt_half_n": 267, "lt_half_pct": 27.0,
        "half_to_full_n": 323, "half_to_full_pct": 33.0,
        "full_n": 6, "full_pct": 1.0,
        "gt_full_n": 3, "gt_full_pct": None,
        "p25": 9000, "p50": 23250, "p75": 33750,
    },
    "indiana-university-maurer-school-of-law": {
        "total_students": 476, "total_receiving": 474, "pct_receiving": 100.0,
        "lt_half_n": 143, "lt_half_pct": 30.0,
        "half_to_full_n": 293, "half_to_full_pct": 62.0,
        "full_n": 2, "full_pct": None,
        "gt_full_n": 36, "gt_full_pct": 8.0,
        "p25": 20000, "p50": 33000, "p75": 50000,
    },
    "indiana-university-indianapolis-law": {
        "total_students": 773, "total_receiving": 779, "pct_receiving": 101.0,
        "lt_half_n": 533, "lt_half_pct": 69.0,
        "half_to_full_n": 196, "half_to_full_pct": 25.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 50, "gt_full_pct": 6.0,
        "p25": 5000, "p50": 15000, "p75": 29232,
    },
    "inter-american-university-pr-law": {
        "total_students": 659, "total_receiving": 52, "pct_receiving": 8.0,
        "lt_half_n": 39, "lt_half_pct": 6.0,
        "half_to_full_n": 6, "half_to_full_pct": 1.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 6, "gt_full_pct": 1.0,
        "p25": 1189, "p50": 2379, "p75": 3568,
    },
    "iowa-law": {
        "total_students": 500, "total_receiving": 383, "pct_receiving": 77.0,
        "lt_half_n": 86, "lt_half_pct": 17.0,
        "half_to_full_n": 155, "half_to_full_pct": 31.0,
        "full_n": 109, "full_pct": 22.0,
        "gt_full_n": 33, "gt_full_pct": 7.0,
        "p25": 24000, "p50": 30043, "p75": 34527,
    },
    "jacksonville-university-law": {
        "total_students": 79, "total_receiving": 79, "pct_receiving": 100.0,
        "lt_half_n": 45, "lt_half_pct": 57.0,
        "half_to_full_n": 34, "half_to_full_pct": 43.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 15000, "p50": 18000, "p75": 22000,
    },
    "kansas-law": {
        "total_students": 365, "total_receiving": 310, "pct_receiving": 85.0,
        "lt_half_n": 172, "lt_half_pct": 47.0,
        "half_to_full_n": 63, "half_to_full_pct": 17.0,
        "full_n": 50, "full_pct": 14.0,
        "gt_full_n": 25, "gt_full_pct": 7.0,
        "p25": 5375, "p50": 15000, "p75": 24700,
    },
    "university-of-kentucky-college-of-law": {
        "total_students": 395, "total_receiving": 386, "pct_receiving": 98.0,
        "lt_half_n": 303, "lt_half_pct": 77.0,
        "half_to_full_n": 66, "half_to_full_pct": 17.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 16, "gt_full_pct": 4.0,
        "p25": 6000, "p50": 11000, "p75": 18000,
    },
    "lewis-and-clark-law-school": {
        "total_students": 508, "total_receiving": 484, "pct_receiving": 95.0,
        "lt_half_n": 272, "lt_half_pct": 54.0,
        "half_to_full_n": 210, "half_to_full_pct": 41.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 2, "gt_full_pct": None,
        "p25": 19000, "p50": 26000, "p75": 36000,
    },
    "liberty-university-law": {
        "total_students": 302, "total_receiving": 302, "pct_receiving": 100.0,
        "lt_half_n": 126, "lt_half_pct": 42.0,
        "half_to_full_n": 176, "half_to_full_pct": 58.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 16000, "p50": 24000, "p75": 32000,
    },
    "lincoln-memorial-university-law": {
        "total_students": 400, "total_receiving": 398, "pct_receiving": 100.0,
        "lt_half_n": 156, "lt_half_pct": 39.0,
        "half_to_full_n": 214, "half_to_full_pct": 54.0,
        "full_n": 3, "full_pct": 1.0,
        "gt_full_n": 25, "gt_full_pct": 6.0,
        "p25": 15988, "p50": 28000, "p75": 32736,
    },
    "louisiana-state-university-law": {
        "total_students": 627, "total_receiving": 492, "pct_receiving": 78.0,
        "lt_half_n": 139, "lt_half_pct": 22.0,
        "half_to_full_n": 181, "half_to_full_pct": 29.0,
        "full_n": 126, "full_pct": 20.0,
        "gt_full_n": 46, "gt_full_pct": 7.0,
        "p25": 9937, "p50": 19750, "p75": 20600,
    },
    "university-of-louisville-louis-d-brandeis-school-of-law": {
        "total_students": 355, "total_receiving": 265, "pct_receiving": 75.0,
        "lt_half_n": 146, "lt_half_pct": 41.0,
        "half_to_full_n": 118, "half_to_full_pct": 33.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 1, "gt_full_pct": None,
        "p25": 5000, "p50": 12000, "p75": 20000,
    },
    "loyola-law-school-los-angeles": {
        "total_students": 980, "total_receiving": 605, "pct_receiving": 62.0,
        "lt_half_n": 140, "lt_half_pct": 14.0,
        "half_to_full_n": 451, "half_to_full_pct": 46.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 14, "gt_full_pct": 1.0,
        "p25": 35000, "p50": 45000, "p75": 52000,
    },
    "loyola-university-chicago-school-of-law": {
        "total_students": 832, "total_receiving": 781, "pct_receiving": 94.0,
        "lt_half_n": 290, "lt_half_pct": 35.0,
        "half_to_full_n": 471, "half_to_full_pct": 57.0,
        "full_n": 18, "full_pct": 2.0,
        "gt_full_n": 2, "gt_full_pct": None,
        "p25": 22000, "p50": 30000, "p75": 35000,
    },
    "loyola-university-new-orleans-college-of-law": {
        "total_students": 620, "total_receiving": 468, "pct_receiving": 75.0,
        "lt_half_n": 252, "lt_half_pct": 41.0,
        "half_to_full_n": 223, "half_to_full_pct": 36.0,
        "full_n": 5, "full_pct": 1.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 15000, "p50": 26000, "p75": 36000,
    },
    "university-of-maine-school-of-law": {
        "total_students": 275, "total_receiving": 198, "pct_receiving": 72.0,
        "lt_half_n": 49, "lt_half_pct": 18.0,
        "half_to_full_n": 67, "half_to_full_pct": 24.0,
        "full_n": 77, "full_pct": 28.0,
        "gt_full_n": 5, "gt_full_pct": 2.0,
        "p25": 15000, "p50": 23535, "p75": 32926,
    },
    "marquette-university-law-school": {
        "total_students": 574, "total_receiving": 541, "pct_receiving": 94.0,
        "lt_half_n": 322, "lt_half_pct": 56.0,
        "half_to_full_n": 218, "half_to_full_pct": 38.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 16000, "p50": 23000, "p75": 30000,
    },
    "university-of-maryland-carey-school-of-law": {
        "total_students": 644, "total_receiving": 507, "pct_receiving": 79.0,
        "lt_half_n": 210, "lt_half_pct": 33.0,
        "half_to_full_n": 197, "half_to_full_pct": 31.0,
        "full_n": 74, "full_pct": 11.0,
        "gt_full_n": 26, "gt_full_pct": 4.0,
        "p25": 10161, "p50": 28000, "p75": 40000,
    },
    "umass-dartmouth-law": {
        "total_students": 381, "total_receiving": 319, "pct_receiving": 84.0,
        "lt_half_n": 227, "lt_half_pct": 60.0,
        "half_to_full_n": 63, "half_to_full_pct": 17.0,
        "full_n": 29, "full_pct": 8.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 8000, "p50": 15000, "p75": 20787,
    },
    "university-of-memphis-school-of-law": {
        "total_students": 239, "total_receiving": 148, "pct_receiving": 62.0,
        "lt_half_n": 76, "lt_half_pct": 32.0,
        "half_to_full_n": 41, "half_to_full_pct": 17.0,
        "full_n": 13, "full_pct": 5.0,
        "gt_full_n": 18, "gt_full_pct": 8.0,
        "p25": 5000, "p50": 11000, "p75": 20000,
    },
    "mercer-university-school-of-law": {
        "total_students": 423, "total_receiving": 366, "pct_receiving": 87.0,
        "lt_half_n": 211, "lt_half_pct": 50.0,
        "half_to_full_n": 142, "half_to_full_pct": 34.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 12, "gt_full_pct": 3.0,
        "p25": 10000, "p50": 20000, "p75": 29000,
    },
    "university-of-miami-law": {
        "total_students": 1074, "total_receiving": 698, "pct_receiving": 65.0,
        "lt_half_n": 104, "lt_half_pct": 10.0,
        "half_to_full_n": 589, "half_to_full_pct": 55.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 5, "gt_full_pct": None,
        "p25": 39374, "p50": 47000, "p75": 52188,
    },
    "michigan-state-university-college-of-law": {
        "total_students": 553, "total_receiving": 553, "pct_receiving": 100.0,
        "lt_half_n": 80, "lt_half_pct": 14.0,
        "half_to_full_n": 188, "half_to_full_pct": 34.0,
        "full_n": 285, "full_pct": 52.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 35000, "p50": 42682, "p75": 47424,
    },
    "michigan-law-school": {
        "total_students": 990, "total_receiving": 902, "pct_receiving": 91.0,
        "lt_half_n": 574, "lt_half_pct": 58.0,
        "half_to_full_n": 297, "half_to_full_pct": 30.0,
        "full_n": 14, "full_pct": 1.0,
        "gt_full_n": 17, "gt_full_pct": 2.0,
        "p25": 20000, "p50": 32000, "p75": 41000,
    },
    "university-of-minnesota-law-school": {
        "total_students": 639, "total_receiving": 603, "pct_receiving": 94.0,
        "lt_half_n": 277, "lt_half_pct": 43.0,
        "half_to_full_n": 319, "half_to_full_pct": 50.0,
        "full_n": 7, "full_pct": 1.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 25000, "p50": 35000, "p75": 48528,
    },
    "mississippi-college-school-of-law": {
        "total_students": 355, "total_receiving": 229, "pct_receiving": 65.0,
        "lt_half_n": 101, "lt_half_pct": 28.0,
        "half_to_full_n": 92, "half_to_full_pct": 26.0,
        "full_n": 36, "full_pct": 10.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 14000, "p50": 22000, "p75": 32000,
    },
    "university-of-missouri-school-of-law": {
        "total_students": 371, "total_receiving": 292, "pct_receiving": 79.0,
        "lt_half_n": 116, "lt_half_pct": 31.0,
        "half_to_full_n": 116, "half_to_full_pct": 31.0,
        "full_n": 14, "full_pct": 4.0,
        "gt_full_n": 46, "gt_full_pct": 12.0,
        "p25": 8000, "p50": 17500, "p75": 24000,
    },
    "umkc-law-school": {
        "total_students": 435, "total_receiving": 371, "pct_receiving": 85.0,
        "lt_half_n": 302, "lt_half_pct": 69.0,
        "half_to_full_n": 51, "half_to_full_pct": 12.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 18, "gt_full_pct": 4.0,
        "p25": 4000, "p50": 7270, "p75": 11250,
    },
    "william-mitchell-college-of-law": {
        "total_students": 1128, "total_receiving": 1131, "pct_receiving": 100.0,
        "lt_half_n": 641, "lt_half_pct": 57.0,
        "half_to_full_n": 439, "half_to_full_pct": 39.0,
        "full_n": 51, "full_pct": 5.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 21372, "p50": 34730, "p75": 42744,
    },
    "university-of-montana-school-of-law": {
        "total_students": 269, "total_receiving": 240, "pct_receiving": 89.0,
        "lt_half_n": 192, "lt_half_pct": 71.0,
        "half_to_full_n": 23, "half_to_full_pct": 9.0,
        "full_n": 8, "full_pct": 3.0,
        "gt_full_n": 17, "gt_full_pct": 6.0,
        "p25": 6630, "p50": 10000, "p75": 15585,
    },
    "university-of-nebraska-college-of-law": {
        "total_students": 442, "total_receiving": 376, "pct_receiving": 85.0,
        "lt_half_n": 77, "lt_half_pct": 17.0,
        "half_to_full_n": 269, "half_to_full_pct": 61.0,
        "full_n": 30, "full_pct": 7.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 9510, "p50": 20000, "p75": 35100,
    },
    "university-of-nevada-las-vegas-school-of-law": {
        "total_students": 416, "total_receiving": 317, "pct_receiving": 76.0,
        "lt_half_n": 169, "lt_half_pct": 41.0,
        "half_to_full_n": 105, "half_to_full_pct": 25.0,
        "full_n": 23, "full_pct": 6.0,
        "gt_full_n": 20, "gt_full_pct": 5.0,
        "p25": 14980, "p50": 22998, "p75": 39946,
    },
    "new-england-law-boston": {
        "total_students": 1133, "total_receiving": 1001, "pct_receiving": 88.0,
        "lt_half_n": 365, "lt_half_pct": 32.0,
        "half_to_full_n": 277, "half_to_full_pct": 24.0,
        "full_n": 327, "full_pct": 29.0,
        "gt_full_n": 7, "gt_full_pct": 1.0,
        "p25": 19500, "p50": 40000, "p75": 59900,
    },
    "university-of-new-hampshire-school-of-law": {
        "total_students": 657, "total_receiving": 456, "pct_receiving": 69.0,
        "lt_half_n": 223, "lt_half_pct": 34.0,
        "half_to_full_n": 220, "half_to_full_pct": 33.0,
        "full_n": 2, "full_pct": None,
        "gt_full_n": 11, "gt_full_pct": 2.0,
        "p25": 7500, "p50": 20000, "p75": 34000,
    },
    "university-of-new-mexico-school-of-law": {
        "total_students": 290, "total_receiving": 175, "pct_receiving": 60.0,
        "lt_half_n": 113, "lt_half_pct": 39.0,
        "half_to_full_n": 38, "half_to_full_pct": 13.0,
        "full_n": 7, "full_pct": 2.0,
        "gt_full_n": 17, "gt_full_pct": 6.0,
        "p25": 5000, "p50": 10000, "p75": 16000,
    },
    "new-york-law-school": {
        "total_students": 1057, "total_receiving": 900, "pct_receiving": 85.0,
        "lt_half_n": 308, "lt_half_pct": 29.0,
        "half_to_full_n": 590, "half_to_full_pct": 56.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 2, "gt_full_pct": None,
        "p25": 21300, "p50": 40000, "p75": 55000,
    },
    "nyu-law-school": {
        "total_students": 1327, "total_receiving": 870, "pct_receiving": 66.0,
        "lt_half_n": 656, "lt_half_pct": 49.0,
        "half_to_full_n": 66, "half_to_full_pct": 5.0,
        "full_n": 148, "full_pct": 11.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 12000, "p50": 20000, "p75": 35750,
    },
    "north-carolina-central-law": {
        "total_students": 391, "total_receiving": 281, "pct_receiving": 72.0,
        "lt_half_n": 129, "lt_half_pct": 33.0,
        "half_to_full_n": 96, "half_to_full_pct": 25.0,
        "full_n": 3, "full_pct": 1.0,
        "gt_full_n": 53, "gt_full_pct": 14.0,
        "p25": 8080, "p50": 12982, "p75": 20907,
    },
    "unc-law-school": {
        "total_students": 544, "total_receiving": 516, "pct_receiving": 95.0,
        "lt_half_n": 209, "lt_half_pct": 38.0,
        "half_to_full_n": 281, "half_to_full_pct": 52.0,
        "full_n": 15, "full_pct": 3.0,
        "gt_full_n": 11, "gt_full_pct": 2.0,
        "p25": 15000, "p50": 17000, "p75": 20000,
    },
    "university-of-north-dakota-school-of-law": {
        "total_students": 227, "total_receiving": 119, "pct_receiving": 52.0,
        "lt_half_n": 86, "lt_half_pct": 38.0,
        "half_to_full_n": 29, "half_to_full_pct": 13.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 4, "gt_full_pct": 2.0,
        "p25": 1775, "p50": 3875, "p75": 10500,
    },
    "unt-dallas-law": {
        "total_students": 435, "total_receiving": 237, "pct_receiving": 54.0,
        "lt_half_n": 193, "lt_half_pct": 44.0,
        "half_to_full_n": 36, "half_to_full_pct": 8.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 8, "gt_full_pct": 2.0,
        "p25": 2000, "p50": 5000, "p75": 7000,
    },
    "northeastern-law-school": {
        "total_students": 729, "total_receiving": 690, "pct_receiving": 95.0,
        "lt_half_n": 360, "lt_half_pct": 49.0,
        "half_to_full_n": 251, "half_to_full_pct": 34.0,
        "full_n": 70, "full_pct": 10.0,
        "gt_full_n": 9, "gt_full_pct": 1.0,
        "p25": 14666, "p50": 31000, "p75": 45000,
    },
    "northern-illinois-university-law": {
        "total_students": 338, "total_receiving": 263, "pct_receiving": 78.0,
        "lt_half_n": 177, "lt_half_pct": 52.0,
        "half_to_full_n": 86, "half_to_full_pct": 25.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 4137, "p50": 8274, "p75": 12411,
    },
    "northern-kentucky-university-law": {
        "total_students": 391, "total_receiving": 228, "pct_receiving": 58.0,
        "lt_half_n": 99, "lt_half_pct": 25.0,
        "half_to_full_n": 125, "half_to_full_pct": 32.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 4, "gt_full_pct": 1.0,
        "p25": 10000, "p50": 18000, "p75": 25000,
    },
    "northwestern-law-school": {
        "total_students": 760, "total_receiving": 581, "pct_receiving": 76.0,
        "lt_half_n": 262, "lt_half_pct": 34.0,
        "half_to_full_n": 301, "half_to_full_pct": 40.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 18, "gt_full_pct": 2.0,
        "p25": 20000, "p50": 40000, "p75": 50000,
    },
    "notre-dame-law-school": {
        "total_students": 553, "total_receiving": 507, "pct_receiving": 92.0,
        "lt_half_n": 299, "lt_half_pct": 54.0,
        "half_to_full_n": 173, "half_to_full_pct": 31.0,
        "full_n": 24, "full_pct": 4.0,
        "gt_full_n": 11, "gt_full_pct": 2.0,
        "p25": 20000, "p50": 30000, "p75": 50000,
    },
    "nova-southeastern-law": {
        "total_students": 620, "total_receiving": 484, "pct_receiving": 78.0,
        "lt_half_n": 372, "lt_half_pct": 60.0,
        "half_to_full_n": 112, "half_to_full_pct": 18.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 10750, "p50": 15000, "p75": 21310,
    },
    "ohio-northern-university-law": {
        "total_students": 314, "total_receiving": 0, "pct_receiving": None,
        "lt_half_n": 137, "lt_half_pct": 44.0,
        "half_to_full_n": 85, "half_to_full_pct": 27.0,
        "full_n": 11, "full_pct": 4.0,
        "gt_full_n": 11, "gt_full_pct": 4.0,
        "p25": 9125, "p50": 17250, "p75": 25700,
    },
    "oklahoma-city-university-law": {
        "total_students": 427, "total_receiving": 330, "pct_receiving": 77.0,
        "lt_half_n": 223, "lt_half_pct": 52.0,
        "half_to_full_n": 76, "half_to_full_pct": 18.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 31, "gt_full_pct": 7.0,
        "p25": 10000, "p50": 14000, "p75": 20000,
    },
    "oregon-law": {
        "total_students": 457, "total_receiving": 338, "pct_receiving": 74.0,
        "lt_half_n": 169, "lt_half_pct": 37.0,
        "half_to_full_n": 149, "half_to_full_pct": 33.0,
        "full_n": 12, "full_pct": 3.0,
        "gt_full_n": 8, "gt_full_pct": 2.0,
        "p25": 15000, "p50": 25000, "p75": 35000,
    },
    "pace-university-school-of-law": {
        "total_students": 865, "total_receiving": 808, "pct_receiving": 93.0,
        "lt_half_n": 406, "lt_half_pct": 47.0,
        "half_to_full_n": 359, "half_to_full_pct": 42.0,
        "full_n": 42, "full_pct": 5.0,
        "gt_full_n": 1, "gt_full_pct": None,
        "p25": 20000, "p50": 29000, "p75": 38500,
    },
    "university-of-pacific-law": {
        "total_students": 577, "total_receiving": 536, "pct_receiving": 93.0,
        "lt_half_n": 269, "lt_half_pct": 47.0,
        "half_to_full_n": 124, "half_to_full_pct": 21.0,
        "full_n": 143, "full_pct": 25.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 3388, "p50": 30426, "p75": 60852,
    },
    "penn-state-dickinson-law": {
        "total_students": 689, "total_receiving": 692, "pct_receiving": 100.0,
        "lt_half_n": 152, "lt_half_pct": 22.0,
        "half_to_full_n": 494, "half_to_full_pct": 72.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 46, "gt_full_pct": 7.0,
        "p25": 26000, "p50": 36000, "p75": 56928,
    },
    "penn-law": {
        "total_students": 759, "total_receiving": 450, "pct_receiving": 59.0,
        "lt_half_n": 188, "lt_half_pct": 25.0,
        "half_to_full_n": 157, "half_to_full_pct": 21.0,
        "full_n": 96, "full_pct": 13.0,
        "gt_full_n": 9, "gt_full_pct": 1.0,
        "p25": 29512, "p50": 42246, "p75": 60000,
    },
    "pepperdine-university-school-of-law": {
        "total_students": 532, "total_receiving": 410, "pct_receiving": 77.0,
        "lt_half_n": 252, "lt_half_pct": 47.0,
        "half_to_full_n": 133, "half_to_full_pct": 25.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 24, "gt_full_pct": 5.0,
        "p25": 30000, "p50": 30000, "p75": 55000,
    },
    "pittsburgh-law-school": {
        "total_students": 373, "total_receiving": 305, "pct_receiving": 82.0,
        "lt_half_n": 77, "lt_half_pct": 21.0,
        "half_to_full_n": 214, "half_to_full_pct": 57.0,
        "full_n": 5, "full_pct": 1.0,
        "gt_full_n": 9, "gt_full_pct": 2.0,
        "p25": 24000, "p50": 30000, "p75": 38000,
    },
    "university-of-puerto-rico-law": {
        "total_students": 449, "total_receiving": 226, "pct_receiving": 50.0,
        "lt_half_n": 209, "lt_half_pct": 47.0,
        "half_to_full_n": 17, "half_to_full_pct": 4.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 1000, "p50": 1000, "p75": 1000,
    },
    "quinnipiac-university-school-of-law": {
        "total_students": 389, "total_receiving": 332, "pct_receiving": 85.0,
        "lt_half_n": 190, "lt_half_pct": 49.0,
        "half_to_full_n": 124, "half_to_full_pct": 32.0,
        "full_n": 11, "full_pct": 3.0,
        "gt_full_n": 7, "gt_full_pct": 2.0,
        "p25": 14000, "p50": 25000, "p75": 35000,
    },
    "regent-university-law": {
        "total_students": 326, "total_receiving": 316, "pct_receiving": 97.0,
        "lt_half_n": 148, "lt_half_pct": 45.0,
        "half_to_full_n": 166, "half_to_full_pct": 51.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 2, "gt_full_pct": 1.0,
        "p25": 10000, "p50": 22500, "p75": 35000,
    },
    "university-of-richmond-school-of-law": {
        "total_students": 392, "total_receiving": 319, "pct_receiving": 81.0,
        "lt_half_n": 110, "lt_half_pct": 28.0,
        "half_to_full_n": 101, "half_to_full_pct": 26.0,
        "full_n": 101, "full_pct": 26.0,
        "gt_full_n": 7, "gt_full_pct": 2.0,
        "p25": 20000, "p50": 40000, "p75": 55550,
    },
    "roger-williams-university-school-of-law": {
        "total_students": 481, "total_receiving": 311, "pct_receiving": 65.0,
        "lt_half_n": 115, "lt_half_pct": 24.0,
        "half_to_full_n": 181, "half_to_full_pct": 38.0,
        "full_n": 13, "full_pct": 3.0,
        "gt_full_n": 2, "gt_full_pct": None,
        "p25": 18562, "p50": 28000, "p75": 34804,
    },
    "rutgers-law-school": {
        "total_students": 1235, "total_receiving": 1068, "pct_receiving": 86.0,
        "lt_half_n": 670, "lt_half_pct": 54.0,
        "half_to_full_n": 386, "half_to_full_pct": 31.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 11, "gt_full_pct": 1.0,
        "p25": 10000, "p50": 16000, "p75": 22000,
    },
    "saint-louis-university-law": {
        "total_students": 552, "total_receiving": 538, "pct_receiving": 97.0,
        "lt_half_n": 316, "lt_half_pct": 57.0,
        "half_to_full_n": 216, "half_to_full_pct": 39.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 6, "gt_full_pct": 1.0,
        "p25": 15000, "p50": 25000, "p75": 35000,
    },
    "samford-university-cumberland-school-of-law": {
        "total_students": 446, "total_receiving": 415, "pct_receiving": 93.0,
        "lt_half_n": 220, "lt_half_pct": 49.0,
        "half_to_full_n": 192, "half_to_full_pct": 43.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 2, "gt_full_pct": None,
        "p25": 7500, "p50": 20000, "p75": 27000,
    },
    "usd-law-school": {
        "total_students": 751, "total_receiving": 684, "pct_receiving": 91.0,
        "lt_half_n": 355, "lt_half_pct": 47.0,
        "half_to_full_n": 305, "half_to_full_pct": 41.0,
        "full_n": 18, "full_pct": 2.0,
        "gt_full_n": 6, "gt_full_pct": 1.0,
        "p25": 27375, "p50": 32000, "p75": 50000,
    },
    "university-of-san-francisco-law": {
        "total_students": 462, "total_receiving": 285, "pct_receiving": 62.0,
        "lt_half_n": 116, "lt_half_pct": 25.0,
        "half_to_full_n": 166, "half_to_full_pct": 36.0,
        "full_n": 2, "full_pct": None,
        "gt_full_n": 1, "gt_full_pct": None,
        "p25": 24343, "p50": 33718, "p75": 43928,
    },
    "santa-clara-university-school-of-law": {
        "total_students": 694, "total_receiving": 426, "pct_receiving": 61.0,
        "lt_half_n": 121, "lt_half_pct": 17.0,
        "half_to_full_n": 167, "half_to_full_pct": 24.0,
        "full_n": 128, "full_pct": 18.0,
        "gt_full_n": 10, "gt_full_pct": 1.0,
        "p25": 30000, "p50": 40000, "p75": 57262,
    },
    "seattle-university-school-of-law": {
        "total_students": 666, "total_receiving": 581, "pct_receiving": 87.0,
        "lt_half_n": 353, "lt_half_pct": 53.0,
        "half_to_full_n": 154, "half_to_full_pct": 23.0,
        "full_n": 61, "full_pct": 9.0,
        "gt_full_n": 13, "gt_full_pct": 2.0,
        "p25": 20000, "p50": 30000, "p75": 40000,
    },
    "seton-hall-university-school-of-law": {
        "total_students": 761, "total_receiving": 673, "pct_receiving": 88.0,
        "lt_half_n": 283, "lt_half_pct": 37.0,
        "half_to_full_n": 383, "half_to_full_pct": 50.0,
        "full_n": 7, "full_pct": 1.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 20500, "p50": 40000, "p75": 60000,
    },
    "university-of-south-carolina-school-of-law": {
        "total_students": 628, "total_receiving": 565, "pct_receiving": 90.0,
        "lt_half_n": 381, "lt_half_pct": 61.0,
        "half_to_full_n": 160, "half_to_full_pct": 25.0,
        "full_n": 7, "full_pct": 1.0,
        "gt_full_n": 17, "gt_full_pct": 3.0,
        "p25": 8000, "p50": 12562, "p75": 18000,
    },
    "university-of-south-dakota-school-of-law": {
        "total_students": 253, "total_receiving": 139, "pct_receiving": 55.0,
        "lt_half_n": 90, "lt_half_pct": 36.0,
        "half_to_full_n": 31, "half_to_full_pct": 12.0,
        "full_n": 13, "full_pct": 5.0,
        "gt_full_n": 5, "gt_full_pct": 2.0,
        "p25": 590, "p50": 5000, "p75": 10000,
    },
    "south-texas-college-of-law-houston": {
        "total_students": 1162, "total_receiving": 985, "pct_receiving": 85.0,
        "lt_half_n": 883, "lt_half_pct": 76.0,
        "half_to_full_n": 97, "half_to_full_pct": 8.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 5, "gt_full_pct": None,
        "p25": 3399, "p50": 11000, "p75": 18000,
    },
    "usc-gould-school-of-law": {
        "total_students": 632, "total_receiving": 601, "pct_receiving": 95.0,
        "lt_half_n": 311, "lt_half_pct": 49.0,
        "half_to_full_n": 282, "half_to_full_pct": 45.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 8, "gt_full_pct": 1.0,
        "p25": 38220, "p50": 40092, "p75": 48160,
    },
    "southern-illinois-university-law": {
        "total_students": 269, "total_receiving": 172, "pct_receiving": 64.0,
        "lt_half_n": 87, "lt_half_pct": 32.0,
        "half_to_full_n": 78, "half_to_full_pct": 29.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 6, "gt_full_pct": 2.0,
        "p25": 6000, "p50": 9500, "p75": 16540,
    },
    "smu-dedman-school-of-law": {
        "total_students": 649, "total_receiving": 611, "pct_receiving": 94.0,
        "lt_half_n": 415, "lt_half_pct": 64.0,
        "half_to_full_n": 173, "half_to_full_pct": 27.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 23, "gt_full_pct": 4.0,
        "p25": 22000, "p50": 26000, "p75": 33333,
    },
    "southern-university-law": {
        "total_students": 800, "total_receiving": 163, "pct_receiving": 20.0,
        "lt_half_n": 101, "lt_half_pct": 13.0,
        "half_to_full_n": 59, "half_to_full_pct": 7.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 3, "gt_full_pct": None,
        "p25": 8000, "p50": 11000, "p75": 13000,
    },
    "southwestern-law-school": {
        "total_students": 1006, "total_receiving": 722, "pct_receiving": 72.0,
        "lt_half_n": 373, "lt_half_pct": 37.0,
        "half_to_full_n": 320, "half_to_full_pct": 32.0,
        "full_n": 12, "full_pct": 1.0,
        "gt_full_n": 17, "gt_full_pct": 2.0,
        "p25": 12500, "p50": 29166, "p75": 42372,
    },
    "st-john-s-university-school-of-law": {
        "total_students": 765, "total_receiving": 643, "pct_receiving": 84.0,
        "lt_half_n": 260, "lt_half_pct": 34.0,
        "half_to_full_n": 383, "half_to_full_pct": 50.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 25000, "p50": 60000, "p75": 72630,
    },
    "st-mary-s-university-school-of-law": {
        "total_students": 833, "total_receiving": 677, "pct_receiving": 81.0,
        "lt_half_n": 603, "lt_half_pct": 72.0,
        "half_to_full_n": 72, "half_to_full_pct": 9.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 2, "gt_full_pct": None,
        "p25": 8500, "p50": 12500, "p75": 17605,
    },
    "st-thomas-university-miami-law": {
        "total_students": 821, "total_receiving": 526, "pct_receiving": 64.0,
        "lt_half_n": 237, "lt_half_pct": 29.0,
        "half_to_full_n": 287, "half_to_full_pct": 35.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 2, "gt_full_pct": None,
        "p25": 12500, "p50": 25320, "p75": 34750,
    },
    "st-thomas-minneapolis-law": {
        "total_students": 475, "total_receiving": 474, "pct_receiving": 100.0,
        "lt_half_n": 192, "lt_half_pct": 40.0,
        "half_to_full_n": 255, "half_to_full_pct": 54.0,
        "full_n": 27, "full_pct": 6.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 19678, "p50": 26865, "p75": 34490,
    },
    "stanford-law-school": {
        "total_students": 565, "total_receiving": 285, "pct_receiving": 50.0,
        "lt_half_n": 77, "lt_half_pct": 14.0,
        "half_to_full_n": 142, "half_to_full_pct": 25.0,
        "full_n": 2, "full_pct": None,
        "gt_full_n": 64, "gt_full_pct": 11.0,
        "p25": 34240, "p50": 52797, "p75": 72247,
    },
    "stetson-university-law": {
        "total_students": 971, "total_receiving": 855, "pct_receiving": 88.0,
        "lt_half_n": 489, "lt_half_pct": 50.0,
        "half_to_full_n": 358, "half_to_full_pct": 37.0,
        "full_n": 7, "full_pct": 1.0,
        "gt_full_n": 1, "gt_full_pct": None,
        "p25": 15000, "p50": 25000, "p75": 35000,
    },
    "suffolk-university-law-school": {
        "total_students": 1194, "total_receiving": 962, "pct_receiving": 81.0,
        "lt_half_n": 646, "lt_half_pct": 54.0,
        "half_to_full_n": 288, "half_to_full_pct": 24.0,
        "full_n": 24, "full_pct": 2.0,
        "gt_full_n": 4, "gt_full_pct": None,
        "p25": 10000, "p50": 22445, "p75": 40000,
    },
    "syracuse-university-law": {
        "total_students": 720, "total_receiving": 658, "pct_receiving": 91.0,
        "lt_half_n": 400, "lt_half_pct": 56.0,
        "half_to_full_n": 247, "half_to_full_pct": 34.0,
        "full_n": 3, "full_pct": None,
        "gt_full_n": 8, "gt_full_pct": 1.0,
        "p25": 14000, "p50": 24351, "p75": 40211,
    },
    "temple-university-beasley-school-of-law": {
        "total_students": 677, "total_receiving": 578, "pct_receiving": 85.0,
        "lt_half_n": 131, "lt_half_pct": 19.0,
        "half_to_full_n": 422, "half_to_full_pct": 62.0,
        "full_n": 4, "full_pct": 1.0,
        "gt_full_n": 21, "gt_full_pct": 3.0,
        "p25": 22271, "p50": 31334, "p75": 41500,
    },
    "university-of-tennessee-college-of-law": {
        "total_students": 401, "total_receiving": 358, "pct_receiving": 89.0,
        "lt_half_n": 267, "lt_half_pct": 67.0,
        "half_to_full_n": 74, "half_to_full_pct": 18.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 17, "gt_full_pct": 4.0,
        "p25": 8000, "p50": 10550, "p75": 14000,
    },
    "texas-am-law-school": {
        "total_students": 394, "total_receiving": 375, "pct_receiving": 95.0,
        "lt_half_n": 88, "lt_half_pct": 22.0,
        "half_to_full_n": 190, "half_to_full_pct": 48.0,
        "full_n": 27, "full_pct": 7.0,
        "gt_full_n": 70, "gt_full_pct": 18.0,
        "p25": 17500, "p50": 26048, "p75": 35534,
    },
    "texas-southern-university-law": {
        "total_students": 552, "total_receiving": 283, "pct_receiving": 51.0,
        "lt_half_n": 182, "lt_half_pct": 33.0,
        "half_to_full_n": 94, "half_to_full_pct": 17.0,
        "full_n": 7, "full_pct": 1.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 3500, "p50": 7363, "p75": 10519,
    },
    "texas-tech-university-school-of-law": {
        "total_students": 441, "total_receiving": 386, "pct_receiving": 88.0,
        "lt_half_n": 242, "lt_half_pct": 55.0,
        "half_to_full_n": 124, "half_to_full_pct": 28.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 20, "gt_full_pct": 5.0,
        "p25": 5869, "p50": 12000, "p75": 17850,
    },
    "ut-austin-law-school": {
        "total_students": 865, "total_receiving": 815, "pct_receiving": 94.0,
        "lt_half_n": 347, "lt_half_pct": 40.0,
        "half_to_full_n": 383, "half_to_full_pct": 44.0,
        "full_n": 23, "full_pct": 3.0,
        "gt_full_n": 62, "gt_full_pct": 7.0,
        "p25": 10000, "p50": 23375, "p75": 38586,
    },
    "moritz-college-of-law": {
        "total_students": 482, "total_receiving": 433, "pct_receiving": 90.0,
        "lt_half_n": 78, "lt_half_pct": 16.0,
        "half_to_full_n": 180, "half_to_full_pct": 37.0,
        "full_n": 142, "full_pct": 29.0,
        "gt_full_n": 33, "gt_full_pct": 7.0,
        "p25": 17360, "p50": 29984, "p75": 37220,
    },
    "university-of-toledo-college-of-law": {
        "total_students": 352, "total_receiving": 288, "pct_receiving": 82.0,
        "lt_half_n": 126, "lt_half_pct": 36.0,
        "half_to_full_n": 68, "half_to_full_pct": 19.0,
        "full_n": 51, "full_pct": 14.0,
        "gt_full_n": 43, "gt_full_pct": 12.0,
        "p25": 8300, "p50": 20000, "p75": 25676,
    },
    "touro-law-center": {
        "total_students": 577, "total_receiving": 441, "pct_receiving": 76.0,
        "lt_half_n": 244, "lt_half_pct": 42.0,
        "half_to_full_n": 192, "half_to_full_pct": 33.0,
        "full_n": 5, "full_pct": 1.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 10000, "p50": 22000, "p75": 35000,
    },
    "tulane-university-law-school": {
        "total_students": 614, "total_receiving": 532, "pct_receiving": 87.0,
        "lt_half_n": 251, "lt_half_pct": 41.0,
        "half_to_full_n": 278, "half_to_full_pct": 45.0,
        "full_n": 3, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 30000, "p50": 37500, "p75": 45000,
    },
    "tulsa-university-law": {
        "total_students": 282, "total_receiving": 279, "pct_receiving": 99.0,
        "lt_half_n": 148, "lt_half_pct": 52.0,
        "half_to_full_n": 119, "half_to_full_pct": 42.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 12, "gt_full_pct": 4.0,
        "p25": 5950, "p50": 12500, "p75": 26633,
    },
    "university-of-utah-sjd": {
        "total_students": 307, "total_receiving": 280, "pct_receiving": 91.0,
        "lt_half_n": 84, "lt_half_pct": 27.0,
        "half_to_full_n": 144, "half_to_full_pct": 47.0,
        "full_n": 24, "full_pct": 8.0,
        "gt_full_n": 28, "gt_full_pct": 9.0,
        "p25": 14205, "p50": 28410, "p75": 37996,
    },
    "vanderbilt-law-school": {
        "total_students": 505, "total_receiving": 453, "pct_receiving": 90.0,
        "lt_half_n": 310, "lt_half_pct": 61.0,
        "half_to_full_n": 120, "half_to_full_pct": 24.0,
        "full_n": 8, "full_pct": 2.0,
        "gt_full_n": 15, "gt_full_pct": 3.0,
        "p25": 17500, "p50": 32792, "p75": 44927,
    },
    "vermont-law-school": {
        "total_students": 557, "total_receiving": 550, "pct_receiving": 99.0,
        "lt_half_n": 355, "lt_half_pct": 64.0,
        "half_to_full_n": 153, "half_to_full_pct": 27.0,
        "full_n": 42, "full_pct": 8.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 18000, "p50": 28000, "p75": 41250,
    },
    "villanova-university-charles-widger-school-of-law": {
        "total_students": 555, "total_receiving": 481, "pct_receiving": 87.0,
        "lt_half_n": 143, "lt_half_pct": 26.0,
        "half_to_full_n": 338, "half_to_full_pct": 61.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 25000, "p50": 40000, "p75": 50000,
    },
    "uva-law-school": {
        "total_students": 915, "total_receiving": 607, "pct_receiving": 66.0,
        "lt_half_n": 312, "lt_half_pct": 34.0,
        "half_to_full_n": 234, "half_to_full_pct": 26.0,
        "full_n": 56, "full_pct": 6.0,
        "gt_full_n": 5, "gt_full_pct": 1.0,
        "p25": 25000, "p50": 35000, "p75": 50000,
    },
    "wake-forest-law-school": {
        "total_students": 517, "total_receiving": 392, "pct_receiving": 76.0,
        "lt_half_n": 93, "lt_half_pct": 18.0,
        "half_to_full_n": 281, "half_to_full_pct": 54.0,
        "full_n": 3, "full_pct": 1.0,
        "gt_full_n": 15, "gt_full_pct": 3.0,
        "p25": 30000, "p50": 40000, "p75": 55162,
    },
    "washburn-university-school-of-law": {
        "total_students": 342, "total_receiving": 272, "pct_receiving": 80.0,
        "lt_half_n": 123, "lt_half_pct": 36.0,
        "half_to_full_n": 90, "half_to_full_pct": 26.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 59, "gt_full_pct": 17.0,
        "p25": 11915, "p50": 15000, "p75": 21500,
    },
    "washington-and-lee-law-school": {
        "total_students": 354, "total_receiving": 342, "pct_receiving": 97.0,
        "lt_half_n": 116, "lt_half_pct": 33.0,
        "half_to_full_n": 222, "half_to_full_pct": 63.0,
        "full_n": 3, "full_pct": 1.0,
        "gt_full_n": 1, "gt_full_pct": None,
        "p25": 19000, "p50": 33500, "p75": 41500,
    },
    "washington-university-school-of-law": {
        "total_students": 781, "total_receiving": 738, "pct_receiving": 94.0,
        "lt_half_n": 120, "lt_half_pct": 15.0,
        "half_to_full_n": 455, "half_to_full_pct": 58.0,
        "full_n": 86, "full_pct": 11.0,
        "gt_full_n": 77, "gt_full_pct": 10.0,
        "p25": 40000, "p50": 48000, "p75": 60000,
    },
    "university-of-washington-school-of-law": {
        "total_students": 544, "total_receiving": 383, "pct_receiving": 70.0,
        "lt_half_n": 335, "lt_half_pct": 62.0,
        "half_to_full_n": 48, "half_to_full_pct": 9.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 7500, "p50": 14000, "p75": 18930,
    },
    "wayne-state-university-law": {
        "total_students": 389, "total_receiving": 322, "pct_receiving": 83.0,
        "lt_half_n": 55, "lt_half_pct": 14.0,
        "half_to_full_n": 76, "half_to_full_pct": 20.0,
        "full_n": 159, "full_pct": 41.0,
        "gt_full_n": 32, "gt_full_pct": 8.0,
        "p25": 25000, "p50": 34636, "p75": 38547,
    },
    "west-virginia-university-college-of-law": {
        "total_students": 296, "total_receiving": 229, "pct_receiving": 77.0,
        "lt_half_n": 96, "lt_half_pct": 32.0,
        "half_to_full_n": 107, "half_to_full_pct": 36.0,
        "full_n": 17, "full_pct": 6.0,
        "gt_full_n": 9, "gt_full_pct": 3.0,
        "p25": 10848, "p50": 17500, "p75": 26110,
    },
    "western-new-england-university-school-of-law": {
        "total_students": 306, "total_receiving": 251, "pct_receiving": 82.0,
        "lt_half_n": 74, "lt_half_pct": 24.0,
        "half_to_full_n": 139, "half_to_full_pct": 45.0,
        "full_n": 37, "full_pct": 12.0,
        "gt_full_n": 1, "gt_full_pct": None,
        "p25": 23000, "p50": 30000, "p75": 35000,
    },
    "western-state-law": {
        "total_students": 365, "total_receiving": 283, "pct_receiving": 78.0,
        "lt_half_n": 157, "lt_half_pct": 43.0,
        "half_to_full_n": 126, "half_to_full_pct": 35.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 13120, "p50": 23616, "p75": 36000,
    },
    "widener-commonwealth-law": {
        "total_students": 401, "total_receiving": 379, "pct_receiving": 95.0,
        "lt_half_n": 228, "lt_half_pct": 57.0,
        "half_to_full_n": 149, "half_to_full_pct": 37.0,
        "full_n": 2, "full_pct": None,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 15625, "p50": 25000, "p75": 34000,
    },
    "widener-delaware-law": {
        "total_students": 684, "total_receiving": 679, "pct_receiving": 99.0,
        "lt_half_n": 310, "lt_half_pct": 45.0,
        "half_to_full_n": 360, "half_to_full_pct": 53.0,
        "full_n": 5, "full_pct": 1.0,
        "gt_full_n": 4, "gt_full_pct": 1.0,
        "p25": 25000, "p50": 32000, "p75": 35000,
    },
    "willamette-university-college-of-law": {
        "total_students": 341, "total_receiving": 319, "pct_receiving": 94.0,
        "lt_half_n": 226, "lt_half_pct": 66.0,
        "half_to_full_n": 81, "half_to_full_pct": 24.0,
        "full_n": 3, "full_pct": 1.0,
        "gt_full_n": 9, "gt_full_pct": 3.0,
        "p25": 10000, "p50": 20000, "p75": 30000,
    },
    "william-and-mary-law-school": {
        "total_students": 516, "total_receiving": 340, "pct_receiving": 66.0,
        "lt_half_n": 123, "lt_half_pct": 24.0,
        "half_to_full_n": 189, "half_to_full_pct": 37.0,
        "full_n": 3, "full_pct": 1.0,
        "gt_full_n": 25, "gt_full_pct": 5.0,
        "p25": 21782, "p50": 29302, "p75": 37112,
    },
    "wilmington-university-law": {
        "total_students": 46, "total_receiving": 25, "pct_receiving": 54.0,
        "lt_half_n": 19, "lt_half_pct": 41.0,
        "half_to_full_n": 5, "half_to_full_pct": 11.0,
        "full_n": 1, "full_pct": 2.0,
        "gt_full_n": 0, "gt_full_pct": None,
        "p25": 3000, "p50": 6000, "p75": 24302,
    },
    "wisconsin-law-school": {
        "total_students": 670, "total_receiving": 631, "pct_receiving": 94.0,
        "lt_half_n": 253, "lt_half_pct": 38.0,
        "half_to_full_n": 250, "half_to_full_pct": 37.0,
        "full_n": 3, "full_pct": None,
        "gt_full_n": 125, "gt_full_pct": 19.0,
        "p25": 15000, "p50": 35000, "p75": 52332,
    },
    "university-of-wyoming-college-of-law": {
        "total_students": 229, "total_receiving": 203, "pct_receiving": 89.0,
        "lt_half_n": 126, "lt_half_pct": 55.0,
        "half_to_full_n": 62, "half_to_full_pct": 27.0,
        "full_n": 1, "full_pct": None,
        "gt_full_n": 14, "gt_full_pct": 6.0,
        "p25": 6000, "p50": 8000, "p75": 13460,
    },
    "yale-law-school": {
        "total_students": 615, "total_receiving": 388, "pct_receiving": 63.0,
        "lt_half_n": 208, "lt_half_pct": 34.0,
        "half_to_full_n": 93, "half_to_full_pct": 15.0,
        "full_n": 0, "full_pct": None,
        "gt_full_n": 87, "gt_full_pct": 14.0,
        "p25": 24825, "p50": 34747, "p75": 47286,
    },
}

# --- BUILD RANKED SCHOOL LIST ---
schools = [s for s in df['school_slug'].unique() if s in slug_map]
ranked_schools = sorted(
    schools,
    key=lambda s: medians.get(s, {}).get('rank', 999)
)

# --- FULL ABA SCHOOL MAP (all ranked schools, 2026-27 US News) ---
full_slug_map = {
    "university-of-akron-school-of-law": "Akron",
    "university-of-alabama-school-of-law": "Alabama",
    "albany-law-school": "Albany",
    "american-university-law-school": "American",
    "appalachian-school-of-law": "Appalachian",
    "asu-law": "Arizona State",
    "university-of-arizona-law-school": "Arizona",
    "arkansas-little-rock-law": "Arkansas LR",
    "university-of-arkansas-school-of-law": "Arkansas",
    "john-marshall-atlanta": "John Marshall ATL",
    "ave-maria-school-of-law": "Ave Maria",
    "university-of-baltimore-school-of-law": "Baltimore",
    "barry-university-law": "Barry",
    "baylor-law-school": "Baylor",
    "belmont-university-law": "Belmont",
    "boston-college-law-school": "Boston College",
    "boston-university-law-school": "Boston University",
    "byu-law-school": "BYU",
    "brooklyn-law-school": "Brooklyn",
    "university-of-buffalo-law": "Buffalo",
    "california-western-school-of-law": "Cal Western",
    "uc-berkeley-law-school": "UC Berkeley",
    "uc-davis-law-school": "UC Davis",
    "uc-irvine-law": "UC Irvine",
    "ucla-law": "UC Los Angeles",
    "uc-law-san-francisco": "UCSF (Hastings)",
    "campbell-university-law": "Campbell",
    "capital-university-law": "Capital",
    "cardozo-school-of-law": "Cardozo",
    "case-western-reserve-university-school-of-law": "Case Western Reserve",
    "catholic-university-law": "Catholic",
    "chapman-university-fowler-school-of-law": "Chapman",
    "charleston-school-of-law": "Charleston",
    "chicago-law-school": "Uchicago",
    "chicago-kent-college-of-law": "Chicago-Kent",
    "university-of-cincinnati-college-of-law": "Cincinnati",
    "cuny-law-school": "CUNY",
    "cleveland-state-law": "Cleveland State",
    "colorado-law": "Colorado",
    "columbia-law-school": "Columbia",
    "uconn-law-school": "Connecticut",
    "cooley-law-school": "Cooley",
    "cornell-law-school": "Cornell",
    "creighton-university-law": "Creighton",
    "university-of-dayton-school-of-law": "Dayton",
    "denver-law-school": "Denver",
    "depaul-university-law": "DePaul",
    "university-of-detroit-mercy-school-of-law": "Detroit Mercy",
    "udc-law-school": "Disctrict of Columbia",
    "drake-university-law": "Drake",
    "drexel-university-thomas-r-kline-school-of-law": "Drexel",
    "duke-law": "Duke",
    "duquesne-university-law": "Duquesne",
    "elon-university-law": "Elon",
    "emory-law-school": "Emory",
    "faulkner-university-law": "Faulkner",
    "florida-am-law": "Florida A&M",
    "florida-international-university-college-of-law": "Florida International",
    "florida-state-university-college-of-law": "Florida State",
    "university-of-florida-levin-college-of-law": "Florida",
    "fordham-law-school": "Fordham",
    "antonin-scalia-law-school": "George Mason",
    "gw-law-school": "George Washington",
    "georgetown-law": "Georgetown",
    "georgia-state-university-college-of-law": "Georgia State",
    "uga-law-school": "Georgia",
    "gonzaga-university-law": "Gonzaga",
    "harvard-law-school": "Harvard",
    "university-of-hawaii-william-s-richardson-school-of-law": "Hawaii",
    "hofstra-university-school-of-law": "Hofstra",
    "university-of-houston-law-center": "Houston",
    "howard-university-school-of-law": "Howard",
    "university-of-idaho-college-of-law": "Idaho",
    "university-of-illinois-chicago-school-of-law": "Illinois",
    "uic-john-marshall-law-school": "Illinois-Chicago",
    "indiana-university-maurer-school-of-law": "Indiana (Bloomington)",
    "indiana-university-indianapolis-law": "Indiana (Indianapolis)",
    "inter-american-university-pr-law": "Inter American PR",
    "iowa-law": "Iowa",
    "jacksonville-university-law": "Jacksonville",
    "kansas-law": "Kansas",
    "university-of-kentucky-college-of-law": "Kentucky",
    "lewis-and-clark-law-school": "Lewis & Clark",
    "liberty-university-law": "Liberty",
    "lincoln-memorial-university-law": "Lincoln Memorial",
    "louisiana-state-university-law": "LSU",
    "university-of-louisville-louis-d-brandeis-school-of-law": "Louisville",
    "loyola-law-school-los-angeles": "LMU LA",
    "loyola-university-chicago-school-of-law": "Loyola Chicago",
    "loyola-university-new-orleans-college-of-law": "LU NO",
    "university-of-maine-school-of-law": "Maine",
    "marquette-university-law-school": "Marquette",
    "university-of-maryland-carey-school-of-law": "Maryland",
    "umass-dartmouth-law": "UMASS Dartmouth",
    "university-of-memphis-school-of-law": "Memphis",
    "mercer-university-school-of-law": "Mercer",
    "university-of-miami-law": "Miami",
    "michigan-state-university-college-of-law": "Michigan State",
    "michigan-law-school": "Michigan",
    "university-of-minnesota-law-school": "Minnesota",
    "mississippi-college-school-of-law": "Mississippi College",
    "university-of-missouri-school-of-law": "Missouri",
    "umkc-law-school": "Missouri-Kansas City",
    "william-mitchell-college-of-law": "Mitchell Hamline",
    "university-of-montana-school-of-law": "Montana",
    "university-of-nebraska-college-of-law": "Nebraska",
    "university-of-nevada-las-vegas-school-of-law": "UNLV",
    "new-england-law-boston": "New England",
    "university-of-new-hampshire-school-of-law": "New Hampshire",
    "university-of-new-mexico-school-of-law": "New Mexico",
    "new-york-law-school": "New York Law School",
    "nyu-law-school": "NYU",
    "north-carolina-central-law": "North Carolina Central",
    "unc-law-school": "North Carolina",
    "university-of-north-dakota-school-of-law": "North Dakota",
    "unt-dallas-law": "North Texas Dallas",
    "northeastern-law-school": "Northeastern",
    "northern-illinois-university-law": "Northern Illinois",
    "northern-kentucky-university-law": "Northern Kentucky",
    "northwestern-law-school": "Northwestern",
    "notre-dame-law-school": "Notre Dame",
    "nova-southeastern-law": "Nova Southeastern",
    "ohio-northern-university-law": "Ohio Northern",
    "oklahoma-city-university-law": "Oklahoma City",
    "oregon-law": "Oregon",
    "pace-university-school-of-law": "Pace",
    "university-of-pacific-law": "Pacific",
    "penn-state-dickinson-law": "Penn State Dickinson",
    "penn-law": "Upenn",
    "pepperdine-university-school-of-law": "Pepperdine",
    "pittsburgh-law-school": "Pittsburgh",
    "university-of-puerto-rico-law": "Puerto Rico",
    "quinnipiac-university-school-of-law": "Quinnipiac",
    "regent-university-law": "Regent",
    "university-of-richmond-school-of-law": "Richmond",
    "roger-williams-university-school-of-law": "Roger Williams",
    "rutgers-law-school": "Rutgers",
    "saint-louis-university-law": "Saint Louis",
    "samford-university-cumberland-school-of-law": "Samford",
    "usd-law-school": "San Diego",
    "university-of-san-francisco-law": "San Francisco",
    "santa-clara-university-school-of-law": "Santa Clara",
    "seattle-university-school-of-law": "Seattle",
    "seton-hall-university-school-of-law": "Seton Hall",
    "university-of-south-carolina-school-of-law": "South Carolina",
    "university-of-south-dakota-school-of-law": "South Dakota",
    "south-texas-college-of-law-houston": "South Texas (Houston)",
    "usc-gould-school-of-law": "USC",
    "southern-illinois-university-law": "Southern Illinois",
    "smu-dedman-school-of-law": "SMU",
    "southern-university-law": "Southern",
    "southwestern-law-school": "Southwestern",
    "st-john-s-university-school-of-law": "St. John's",
    "st-mary-s-university-school-of-law": "St. Mary's",
    "st-thomas-university-miami-law": "St. Thomas Miami",
    "st-thomas-minneapolis-law": "St. Thomas Minneapolis",
    "stanford-law-school": "Stanford",
    "stetson-university-law": "Stetson",
    "suffolk-university-law-school": "Suffolk",
    "syracuse-university-law": "Syracuse",
    "temple-university-beasley-school-of-law": "Temple",
    "university-of-tennessee-college-of-law": "Tennessee",
    "texas-am-law-school": "Texas A&M",
    "texas-southern-university-law": "Texas Southern",
    "texas-tech-university-school-of-law": "Texas Tech",
    "ut-austin-law-school": "Texas",
    "moritz-college-of-law": "Ohio State",
    "university-of-toledo-college-of-law": "Toledo",
    "touro-law-center": "Touro",
    "tulane-university-law-school": "Tulane",
    "tulsa-university-law": "Tulsa",
    "university-of-utah-sjd": "Utah",
    "vanderbilt-law-school": "Vanderbilt",
    "vermont-law-school": "Vermont",
    "villanova-university-charles-widger-school-of-law": "Villanova",
    "uva-law-school": "UVA",
    "wake-forest-law-school": "Wake Forest",
    "washburn-university-school-of-law": "Washburn",
    "washington-and-lee-law-school": "Washington and Lee",
    "washington-university-school-of-law": "WashU",
    "university-of-washington-school-of-law": "Washington",
    "wayne-state-university-law": "Wayne State",
    "west-virginia-university-college-of-law": "West Virginia",
    "western-new-england-university-school-of-law": "Western New England",
    "western-state-law": "Western State",
    "widener-commonwealth-law": "Widener Commonwealth",
    "widener-delaware-law": "Widener Delaware",
    "willamette-university-college-of-law": "Willamette University",
    "william-and-mary-law-school": "William & Mary",
    "wilmington-university-law": "Wilmington",
    "wisconsin-law-school": "Wisconsin",
    "university-of-wyoming-college-of-law": "Wyoming",
    "yale-law-school": "Yale",
}

# --- US NEWS RANKINGS (2025, with 2024 for history) ---
rankings = {
    "university-of-akron-school-of-law": {"rank": 131, "rank_2024": None, "rank_delta": 0},
    "university-of-alabama-school-of-law": {"rank": 40, "rank_2024": None, "rank_delta": -9},
    "albany-law-school": {"rank": 120, "rank_2024": None, "rank_delta": 0},
    "american-university-law-school": {"rank": 108, "rank_2024": None, "rank_delta": 0},
    "appalachian-school-of-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "asu-law": {"rank": 44, "rank_2024": None, "rank_delta": 1},
    "university-of-arizona-law-school": {"rank": 70, "rank_2024": None, "rank_delta": -11},
    "arkansas-little-rock-law": {"rank": 100, "rank_2024": None, "rank_delta": 0},
    "university-of-arkansas-school-of-law": {"rank": 140, "rank_2024": None, "rank_delta": 0},
    "john-marshall-atlanta": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "ave-maria-school-of-law": {"rank": 164, "rank_2024": None, "rank_delta": 0},
    "university-of-baltimore-school-of-law": {"rank": 136, "rank_2024": None, "rank_delta": 0},
    "barry-university-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "baylor-law-school": {"rank": 34, "rank_2024": None, "rank_delta": 9},
    "belmont-university-law": {"rank": 85, "rank_2024": None, "rank_delta": -1},
    "boston-college-law-school": {"rank": 20, "rank_2024": None, "rank_delta": 5},
    "boston-university-law-school": {"rank": 24, "rank_2024": None, "rank_delta": -2},
    "byu-law-school": {"rank": 24, "rank_2024": None, "rank_delta": 4},
    "brooklyn-law-school": {"rank": 105, "rank_2024": None, "rank_delta": 0},
    "university-of-buffalo-law": {"rank": 82, "rank_2024": None, "rank_delta": 12},
    "california-western-school-of-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "uc-berkeley-law-school": {"rank": 16, "rank_2024": None, "rank_delta": -3},
    "uc-davis-law-school": {"rank": 52, "rank_2024": None, "rank_delta": -2},
    "uc-irvine-law": {"rank": 34, "rank_2024": None, "rank_delta": 4},
    "ucla-law": {"rank": 13, "rank_2024": None, "rank_delta": -1},
    "uc-law-san-francisco": {"rank": 85, "rank_2024": None, "rank_delta": 3},
    "campbell-university-law": {"rank": 131, "rank_2024": None, "rank_delta": 0},
    "capital-university-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "cardozo-school-of-law": {"rank": 59, "rank_2024": None, "rank_delta": 4},
    "case-western-reserve-university-school-of-law": {"rank": 100, "rank_2024": None, "rank_delta": 0},
    "catholic-university-law": {"rank": 70, "rank_2024": None, "rank_delta": 1},
    "chapman-university-fowler-school-of-law": {"rank": 112, "rank_2024": None, "rank_delta": 0},
    "charleston-school-of-law": {"rank": 171, "rank_2024": None, "rank_delta": 0},
    "chicago-law-school": {"rank": 2, "rank_2024": None, "rank_delta": 1},
    "chicago-kent-college-of-law": {"rank": 105, "rank_2024": None, "rank_delta": 0},
    "university-of-cincinnati-college-of-law": {"rank": 82, "rank_2024": None, "rank_delta": -11},
    "cuny-law-school": {"rank": 171, "rank_2024": None, "rank_delta": 0},
    "cleveland-state-law": {"rank": 136, "rank_2024": None, "rank_delta": 0},
    "colorado-law": {"rank": 54, "rank_2024": None, "rank_delta": -8},
    "columbia-law-school": {"rank": 9, "rank_2024": None, "rank_delta": 1},
    "uconn-law-school": {"rank": 58, "rank_2024": None, "rank_delta": -8},
    "cooley-law-school": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "cornell-law-school": {"rank": 13, "rank_2024": None, "rank_delta": 5},
    "creighton-university-law": {"rank": 144, "rank_2024": None, "rank_delta": 0},
    "university-of-dayton-school-of-law": {"rank": 112, "rank_2024": None, "rank_delta": 0},
    "denver-law-school": {"rank": 91, "rank_2024": None, "rank_delta": -3},
    "depaul-university-law": {"rank": 131, "rank_2024": None, "rank_delta": 0},
    "university-of-detroit-mercy-school-of-law": {"rank": 142, "rank_2024": None, "rank_delta": 0},
    "udc-law-school": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "drake-university-law": {"rank": 91, "rank_2024": None, "rank_delta": -7},
    "drexel-university-thomas-r-kline-school-of-law": {"rank": 82, "rank_2024": None, "rank_delta": -3},
    "duke-law": {"rank": 7, "rank_2024": None, "rank_delta": -1},
    "duquesne-university-law": {"rank": 85, "rank_2024": None, "rank_delta": 7},
    "elon-university-law": {"rank": 144, "rank_2024": None, "rank_delta": 0},
    "emory-law-school": {"rank": 40, "rank_2024": None, "rank_delta": -2},
    "faulkner-university-law": {"rank": 164, "rank_2024": None, "rank_delta": 0},
    "florida-am-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "florida-international-university-college-of-law": {"rank": 77, "rank_2024": None, "rank_delta": 7},
    "florida-state-university-college-of-law": {"rank": 34, "rank_2024": None, "rank_delta": 4},
    "university-of-florida-levin-college-of-law": {"rank": 34, "rank_2024": None, "rank_delta": 4},
    "fordham-law-school": {"rank": 42, "rank_2024": None, "rank_delta": -4},
    "antonin-scalia-law-school": {"rank": 32, "rank_2024": None, "rank_delta": -1},
    "gw-law-school": {"rank": 26, "rank_2024": None, "rank_delta": 5},
    "georgetown-law": {"rank": 18, "rank_2024": None, "rank_delta": -4},
    "georgia-state-university-college-of-law": {"rank": 77, "rank_2024": None, "rank_delta": 2},
    "uga-law-school": {"rank": 26, "rank_2024": None, "rank_delta": -4},
    "gonzaga-university-law": {"rank": 144, "rank_2024": None, "rank_delta": 0},
    "harvard-law-school": {"rank": 6, "rank_2024": None, "rank_delta": 0},
    "university-of-hawaii-william-s-richardson-school-of-law": {"rank": 91, "rank_2024": None, "rank_delta": 8},
    "hofstra-university-school-of-law": {"rank": 117, "rank_2024": None, "rank_delta": 0},
    "university-of-houston-law-center": {"rank": 54, "rank_2024": None, "rank_delta": 9},
    "howard-university-school-of-law": {"rank": 117, "rank_2024": None, "rank_delta": 0},
    "university-of-idaho-college-of-law": {"rank": 148, "rank_2024": None, "rank_delta": 0},
    "university-of-illinois-chicago-school-of-law": {"rank": 46, "rank_2024": None, "rank_delta": 2},
    "uic-john-marshall-law-school": {"rank": 167, "rank_2024": None, "rank_delta": 0},
    "indiana-university-maurer-school-of-law": {"rank": 49, "rank_2024": None, "rank_delta": -3},
    "indiana-university-indianapolis-law": {"rank": 124, "rank_2024": None, "rank_delta": 0},
    "inter-american-university-pr-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "iowa-law": {"rank": 32, "rank_2024": None, "rank_delta": 4},
    "jacksonville-university-law": {"rank": 999, "rank_2024": None, "rank_delta": 0},
    "kansas-law": {"rank": 46, "rank_2024": None, "rank_delta": 4},
    "university-of-kentucky-college-of-law": {"rank": 70, "rank_2024": None, "rank_delta": -2},
    "lewis-and-clark-law-school": {"rank": 112, "rank_2024": None, "rank_delta": -13},
    "liberty-university-law": {"rank": 144, "rank_2024": None, "rank_delta": 0},
    "lincoln-memorial-university-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "louisiana-state-university-law": {"rank": 85, "rank_2024": None, "rank_delta": -1},
    "university-of-louisville-louis-d-brandeis-school-of-law": {"rank": 124, "rank_2024": None, "rank_delta": 0},
    "loyola-law-school-los-angeles": {"rank": 70, "rank_2024": None, "rank_delta": 1},
    "loyola-university-chicago-school-of-law": {"rank": 70, "rank_2024": None, "rank_delta": 9},
    "loyola-university-new-orleans-college-of-law": {"rank": 135, "rank_2024": None, "rank_delta": 0},
    "university-of-maine-school-of-law": {"rank": 91, "rank_2024": None, "rank_delta": -3},
    "marquette-university-law-school": {"rank": 59, "rank_2024": None, "rank_delta": 0},
    "university-of-maryland-carey-school-of-law": {"rank": 62, "rank_2024": None, "rank_delta": 1},
    "umass-dartmouth-law": {"rank": 171, "rank_2024": None, "rank_delta": 0},
    "university-of-memphis-school-of-law": {"rank": 140, "rank_2024": None, "rank_delta": 0},
    "mercer-university-school-of-law": {"rank": 108, "rank_2024": None, "rank_delta": 0},
    "university-of-miami-law": {"rank": 70, "rank_2024": None, "rank_delta": 22},
    "michigan-state-university-college-of-law": {"rank": 108, "rank_2024": None, "rank_delta": 0},
    "michigan-law-school": {"rank": 9, "rank_2024": None, "rank_delta": -1},
    "university-of-minnesota-law-school": {"rank": 22, "rank_2024": None, "rank_delta": -2},
    "mississippi-college-school-of-law": {"rank": 161, "rank_2024": None, "rank_delta": 0},
    "university-of-missouri-school-of-law": {"rank": 59, "rank_2024": None, "rank_delta": -2},
    "umkc-law-school": {"rank": 100, "rank_2024": None, "rank_delta": -1},
    "william-mitchell-college-of-law": {"rank": 152, "rank_2024": None, "rank_delta": 0},
    "university-of-montana-school-of-law": {"rank": 90, "rank_2024": None, "rank_delta": 9},
    "university-of-nebraska-college-of-law": {"rank": 62, "rank_2024": None, "rank_delta": 9},
    "university-of-nevada-las-vegas-school-of-law": {"rank": 91, "rank_2024": None, "rank_delta": -12},
    "new-england-law-boston": {"rank": 168, "rank_2024": None, "rank_delta": 0},
    "university-of-new-hampshire-school-of-law": {"rank": 136, "rank_2024": None, "rank_delta": 0},
    "university-of-new-mexico-school-of-law": {"rank": 117, "rank_2024": None, "rank_delta": 0},
    "new-york-law-school": {"rank": 112, "rank_2024": None, "rank_delta": 0},
    "nyu-law-school": {"rank": 7, "rank_2024": None, "rank_delta": 1},
    "north-carolina-central-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "unc-law-school": {"rank": 18, "rank_2024": None, "rank_delta": 0},
    "university-of-north-dakota-school-of-law": {"rank": 156, "rank_2024": None, "rank_delta": 0},
    "unt-dallas-law": {"rank": 159, "rank_2024": None, "rank_delta": 0},
    "northeastern-law-school": {"rank": 77, "rank_2024": None, "rank_delta": -9},
    "northern-illinois-university-law": {"rank": 161, "rank_2024": None, "rank_delta": 0},
    "northern-kentucky-university-law": {"rank": 131, "rank_2024": None, "rank_delta": 0},
    "northwestern-law-school": {"rank": 9, "rank_2024": None, "rank_delta": 1},
    "notre-dame-law-school": {"rank": 20, "rank_2024": None, "rank_delta": 0},
    "nova-southeastern-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "ohio-northern-university-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "oklahoma-city-university-law": {"rank": 62, "rank_2024": None, "rank_delta": 0},
    "oregon-law": {"rank": 91, "rank_2024": None, "rank_delta": 3},
    "pace-university-school-of-law": {"rank": 142, "rank_2024": None, "rank_delta": 0},
    "university-of-pacific-law": {"rank": 152, "rank_2024": None, "rank_delta": 0},
    "penn-state-dickinson-law": {"rank": 62, "rank_2024": None, "rank_delta": -3},
    "penn-law": {"rank": 4, "rank_2024": None, "rank_delta": 1},
    "pepperdine-university-school-of-law": {"rank": 46, "rank_2024": None, "rank_delta": 9},
    "pittsburgh-law-school": {"rank": 77, "rank_2024": None, "rank_delta": 2},
    "university-of-puerto-rico-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "quinnipiac-university-school-of-law": {"rank": 136, "rank_2024": None, "rank_delta": 0},
    "regent-university-law": {"rank": 91, "rank_2024": None, "rank_delta": 3},
    "university-of-richmond-school-of-law": {"rank": 62, "rank_2024": None, "rank_delta": 9},
    "roger-williams-university-school-of-law": {"rank": 171, "rank_2024": None, "rank_delta": 0},
    "rutgers-law-school": {"rank": 100, "rank_2024": None, "rank_delta": 0},
    "saint-louis-university-law": {"rank": 91, "rank_2024": None, "rank_delta": 3},
    "samford-university-cumberland-school-of-law": {"rank": 122, "rank_2024": None, "rank_delta": 0},
    "usd-law-school": {"rank": 54, "rank_2024": None, "rank_delta": 3},
    "university-of-san-francisco-law": {"rank": 85, "rank_2024": None, "rank_delta": 0},
    "santa-clara-university-school-of-law": {"rank": 150, "rank_2024": None, "rank_delta": 0},
    "seattle-university-school-of-law": {"rank": 128, "rank_2024": None, "rank_delta": 0},
    "seton-hall-university-school-of-law": {"rank": 70, "rank_2024": None, "rank_delta": 1},
    "university-of-south-carolina-school-of-law": {"rank": 62, "rank_2024": None, "rank_delta": 1},
    "university-of-south-dakota-school-of-law": {"rank": 122, "rank_2024": None, "rank_delta": 0},
    "south-texas-college-of-law-houston": {"rank": 128, "rank_2024": None, "rank_delta": 0},
    "usc-gould-school-of-law": {"rank": 26, "rank_2024": None, "rank_delta": 0},
    "southern-illinois-university-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "smu-dedman-school-of-law": {"rank": 42, "rank_2024": None, "rank_delta": 1},
    "southern-university-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "southwestern-law-school": {"rank": 161, "rank_2024": None, "rank_delta": 0},
    "st-john-s-university-school-of-law": {"rank": 62, "rank_2024": None, "rank_delta": 1},
    "st-mary-s-university-school-of-law": {"rank": 150, "rank_2024": None, "rank_delta": 0},
    "st-thomas-university-miami-law": {"rank": 105, "rank_2024": None, "rank_delta": 0},
    "st-thomas-minneapolis-law": {"rank": 105, "rank_2024": None, "rank_delta": -11},
    "stanford-law-school": {"rank": 1, "rank_2024": None, "rank_delta": 0},
    "stetson-university-law": {"rank": 91, "rank_2024": None, "rank_delta": 8},
    "suffolk-university-law-school": {"rank": 128, "rank_2024": None, "rank_delta": 0},
    "syracuse-university-law": {"rank": 100, "rank_2024": None, "rank_delta": 0},
    "temple-university-beasley-school-of-law": {"rank": 49, "rank_2024": None, "rank_delta": 1},
    "university-of-tennessee-college-of-law": {"rank": 57, "rank_2024": None, "rank_delta": -2},
    "texas-am-law-school": {"rank": 22, "rank_2024": None, "rank_delta": 0},
    "texas-southern-university-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "texas-tech-university-school-of-law": {"rank": 85, "rank_2024": None, "rank_delta": 3},
    "ut-austin-law-school": {"rank": 16, "rank_2024": None, "rank_delta": -2},
    "moritz-college-of-law": {"rank": 30, "rank_2024": None, "rank_delta": -2},
    "university-of-toledo-college-of-law": {"rank": 149, "rank_2024": None, "rank_delta": 0},
    "touro-law-center": {"rank": 164, "rank_2024": None, "rank_delta": 0},
    "tulane-university-law-school": {"rank": 70, "rank_2024": None, "rank_delta": 1},
    "tulsa-university-law": {"rank": 120, "rank_2024": None, "rank_delta": 0},
    "university-of-utah-sjd": {"rank": 44, "rank_2024": None, "rank_delta": -13},
    "vanderbilt-law-school": {"rank": 12, "rank_2024": None, "rank_delta": 2},
    "vermont-law-school": {"rank": 154, "rank_2024": None, "rank_delta": 0},
    "villanova-university-charles-widger-school-of-law": {"rank": 49, "rank_2024": None, "rank_delta": -1},
    "uva-law-school": {"rank": 4, "rank_2024": None, "rank_delta": 0},
    "wake-forest-law-school": {"rank": 30, "rank_2024": None, "rank_delta": -4},
    "washburn-university-school-of-law": {"rank": 108, "rank_2024": None, "rank_delta": 0},
    "washington-and-lee-law-school": {"rank": 34, "rank_2024": None, "rank_delta": 2},
    "washington-university-school-of-law": {"rank": 13, "rank_2024": None, "rank_delta": 1},
    "university-of-washington-school-of-law": {"rank": 52, "rank_2024": None, "rank_delta": -2},
    "wayne-state-university-law": {"rank": 62, "rank_2024": None, "rank_delta": 9},
    "west-virginia-university-college-of-law": {"rank": 124, "rank_2024": None, "rank_delta": 0},
    "western-new-england-university-school-of-law": {"rank": 159, "rank_2024": None, "rank_delta": 0},
    "western-state-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "widener-commonwealth-law": {"rank": 175, "rank_2024": None, "rank_delta": 0},
    "widener-delaware-law": {"rank": 154, "rank_2024": None, "rank_delta": 0},
    "willamette-university-college-of-law": {"rank": 168, "rank_2024": None, "rank_delta": 0},
    "william-and-mary-law-school": {"rank": 34, "rank_2024": None, "rank_delta": -3},
    "wilmington-university-law": {"rank": 999, "rank_2024": None, "rank_delta": 0},
    "wisconsin-law-school": {"rank": 26, "rank_2024": None, "rank_delta": 2},
    "university-of-wyoming-college-of-law": {"rank": 112, "rank_2024": None, "rank_delta": 0},
    "yale-law-school": {"rank": 2, "rank_2024": None, "rank_delta": -1},
}

# --- RANKS for additional ABA schools (not in rankings dict above) ---
additional_ranks = {
    s: v["rank"] for s, v in rankings.items()
}

# Merged lookup for rank display, sorting, and position indicators in full mode
full_medians = {}
# Start with rankings dict for all ranked schools
for s, r in rankings.items():
    rank_val = r["rank"] if r["rank"] not in (999, None) else None  # 999 = truly unranked (UR)
    full_medians[s] = {"rank": rank_val, "rank_2024": r.get("rank_2024"), "rank_delta": r.get("rank_delta", 0)}
# Overlay lsat/gpa from admissions_data for all schools
for s, adm in admissions_data.items():
    if s not in full_medians:
        full_medians[s] = {}
    if adm.get("lsat50"):
        full_medians[s]["lsat"] = adm["lsat50"]
    if adm.get("gpa50"):
        full_medians[s]["gpa"]  = adm["gpa50"]
    if adm.get("lsat25"):
        full_medians[s]["lsat25"] = adm["lsat25"]
    if adm.get("lsat75"):
        full_medians[s]["lsat75"] = adm["lsat75"]
    if adm.get("gpa25"):
        full_medians[s]["gpa25"] = adm["gpa25"]
    if adm.get("gpa75"):
        full_medians[s]["gpa75"] = adm["gpa75"]
# Overlay curated medians (most accurate lsat/gpa + rank)
for s, v in medians.items():
    if s not in full_medians:
        full_medians[s] = {}
    full_medians[s].update(v)
    # Patch rank from rankings dict if available
    if s in rankings:
        full_medians[s]["rank"] = rankings[s]["rank"]
        full_medians[s]["rank_2024"] = rankings[s].get("rank_2024")
        full_medians[s]["rank_delta"] = rankings[s].get("rank_delta", 0)

# Build all_ranked_schools from full_slug_map, sorted by rank
all_slugs_in_data = set(df['school_slug'].unique())
unmatched = all_slugs_in_data - set(full_slug_map.keys())
if unmatched:
    print(f"⚠️  Slugs in data but NOT in full_slug_map ({len(unmatched)}): {sorted(unmatched)}")
else:
    print("✅ All slugs matched!")
all_ranked_schools = sorted(
    full_slug_map.keys(),
    key=lambda s: (rankings.get(s, {}).get('rank') or 9999)
)

# --- STATE → PUBLIC SCHOOLS MAPPING ---
STATE_TO_PUBLIC_SCHOOLS = {
    "AL": ["university-of-alabama-school-of-law"],
    "AR": ["arkansas-little-rock-law", "university-of-arkansas-school-of-law"],
    "AZ": ["asu-law", "university-of-arizona-law-school"],
    "CA": ["uc-berkeley-law-school", "uc-davis-law-school", "uc-irvine-law", "uc-law-san-francisco", "ucla-law"],
    "CO": ["colorado-law"],
    "CT": ["uconn-law-school"],
    "DC": ["udc-law-school"],
    "FL": ["florida-am-law", "florida-international-university-college-of-law", "florida-state-university-college-of-law", "university-of-florida-levin-college-of-law"],
    "GA": ["georgia-state-university-college-of-law", "uga-law-school"],
    "HI": ["university-of-hawaii-william-s-richardson-school-of-law"],
    "IA": ["iowa-law"],
    "ID": ["university-of-idaho-college-of-law"],
    "IL": ["uic-john-marshall-law-school", "university-of-illinois-chicago-school-of-law", "southern-illinois-university-law"],
    "IN": ["indiana-university-maurer-school-of-law", "indiana-university-indianapolis-law"],
    "KS": ["kansas-law", "washburn-university-school-of-law"],
    "KY": ["university-of-kentucky-college-of-law", "northern-kentucky-university-law", "university-of-louisville-louis-d-brandeis-school-of-law"],
    "LA": ["louisiana-state-university-law", "southern-university-law"],
    "MA": ["umass-dartmouth-law"],
    "MD": ["university-of-baltimore-school-of-law", "university-of-maryland-carey-school-of-law"],
    "ME": ["university-of-maine-school-of-law"],
    "MI": ["michigan-law-school", "michigan-state-university-college-of-law", "wayne-state-university-law"],
    "MN": ["university-of-minnesota-law-school"],
    "MO": ["umkc-law-school", "university-of-missouri-school-of-law"],
    "MT": ["university-of-montana-school-of-law"],
    "NC": ["north-carolina-central-law", "unc-law-school"],
    "ND": ["university-of-north-dakota-school-of-law"],
    "NE": ["university-of-nebraska-college-of-law"],
    "NH": ["university-of-new-hampshire-school-of-law"],
    "NM": ["university-of-new-mexico-school-of-law"],
    "NV": ["university-of-nevada-las-vegas-school-of-law"],
    "NY": ["cuny-law-school", "university-of-buffalo-law"],
    "OH": ["cleveland-state-law", "moritz-college-of-law", "university-of-akron-school-of-law", "university-of-cincinnati-college-of-law", "university-of-toledo-college-of-law"],
    "OK": ["university-of-oklahoma-college-of-law"],
    "OR": ["oregon-law"],
    "PA": ["pittsburgh-law-school"],
    "SC": ["university-of-south-carolina-school-of-law"],
    "SD": ["university-of-south-dakota-school-of-law"],
    "TN": ["university-of-memphis-school-of-law", "university-of-tennessee-college-of-law"],
    "TX": ["texas-am-law-school", "texas-southern-university-law", "texas-tech-university-school-of-law", "university-of-houston-law-center", "unt-dallas-law", "ut-austin-law-school"],
    "UT": ["university-of-utah-sjd"],
    "VA": ["antonin-scalia-law-school", "uva-law-school", "william-and-mary-law-school"],
    "VT": ["vermont-law-school"],
    "WA": ["university-of-washington-school-of-law"],
    "WI": ["wisconsin-law-school"],
    "WV": ["west-virginia-university-college-of-law"],
    "WY": ["university-of-wyoming-college-of-law"],
}
def get_prediction(slug, u_lsat, u_gpa, user_state=None):
    """Returns (pred_value, source_label, n_nearby, tier) where tier is one of: gt_full, full, half_to_full, lt_half, None."""
    import math

    def norm_cdf(z):
        return 0.5 * (1 + math.erf(z / math.sqrt(2)))

    # These schools use need-based aid only — merit predictions don't apply
    NO_MERIT_AID = {'harvard-law-school', 'stanford-law-school', 'yale-law-school'}
    if slug in NO_MERIT_AID:
        return None, None, 0, None
    df_s = df[df['school_slug'] == slug].copy().dropna(subset=['lsat', 'gpa'])
    is_instate = bool(user_state and slug in STATE_TO_PUBLIC_SCHOOLS.get(user_state, []))
    if user_state and 'is_in_state' in df_s.columns:
        f = df_s[df_s['is_in_state'].astype(bool) == is_instate]
        if len(f) >= 5: df_s = f
    df_schol = df_s[df_s['scholarship'] > 0].copy()
    if len(df_schol) >= 5:
        df_schol['_dist'] = (((df_schol['lsat']-u_lsat)/60)**2 + ((df_schol['gpa']-u_gpa)/2)**2)**0.5
        n_nearby = int((df_schol['_dist'] <= ((7/60)**2 + (0.35/2)**2)**0.5).sum())
        if n_nearby >= 5:
            # If user is significantly above the school's 75th LSAT, KNN may
            # underpredict because few reporters represent true outlier offers.
            # In that case, trust ABA tier system instead.
            adm_check = admissions_data.get(slug, {})
            _lsat75 = adm_check.get('lsat75')
            _gpa75  = adm_check.get('gpa75')
            _lsat_outlier = _lsat75 and (u_lsat - _lsat75) >= 6
            _gpa_outlier  = _gpa75  and (u_gpa  - _gpa75)  >= 0.1
            use_knn_over_aba = not (_lsat_outlier and _gpa_outlier)

            if use_knn_over_aba:
                k = max(30, int(len(df_schol) * 0.2))
                neighbors = df_schol.nsmallest(k, '_dist')
                pred_med = neighbors['scholarship'].median()
                aba_pct_r = grant_data.get(slug, {}).get('pct_receiving')
                label = f"lsd.law KNN ({n_nearby} nearby · {aba_pct_r:.0f}% receive aid per ABA)" if aba_pct_r else f"lsd.law KNN ({n_nearby} nearby)"
                # Determine tier from predicted amount vs tuition
                si = school_info.get(slug, {})
                _t = si.get('tuition') or si.get('tuition_res') or (si.get('credit_oos') or si.get('credit_res') or 0) * 31
                full_3yr = (_t * 3) if _t else None
                tier = None
                if full_3yr:
                    ratio = pred_med / full_3yr
                    if ratio >= 1.0:    tier = 'gt_full'
                    elif ratio >= 0.9:  tier = 'full'
                    elif ratio >= 0.45: tier = 'half_to_full'
                    else:               tier = 'lt_half'
                return pred_med, label, n_nearby, tier

    # --- Fall back to tier-adjusted ABA ---
    g = grant_data.get(slug, {})
    aba_med = (g.get("p50") or 0) * 3 or None
    if not aba_med:
        return None, None, 0, None

    predicted_tier = None
    adm = admissions_data.get(slug, {})
    lsat25 = adm.get('lsat25'); lsat75 = adm.get('lsat75')
    gpa25  = adm.get('gpa25');  gpa75  = adm.get('gpa75')
    if lsat25 and lsat75 and gpa25 and gpa75:
        lsat_iqr = max(1, lsat75 - lsat25)
        gpa_iqr  = max(0.01, gpa75 - gpa25)
        lsat_z = (u_lsat - (lsat75 + lsat25) / 2) / (lsat_iqr / 1.35)
        gpa_z  = (u_gpa  - (gpa75  + gpa25)  / 2) / (gpa_iqr  / 1.35)
        # 70/30 weighting
        combined_z = lsat_z * 0.7 + gpa_z * 0.3
        # Bonus for LSAT above 75th (schools recruit outlier applicants)
        if u_lsat > lsat75:
            combined_z += (u_lsat - lsat75) / lsat_iqr * 0.5
        # Stronger penalty for GPA below 25th (schools protect median stats)
        if u_gpa < gpa25:
            combined_z -= (gpa25 - u_gpa) / gpa_iqr * 1.0
        # Moderate penalty for GPA below median (between 25th and 50th)
        elif u_gpa < (gpa75 + gpa25) / 2:
            combined_z -= ((gpa75 + gpa25) / 2 - u_gpa) / gpa_iqr * 0.4
        pct_receiving = (g.get('pct_receiving') or 0) / 100
        # If pct_receiving is missing, estimate from tier data
        if not pct_receiving:
            gt_n = g.get('gt_full_n') or 0; fp_n = g.get('full_n') or 0
            hf_n = g.get('half_to_full_n') or 0; lt_n = g.get('lt_half_n') or 0
            total_recv = gt_n + fp_n + hf_n + lt_n
            total_students = g.get('total_students') or 0
            if total_students > 0 and total_recv > 0:
                pct_receiving = total_recv / total_students
            else:
                pct_receiving = 0.75  # reasonable default
        if pct_receiving > 0:
            rank_among_recipients = max(0, min(1,
                (norm_cdf(combined_z) - (1 - pct_receiving)) / pct_receiving
            ))
            gt_full_pct   = (g.get('gt_full_pct')     or 0) / 100
            full_pct      = (g.get('full_pct')         or 0) / 100
            half_full_pct = (g.get('half_to_full_pct') or 0) / 100
            si = school_info.get(slug, {})
            is_instate_for_tier = bool(user_state and slug in STATE_TO_PUBLIC_SCHOOLS.get(user_state, []))
            if is_instate_for_tier:
                _t = si.get('tuition_res') or (si.get('credit_res') or 0) * 31 or si.get('tuition')
            else:
                _t = si.get('tuition') or si.get('tuition_res') or (si.get('credit_oos') or si.get('credit_res') or 0) * 31
            full_3yr = (_t * 3) if _t else None
            if full_3yr:
                if rank_among_recipients >= (1 - gt_full_pct) and gt_full_pct > 0:
                    aba_med = int(full_3yr * 1.1)
                    predicted_tier = 'gt_full'
                elif rank_among_recipients >= (1 - gt_full_pct - full_pct) and (gt_full_pct + full_pct) > 0:
                    aba_med = full_3yr
                    predicted_tier = 'full'
                elif rank_among_recipients >= (1 - gt_full_pct - full_pct - half_full_pct):
                    # Interpolate within half_to_full band based on position in band
                    band_bottom = 1 - gt_full_pct - full_pct - half_full_pct
                    band_top    = 1 - gt_full_pct - full_pct
                    if band_top > band_bottom:
                        position = (rank_among_recipients - band_bottom) / (band_top - band_bottom)
                    else:
                        position = 0.5
                    # Scale from 50% tuition (bottom of band) to 100% tuition (top of band)
                    aba_med = int(full_3yr * (0.5 + position * 0.5))
                    predicted_tier = 'half_to_full'
                else:
                    # Interpolate within lt_half band
                    band_top = 1 - gt_full_pct - full_pct - half_full_pct
                    if band_top > 0:
                        position = rank_among_recipients / band_top
                    else:
                        position = 0.5
                    # Scale from 0% (bottom) to 50% tuition (top of lt_half)
                    aba_med = int(full_3yr * position * 0.5)
                    predicted_tier = 'lt_half'

    return aba_med, "adjusted ABA estimate", 0, predicted_tier


app = Dash(__name__, suppress_callback_exceptions=True)
server = app.server  # expose for gunicorn

# Inject CSS for the school list hover/active states
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>Law School Scholarship Visualizer</title>
        {%favicon%}
        {%css%}
        <style>
            @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500&display=swap');

            body {
                background-color: #0a0a0a;
                color: white;
                font-family: 'DM Sans', sans-serif;
                margin: 0;
                padding: 16px;
            }

            .school-item {
                display: flex;
                align-items: center;
                gap: 6px;
                padding: 6px 10px;
                border-radius: 6px;
                cursor: pointer;
                transition: background 0.15s ease;
                border-left: 3px solid transparent;
                color: #999;
                font-size: 14px;
                font-weight: 400;
                letter-spacing: 0.01em;
                user-select: none;
            }

            .school-item:hover {
                background: #1e1e1e;
                color: #e0e0e0;
            }

            .school-item.active {
                background: #1a1a2e;
                border-left: 3px solid #c8a96e;
                color: #f0e6d0;
                font-weight: 500;
            }

            .school-rank {
                font-size: 10px;
                color: #555;
                min-width: 22px;
                text-align: right;
                font-variant-numeric: tabular-nums;
            }

            .school-item.active .school-rank {
                color: #c8a96e;
            }

            .grant-bar-row {
                margin-bottom: 5px;
            }
            .grant-bar-label {
                display: flex;
                justify-content: space-between;
                font-size: 14px;
                color: #aaa;
                margin-bottom: 2px;
            }
            .grant-bar-label span:last-child {
                color: #ddd;
                font-variant-numeric: tabular-nums;
            }
            .grant-bar-track {
                height: 5px;
                background: #1e1e1e;
                border-radius: 3px;
                overflow: hidden;
            }
            .grant-bar-fill {
                height: 100%;
                border-radius: 3px;
                transition: width 0.4s ease;
            }
            .grant-section-header {
                font-size: 11px;
                letter-spacing: 0.12em;
                color: #555;
                font-weight: 500;
                margin: 14px 0 8px;
                text-transform: uppercase;
            }
            .percentile-row {
                display: flex;
                justify-content: space-between;
                font-size: 14px;
                color: #888;
                padding: 4px 0;
            }
            .percentile-row span:last-child {
                color: #c8a96e;
                font-weight: 500;
            }
                width: 4px;
            }
            ::-webkit-scrollbar-track {
                background: transparent;
            }
            ::-webkit-scrollbar-thumb {
                background: #333;
                border-radius: 2px;
            }
            /* State dropdown styling */
            /* State dropdown — catch all backgrounds */
            .state-dropdown,
            .state-dropdown > div,
            .state-dropdown > div > div,
            .state-dropdown .Select-control,
            .state-dropdown .VirtualizedSelectFocusedOption,
            .state-dropdown .Select-menu,
            .state-dropdown .Select-menu-outer,
            .state-dropdown .Select-option,
            .state-dropdown [class*="menu"],
            .state-dropdown [class*="control"],
            .state-dropdown [class*="option"],
            .state-dropdown [class*="single-value"],
            .state-dropdown [class*="placeholder"],
            .state-dropdown [class*="container"],
            .state-dropdown [class*="indicator"] {
                background-color: #111 !important;
                color: #c8a96e !important;
                border-color: #c8a96e44 !important;
            }
            .state-dropdown .Select-control {
                border: 1px solid #c8a96e44 !important;
                border-radius: 6px !important;
                height: 38px !important;
                cursor: pointer !important;
                box-shadow: none !important;
            }
            .state-dropdown .Select-control:hover {
                border-color: #c8a96e88 !important;
                box-shadow: none !important;
            }
            .state-dropdown .Select-value-label,
            .state-dropdown .Select-value {
                color: #c8a96e !important;
                font-size: 16px !important;
                font-weight: 600 !important;
                font-family: 'DM Serif Display', serif !important;
                line-height: 36px !important;
            }
            .state-dropdown .Select-placeholder {
                color: #555 !important;
                font-size: 14px !important;
                line-height: 36px !important;
            }
            .state-dropdown .Select-arrow-zone .Select-arrow {
                border-color: #555 transparent transparent !important;
            }
            .state-dropdown.is-open .Select-arrow {
                border-color: transparent transparent #555 !important;
            }
            .state-dropdown .Select-menu-outer {
                border: 1px solid #c8a96e44 !important;
                border-radius: 6px !important;
                margin-top: 2px !important;
                z-index: 9999 !important;
                box-shadow: 0 4px 12px rgba(0,0,0,0.5) !important;
            }
            .state-dropdown .Select-menu {
                border-radius: 6px !important;
                max-height: 200px !important;
            }
            .state-dropdown .Select-option {
                font-size: 13px !important;
                padding: 6px 10px !important;
                cursor: pointer !important;
            }
            .state-dropdown .Select-option:hover,
            .state-dropdown .Select-option.is-focused {
                background-color: #1e1e0e !important;
                color: #f0e6d0 !important;
            }
            .state-dropdown .Select-option.is-selected {
                background-color: #2a2a1a !important;
                font-weight: 600 !important;
            }
            .state-dropdown .Select-clear-zone {
                color: #555 !important;
            }
            .state-dropdown .Select-clear-zone:hover {
                color: #c8a96e !important;
            }
            .state-dropdown input {
                color: #c8a96e !important;
                background: transparent !important;
            }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
'''

CITY_MAP = {
    "stanford-law-school": "Palo Alto, CA",
    "yale-law-school": "New Haven, CT",
    "chicago-law-school": "Chicago, IL",
    "columbia-law-school": "New York, NY",
    "harvard-law-school": "Cambridge, MA",
    "penn-law": "Philadelphia, PA",
    "uva-law-school": "Charlottesville, VA",
    "northwestern-law-school": "Chicago, IL",
    "duke-law": "Durham, NC",
    "georgetown-law": "Washington, DC",
    "cornell-law-school": "Ithaca, NY",
    "texas-am-law-school": "Fort Worth, TX",
    "ut-austin-law-school": "Austin, TX",
    "michigan-law-school": "Ann Arbor, MI",
    "university-of-minnesota-law-school": "Minneapolis, MN",
    "vanderbilt-law-school": "Nashville, TN",
    "washington-university-school-of-law": "St. Louis, MO",
    "emory-law-school": "Atlanta, GA",
    "uga-law-school": "Athens, GA",
    "georgia-state-university-college-of-law": "Atlanta, GA",
    "notre-dame-law-school": "Notre Dame, IN",
    "boston-college-law-school": "Newton, MA",
    "boston-university-law-school": "Boston, MA",
    "northeastern-law-school": "Boston, MA",
    "suffolk-university-law-school": "Boston, MA",
    "new-england-law-boston": "Boston, MA",
    "tulane-university-law-school": "New Orleans, LA",
    "loyola-university-new-orleans-college-of-law": "New Orleans, LA",
    "louisiana-state-university-law": "Baton Rouge, LA",
    "fordham-law-school": "New York, NY",
    "cardozo-school-of-law": "New York, NY",
    "brooklyn-law-school": "Brooklyn, NY",
    "nyu-law-school": "New York, NY",
    "cuny-law-school": "Long Island City, NY",
    "new-york-law-school": "New York, NY",
    "st-john-s-university-school-of-law": "Jamaica, NY",
    "hofstra-university-school-of-law": "Hempstead, NY",
    "pace-university-school-of-law": "White Plains, NY",
    "touro-law-center": "Central Islip, NY",
    "albany-law-school": "Albany, NY",
    "syracuse-university-law": "Syracuse, NY",
    "university-of-buffalo-law": "Buffalo, NY",
    "uc-berkeley-law-school": "Berkeley, CA",
    "uc-davis-law-school": "Davis, CA",
    "uc-irvine-law": "Irvine, CA",
    "uc-law-san-francisco": "San Francisco, CA",
    "ucla-law": "Los Angeles, CA",
    "usc-gould-school-of-law": "Los Angeles, CA",
    "loyola-law-school-los-angeles": "Los Angeles, CA",
    "southwestern-law-school": "Los Angeles, CA",
    "pepperdine-university-school-of-law": "Malibu, CA",
    "santa-clara-university-school-of-law": "Santa Clara, CA",
    "university-of-san-francisco-law": "San Francisco, CA",
    "usd-law-school": "San Diego, CA",
    "california-western-school-of-law": "San Diego, CA",
    "university-of-pacific-law": "Sacramento, CA",
    "chapman-university-fowler-school-of-law": "Orange, CA",
    "western-state-law": "Fullerton, CA",
    "moritz-college-of-law": "Columbus, OH",
    "case-western-reserve-university-school-of-law": "Cleveland, OH",
    "cleveland-state-law": "Cleveland, OH",
    "university-of-akron-school-of-law": "Akron, OH",
    "university-of-cincinnati-college-of-law": "Cincinnati, OH",
    "university-of-toledo-college-of-law": "Toledo, OH",
    "ohio-northern-university-law": "Ada, OH",
    "gw-law-school": "Washington, DC",
    "american-university-law-school": "Washington, DC",
    "catholic-university-law": "Washington, DC",
    "howard-university-school-of-law": "Washington, DC",
    "udc-law-school": "Washington, DC",
    "university-of-maryland-carey-school-of-law": "Baltimore, MD",
    "university-of-baltimore-school-of-law": "Baltimore, MD",
    "pittsburgh-law-school": "Pittsburgh, PA",
    "duquesne-university-law": "Pittsburgh, PA",
    "drexel-university-thomas-r-kline-school-of-law": "Philadelphia, PA",
    "temple-university-beasley-school-of-law": "Philadelphia, PA",
    "villanova-university-charles-widger-school-of-law": "Villanova, PA",
    "widener-commonwealth-law": "Carlisle, PA",
    "widener-delaware-law": "Wilmington, DE",
    "rutgers-law-school": "Newark, NJ",
    "seton-hall-university-school-of-law": "Newark, NJ",
    "florida-state-university-college-of-law": "Tallahassee, FL",
    "university-of-florida-levin-college-of-law": "Gainesville, FL",
    "florida-am-law": "Orlando, FL",
    "florida-international-university-college-of-law": "Miami, FL",
    "university-of-miami-law": "Coral Gables, FL",
    "nova-southeastern-law": "Fort Lauderdale, FL",
    "barry-university-law": "Orlando, FL",
    "stetson-university-law": "Gulfport, FL",
    "st-thomas-university-miami-law": "Miami Gardens, FL",
    "university-of-washington-school-of-law": "Seattle, WA",
    "seattle-university-school-of-law": "Seattle, WA",
    "gonzaga-university-law": "Spokane, WA",
    "lewis-and-clark-law-school": "Portland, OR",
    "oregon-law": "Eugene, OR",
    "willamette-university-college-of-law": "Salem, OR",
    "colorado-law": "Boulder, CO",
    "denver-law-school": "Denver, CO",
    "asu-law": "Phoenix, AZ",
    "university-of-arizona-law-school": "Tucson, AZ",
    "iowa-law": "Iowa City, IA",
    "drake-university-law": "Des Moines, IA",
    "kansas-law": "Lawrence, KS",
    "washburn-university-school-of-law": "Topeka, KS",
    "university-of-nebraska-college-of-law": "Lincoln, NE",
    "creighton-university-law": "Omaha, NE",
    "university-of-south-dakota-school-of-law": "Vermillion, SD",
    "university-of-north-dakota-school-of-law": "Grand Forks, ND",
    "university-of-montana-school-of-law": "Missoula, MT",
    "university-of-wyoming-college-of-law": "Laramie, WY",
    "university-of-idaho-college-of-law": "Moscow, ID",
    "university-of-nevada-las-vegas-school-of-law": "Las Vegas, NV",
    "university-of-utah-sjd": "Salt Lake City, UT",
    "byu-law-school": "Provo, UT",
    "university-of-hawaii-william-s-richardson-school-of-law": "Honolulu, HI",
    "university-of-new-mexico-school-of-law": "Albuquerque, NM",
    "university-of-new-hampshire-school-of-law": "Concord, NH",
    "uconn-law-school": "Hartford, CT",
    "quinnipiac-university-school-of-law": "Hamden, CT",
    "university-of-maine-school-of-law": "Portland, ME",
    "vermont-law-school": "South Royalton, VT",
    "roger-williams-university-school-of-law": "Bristol, RI",
    "western-new-england-university-school-of-law": "Springfield, MA",
    "umass-dartmouth-law": "Dartmouth, MA",
    "william-and-mary-law-school": "Williamsburg, VA",
    "antonin-scalia-law-school": "Arlington, VA",
    "regent-university-law": "Virginia Beach, VA",
    "university-of-richmond-school-of-law": "Richmond, VA",
    "washington-and-lee-law-school": "Lexington, VA",
    "appalachian-school-of-law": "Grundy, VA",
    "liberty-university-law": "Lynchburg, VA",
    "unc-law-school": "Chapel Hill, NC",
    "north-carolina-central-law": "Durham, NC",
    "campbell-university-law": "Raleigh, NC",
    "elon-university-law": "Greensboro, NC",
    "wake-forest-law-school": "Winston-Salem, NC",
    "charleston-school-of-law": "Charleston, SC",
    "university-of-south-carolina-school-of-law": "Columbia, SC",
    "mercer-university-school-of-law": "Macon, GA",
    "john-marshall-atlanta": "Atlanta, GA",
    "samford-university-cumberland-school-of-law": "Birmingham, AL",
    "university-of-alabama-school-of-law": "Tuscaloosa, AL",
    "faulkner-university-law": "Montgomery, AL",
    "mississippi-college-school-of-law": "Jackson, MS",
    "university-of-kentucky-college-of-law": "Lexington, KY",
    "university-of-louisville-louis-d-brandeis-school-of-law": "Louisville, KY",
    "northern-kentucky-university-law": "Highland Heights, KY",
    "university-of-tennessee-college-of-law": "Knoxville, TN",
    "belmont-university-law": "Nashville, TN",
    "lincoln-memorial-university-law": "Harrogate, TN",
    "indiana-university-maurer-school-of-law": "Bloomington, IN",
    "indiana-university-indianapolis-law": "Indianapolis, IN",
    "notre-dame-law-school": "Notre Dame, IN",
    "valparaiso-university-law": "Valparaiso, IN",
    "university-of-illinois-chicago-school-of-law": "Chicago, IL",
    "uic-john-marshall-law-school": "Chicago, IL",
    "depaul-university-law": "Chicago, IL",
    "loyola-university-chicago-school-of-law": "Chicago, IL",
    "chicago-kent-college-of-law": "Chicago, IL",
    "northern-illinois-university-law": "DeKalb, IL",
    "southern-illinois-university-law": "Carbondale, IL",
    "michigan-state-university-college-of-law": "East Lansing, MI",
    "wayne-state-university-law": "Detroit, MI",
    "university-of-detroit-mercy-school-of-law": "Detroit, MI",
    "cooley-law-school": "Lansing, MI",
    "marquette-university-law-school": "Milwaukee, WI",
    "wisconsin-law-school": "Madison, WI",
    "william-mitchell-college-of-law": "St. Paul, MN",
    "st-thomas-minneapolis-law": "Minneapolis, MN",
    "hamline-university-law": "St. Paul, MN",
    "university-of-missouri-school-of-law": "Columbia, MO",
    "umkc-law-school": "Kansas City, MO",
    "saint-louis-university-law": "St. Louis, MO",
    "smu-dedman-school-of-law": "Dallas, TX",
    "texas-tech-university-school-of-law": "Lubbock, TX",
    "texas-southern-university-law": "Houston, TX",
    "university-of-houston-law-center": "Houston, TX",
    "south-texas-college-of-law-houston": "Houston, TX",
    "unt-dallas-law": "Dallas, TX",
    "baylor-law-school": "Waco, TX",
    "st-mary-s-university-school-of-law": "San Antonio, TX",
    "arkansas-little-rock-law": "Little Rock, AR",
    "university-of-arkansas-school-of-law": "Fayetteville, AR",
    "university-of-oklahoma-college-of-law": "Norman, OK",
    "oklahoma-city-university-law": "Oklahoma City, OK",
    "tulsa-university-law": "Tulsa, OK",
    "penn-state-dickinson-law": "Carlisle, PA",
    "inter-american-university-pr-law": "San Juan, PR",
    "university-of-puerto-rico-law": "San Juan, PR",
    "jacksonville-university-law": "Jacksonville, FL",
}


app.layout = html.Div([
    dcc.Location(id='url', refresh=False),

    # --- NAV BAR ---
    html.Div([
        html.Span("Law School Scholarship Visualizer", style={
            "fontFamily": "'DM Serif Display', serif",
            "fontSize": "18px",
            "color": "#f0e6d0",
            "letterSpacing": "0.02em",
            "marginRight": "32px",
        }),
        html.A("Visualizer", href="/", id='nav-vis', style={
            "fontSize": "11px", "letterSpacing": "0.1em", "color": "#c8a96e",
            "textDecoration": "none", "padding": "4px 12px",
            "borderRadius": "4px", "border": "1px solid #c8a96e44",
            "backgroundColor": "#1a1a12", "marginRight": "8px",
        }),
        html.A("Cost Comparison", href="/comparison", id='nav-comp', style={
            "fontSize": "11px", "letterSpacing": "0.1em", "color": "#555",
            "textDecoration": "none", "padding": "4px 12px",
            "borderRadius": "4px", "border": "1px solid #2a2a2a",
        }),
    ], style={
        "display": "flex", "alignItems": "center",
        "padding": "10px 20px", "borderBottom": "1px solid #1a1a1a",
        "backgroundColor": "#060606",
    }),

    # school-title lives here permanently so callbacks always find it
    html.Div(id='school-title', style={"display": "none"}),

    # Page header — shows school name on visualizer, "Cost Comparison" on comparison page
    html.Div(id='page-header', style={
        "textAlign": "center",
        "marginBottom": "8px",
        "marginTop": "8px",
        "minHeight": "24px",
    }),

    # --- PAGE CONTENT ---
    html.Div(id='page-content'),

])

visualizer_layout = html.Div([

        # LEFT: SCHOOL LIST
        html.Div([
            html.Div([
                html.Span("SCHOOLS", style={
                    "fontSize": "10px",
                    "letterSpacing": "0.15em",
                    "color": "#555",
                    "fontWeight": "500",
                }),
            ], style={
                "display": "flex",
                "justifyContent": "space-between",
                "alignItems": "center",
                "padding": "4px 12px 8px",
            }),
            # Search box
            dcc.Input(
                id='school-search',
                type='text',
                placeholder='Search schools...',
                debounce=False,
                style={
                    "width": "100%",
                    "padding": "6px 10px",
                    "backgroundColor": "#1a1a1a",
                    "border": "1px solid #2a2a2a",
                    "borderRadius": "6px",
                    "color": "#ccc",
                    "fontSize": "12px",
                    "fontFamily": "'DM Sans', sans-serif",
                    "boxSizing": "border-box",
                    "margin": "0 0 8px 0",
                    "outline": "none",
                }
            ),
            html.Div(style={"borderBottom": "1px solid #1e1e1e", "marginBottom": "6px"}),
            dcc.Store(id='school-dropdown', data=all_ranked_schools[0] if all_ranked_schools else None),
            dcc.Store(id='starred-schools', data=list(slug_map.keys())),
            html.Div(id='school-list')
        ], style={
            "flex": "1",
            "maxWidth": "210px",
            "backgroundColor": "#0f0f0f",
            "paddingTop": "12px",
            "borderRadius": "10px",
            "overflowY": "auto",
            "height": "85vh",
            "border": "1px solid #1a1a1a"
        }),

        # CENTER: GRAPH + TOGGLES
        html.Div([
            dcc.Graph(id='plot', style={"height": "78vh", "width": "100%"}),
            # OVERLAY TOGGLES below graph
            html.Div([
                dcc.Checklist(
                    id='overlay-toggles',
                    options=[
                        {'label': html.Span('Median LSAT',   style={"color": "#aaa", "fontSize": "11px", "marginLeft": "4px", "marginRight": "12px"}), 'value': 'lsat_med'},
                        {'label': html.Span('Median GPA',    style={"color": "#aaa", "fontSize": "11px", "marginLeft": "4px", "marginRight": "12px"}), 'value': 'gpa_med'},
                        {'label': html.Span('25/75 LSAT',    style={"color": "#aaa", "fontSize": "11px", "marginLeft": "4px", "marginRight": "12px"}), 'value': 'lsat_pct'},
                        {'label': html.Span('25/75 GPA',     style={"color": "#aaa", "fontSize": "11px", "marginLeft": "4px", "marginRight": "12px"}), 'value': 'gpa_pct'},
                        {'label': html.Span('Tuition+Fees (3yr)',style={"color": "#aaa", "fontSize": "11px", "marginLeft": "4px", "marginRight": "12px"}), 'value': 'coa'},
                        {'label': html.Span('50% T+F (3yr)',    style={"color": "#aaa", "fontSize": "11px", "marginLeft": "4px", "marginRight": "12px"}), 'value': 'half_coa'},
                        {'label': html.Span('IS T+F (3yr)',     style={"color": "#aaa", "fontSize": "11px", "marginLeft": "4px", "marginRight": "12px"}), 'value': 'is_coa'},
                        {'label': html.Span('Your Profile',  style={"color": "#aaa", "fontSize": "11px", "marginLeft": "4px", "marginRight": "12px"}), 'value': 'you'},
                        {'label': html.Span('Scholarship Only', style={"color": "#c8a96e", "fontSize": "11px", "marginLeft": "4px", "marginRight": "12px"}), 'value': 'hide_no_schol'},
                        {'label': html.Span('Color by Median', style={"color": "#aaa", "fontSize": "11px", "marginLeft": "4px", "marginRight": "12px"}), 'value': 'color_median'},
                        {'label': html.Span('Surface View',   style={"color": "#aaa", "fontSize": "11px", "marginLeft": "4px", "marginRight": "12px"}), 'value': 'surface_view'},
                    ],
                    value=['coa', 'you', 'hide_no_schol', 'color_median'],
                    inputStyle={"accentColor": "#c8a96e"},
                    labelStyle={"display": "inline-flex", "alignItems": "center"},
                    style={"display": "flex", "flexWrap": "wrap", "gap": "4px", "flex": "1"},
                ),
                html.Button(
                    "⟳ Reset View",
                    id='reset-view-btn',
                    n_clicks=0,
                    style={
                        "fontSize": "10px",
                        "color": "#555",
                        "backgroundColor": "transparent",
                        "border": "1px solid #2a2a2a",
                        "borderRadius": "4px",
                        "padding": "4px 10px",
                        "cursor": "pointer",
                        "fontFamily": "'DM Sans', sans-serif",
                        "letterSpacing": "0.05em",
                        "whiteSpace": "nowrap",
                    }
                ),
                dcc.Store(id='reset-view-store', data=0),
            ], style={
                "backgroundColor": "#0f0f0f",
                "border": "1px solid #1e1e1e",
                "borderRadius": "8px",
                "padding": "10px 14px",
                "marginTop": "8px",
                "display": "flex",
                "alignItems": "center",
                "gap": "12px",
            }),
            # AXIS RANGE CONTROLS
            html.Div([
                html.Div([
                    html.Span("LSAT", style={"fontSize": "9px", "color": "#555", "letterSpacing": "0.1em", "marginRight": "8px", "whiteSpace": "nowrap"}),
                    dcc.RangeSlider(
                        id='lsat-range',
                        min=120, max=180, step=1,
                        value=[140, 180],
                        marks={120: {'label': '120', 'style': {'color': '#444', 'fontSize': '9px'}},
                               140: {'label': '140', 'style': {'color': '#444', 'fontSize': '9px'}},
                               160: {'label': '160', 'style': {'color': '#444', 'fontSize': '9px'}},
                               180: {'label': '180', 'style': {'color': '#444', 'fontSize': '9px'}}},
                        tooltip={"placement": "bottom", "always_visible": False},
                        allowCross=False,
                    ),
                ], style={"display": "flex", "alignItems": "center", "flex": "1", "gap": "4px"}),
                html.Div([
                    html.Span("GPA", style={"fontSize": "9px", "color": "#555", "letterSpacing": "0.1em", "marginRight": "8px", "whiteSpace": "nowrap"}),
                    dcc.RangeSlider(
                        id='gpa-range',
                        min=2.0, max=4.0, step=0.05,
                        value=[2.5, 4.0],
                        marks={2.0: {'label': '2.0', 'style': {'color': '#444', 'fontSize': '9px'}},
                               2.5: {'label': '2.5', 'style': {'color': '#444', 'fontSize': '9px'}},
                               3.0: {'label': '3.0', 'style': {'color': '#444', 'fontSize': '9px'}},
                               3.5: {'label': '3.5', 'style': {'color': '#444', 'fontSize': '9px'}},
                               4.0: {'label': '4.0', 'style': {'color': '#444', 'fontSize': '9px'}}},
                        tooltip={"placement": "bottom", "always_visible": False},
                        allowCross=False,
                    ),
                ], style={"display": "flex", "alignItems": "center", "flex": "1", "gap": "4px"}),
                html.Div([
                    html.Span("AID $k", style={"fontSize": "9px", "color": "#555", "letterSpacing": "0.1em", "marginRight": "8px", "whiteSpace": "nowrap"}),
                    dcc.RangeSlider(
                        id='schol-range',
                        min=0, max=400, step=10,
                        value=[0, 400],
                        marks={0:   {'label': '0',    'style': {'color': '#444', 'fontSize': '9px'}},
                               100: {'label': '100k', 'style': {'color': '#444', 'fontSize': '9px'}},
                               200: {'label': '200k', 'style': {'color': '#444', 'fontSize': '9px'}},
                               300: {'label': '300k', 'style': {'color': '#444', 'fontSize': '9px'}},
                               400: {'label': '400k', 'style': {'color': '#444', 'fontSize': '9px'}}},
                        tooltip={"placement": "bottom", "always_visible": False},
                        allowCross=False,
                    ),
                ], style={"display": "flex", "alignItems": "center", "flex": "1", "gap": "4px"}),
            ], style={
                "backgroundColor": "#0a0a0a",
                "border": "1px solid #1a1a1a",
                "borderRadius": "8px",
                "padding": "8px 14px",
                "marginTop": "6px",
                "display": "flex",
                "gap": "16px",
                "alignItems": "center",
            }),
            html.Div("Scholarship data: lsd.law", style={
                "fontSize": "10px", "color": "#333",
                "textAlign": "right", "marginTop": "5px",
                "letterSpacing": "0.04em",
            }),
        ], style={"flex": "0 0 auto", "width": "min(80vh, 60vw)", "margin": "0 auto"}),
        # RIGHT: MEDIANS + NET COST
        html.Div([
            # YOUR PROFILE card — half width
            html.Div([
            html.Div([
                html.Div("YOUR PROFILE", style={
                    "fontSize": "10px",
                    "letterSpacing": "0.15em",
                    "color": "#c8a96e",
                    "fontWeight": "600",
                    "marginBottom": "10px",
                }),
                html.Div([
                    html.Div([
                        html.Div("LSAT", style={"fontSize": "10px", "color": "#555", "letterSpacing": "0.1em", "marginBottom": "4px"}),
                        dcc.Input(
                            id='user-lsat',
                            type='number',
                            placeholder='172',
                            value=175,
                            min=120,
                            max=180,
                            style={
                                "width": "70px",
                                "padding": "8px 6px",
                                "borderRadius": "6px",
                                "border": "1px solid #c8a96e44",
                                "backgroundColor": "#1a1a12",
                                "color": "#c8a96e",
                                "textAlign": "center",
                                "fontSize": "16px",
                                "fontWeight": "600",
                                "fontFamily": "'DM Serif Display', serif",
                                "boxSizing": "border-box",
                            }
                        ),
                    ], style={"flex": "0 0 auto"}),
                    html.Div([
                        html.Div("GPA", style={"fontSize": "10px", "color": "#555", "letterSpacing": "0.1em", "marginBottom": "4px"}),
                        dcc.Input(
                            id='user-gpa',
                            type='number',
                            placeholder='3.9',
                            value=3.91,
                            step=0.01,
                            min=0,
                            max=4.3,
                            style={
                                "width": "70px",
                                "padding": "8px 6px",
                                "borderRadius": "6px",
                                "border": "1px solid #c8a96e44",
                                "backgroundColor": "#1a1a12",
                                "color": "#c8a96e",
                                "textAlign": "center",
                                "fontSize": "16px",
                                "fontWeight": "600",
                                "fontFamily": "'DM Serif Display', serif",
                                "boxSizing": "border-box",
                            }
                        ),
                    ], style={"flex": "0 0 auto"}),
                    html.Div([
                        html.Div("STATE", style={"fontSize": "10px", "color": "#555", "letterSpacing": "0.1em", "marginBottom": "4px"}),
                        dcc.Dropdown(
                            id='user-state',
                            options=[{"label": s, "value": s} for s in sorted(STATE_TO_PUBLIC_SCHOOLS.keys())],
                            placeholder="ST",
                            clearable=True,
                            style={
                                "width": "80px",
                                "fontSize": "16px",
                                "fontWeight": "600",
                                "fontFamily": "'DM Serif Display', serif",
                                "backgroundColor": "#1a1a12",
                                "border": "1px solid #c8a96e44",
                                "borderRadius": "6px",
                                "color": "#c8a96e",
                            },
                            className="state-dropdown",
                        ),
                    ], style={"flex": "0 0 auto"}),
                ], style={"display": "flex", "gap": "8px", "alignItems": "flex-end"}),
            ], style={
                "backgroundColor": "#0f0f08",
                "border": "1px solid #c8a96e33",
                "borderRadius": "8px",
                "padding": "12px",
                "marginBottom": "12px",
            }),
            ], style={"width": "50%"}),
            html.Div(
                id='median-panel',
                children="Select a school",
                style={"color": "white", "fontSize": "16px", "lineHeight": "1.6"}
            )
        ], style={
            "flex": "1",
            "backgroundColor": "#111",
            "padding": "10px",
            "borderRadius": "12px",
            "color": "white",
            "overflowY": "auto",
            "height": "85vh",
        })

    ], style={
        "display": "flex",
        "gap": "20px",
        "alignItems": "stretch",
        "width": "100%"
    })

# Keep visualizer always in DOM to avoid callback ID issues
# Page routing just shows/hides via display style


# --- ROUTING CALLBACK ---
@app.callback(
    Output('page-header', 'children'),
    Input('url', 'pathname'),
    Input('school-title', 'children'),
)
def update_page_header(pathname, school_title_children):
    if pathname == '/comparison':
        return html.Span("Cost Comparison", style={
            "fontFamily": "'DM Serif Display', serif",
            "fontSize": "22px",
            "color": "#f0e6d0",
            "letterSpacing": "0.02em",
        })
    if not school_title_children:
        return None
    return html.Div(school_title_children, style={
        "display": "flex",
        "flexDirection": "column",
        "alignItems": "center",
    })



@app.callback(
    Output('nav-vis', 'style'),
    Output('nav-comp', 'style'),
    Input('url', 'pathname'),
)
def update_nav(pathname):
    active   = {"fontSize": "11px", "letterSpacing": "0.1em", "color": "#c8a96e",
                 "textDecoration": "none", "padding": "4px 12px",
                 "borderRadius": "4px", "border": "1px solid #c8a96e44",
                 "backgroundColor": "#1a1a12", "marginRight": "8px"}
    inactive = {"fontSize": "11px", "letterSpacing": "0.1em", "color": "#555",
                 "textDecoration": "none", "padding": "4px 12px",
                 "borderRadius": "4px", "border": "1px solid #2a2a2a", "marginRight": "8px"}
    inactive_comp = {**inactive, "marginRight": "0"}
    if pathname == '/comparison':
        return inactive, {**active, "marginRight": "0"}
    return active, inactive_comp



@app.callback(
    Output('page-content', 'children'),
    Input('url', 'pathname'),
)
def route(pathname):
    vis_style = {"display": "none"} if pathname == '/comparison' else {}
    cmp_style = {"display": "none"} if pathname != '/comparison' else {}
    return html.Div([
        html.Div(visualizer_layout, style=vis_style),
        html.Div(build_comparison_layout(), id='cmp-page', style=cmp_style),
    ])
def build_comparison_layout():
    return html.Div([
        # Profile note — uses visualizer inputs
        html.Div([
            html.Span("Using profile from Visualizer  ·  ", style={"fontSize": "11px", "color": "#444", "fontStyle": "italic"}),
            html.Span(id='cmp-profile-display', style={"fontSize": "11px", "color": "#c8a96e"}),
            html.Div([
                html.Div("SORT BY", style={"fontSize": "9px", "color": "#555", "letterSpacing": "0.1em", "marginBottom": "4px"}),
                html.Div([
                    html.Button(label, id=f'cmp-sort-{key}', n_clicks=0, style={
                        "background": "#111", "border": "1px solid #2a2a2a", "borderRadius": "6px",
                        "color": "#555", "fontSize": "11px", "fontFamily": "'DM Sans', sans-serif",
                        "padding": "6px 12px", "cursor": "pointer", "letterSpacing": "0.05em",
                    })
                    for key, label in [('rank','Rank'), ('tf','T+F'), ('net_tf','Net T+F'), ('net_coa','Net COA')]
                ], style={"display": "flex", "gap": "6px"}),
            ], style={"marginLeft": "auto"}),
        ], style={"display": "flex", "alignItems": "center", "gap": "16px",
                  "padding": "12px 24px", "borderBottom": "1px solid #1a1a1a", "backgroundColor": "#0a0a0a"}),

        dcc.Store(id='cmp-sort-store', data='rank'),
        dcc.Store(id='cmp-selected-school', data=None),
        html.Div(id='cmp-stats', style={"padding": "8px 24px", "fontSize": "11px", "color": "#444",
                                         "borderBottom": "1px solid #111", "backgroundColor": "#0a0a0a"}),
        # Split panel
        html.Div([
            # Table side
            html.Div(id='cmp-table', style={"overflowX": "auto", "overflowY": "auto", "height": "calc(100vh - 160px)"}),
            # Detail panel
            html.Div(id='cmp-detail', style={"display": "none"}),
        ], id='cmp-split', style={"display": "flex", "gap": "0"}),
    ])


# --- COMPARISON CALLBACKS ---
@app.callback(
    Output('cmp-sort-store', 'data'),
    [Input(f'cmp-sort-{k}', 'n_clicks') for k in ['rank','tf','net_tf','net_coa']],
    prevent_initial_call=True,
)
def update_cmp_sort(*args):
    triggered = ctx.triggered_id or 'cmp-sort-rank'
    return triggered.replace('cmp-sort-', '')


def _predict(slug, u_lsat, u_gpa, is_instate):
    df_s = df[df['school_slug'] == slug].dropna(subset=['lsat','gpa','scholarship'])
    df_s = df_s[df_s['scholarship'] > 0].copy()
    if 'is_in_state' in df_s.columns:
        f = df_s[df_s['is_in_state'].astype(bool) == is_instate]
        if len(f) >= 5: df_s = f
    if len(df_s) < 5: return None, None, None
    df_s['_d'] = (((df_s['lsat']-u_lsat)/60)**2 + ((df_s['gpa']-u_gpa)/2)**2)**0.5
    nb = df_s.nsmallest(max(30, int(len(df_s)*0.2)), '_d')
    return nb['scholarship'].quantile(0.25), nb['scholarship'].median(), nb['scholarship'].quantile(0.75)


@app.callback(
    Output('cmp-profile-display', 'children'),
    Input('user-lsat', 'value'),
    Input('user-gpa', 'value'),
    Input('user-state', 'value'),
    prevent_initial_call=True,
)
def update_cmp_profile(user_lsat, user_gpa, user_state):
    parts = []
    if user_lsat: parts.append(f"LSAT {int(user_lsat)}")
    if user_gpa:  parts.append(f"GPA {user_gpa}")
    if user_state: parts.append(user_state)
    return "  ·  ".join(parts) if parts else "No profile entered"


@app.callback(
    Output('cmp-selected-school', 'data'),
    Input({'type': 'cmp-row', 'index': ALL}, 'n_clicks'),
    prevent_initial_call=True,
)
def select_cmp_school(n_clicks):
    if not ctx.triggered_id: return None
    slug = ctx.triggered_id['index']
    return slug


@app.callback(
    Output('cmp-split', 'style'),
    Output('cmp-table', 'style'),
    Output('cmp-detail', 'style'),
    Output('cmp-detail', 'children'),
    Output('school-dropdown', 'data', allow_duplicate=True),
    Input('cmp-selected-school', 'data'),
    Input('user-lsat', 'value'),
    Input('user-gpa', 'value'),
    Input('user-state', 'value'),
    prevent_initial_call=True,
)
def update_cmp_detail(selected_slug, user_lsat, user_gpa, user_state):
    split_style = {"display": "flex", "gap": "0", "height": "calc(100vh - 120px)"}
    no_detail_table = {"overflowX": "auto", "overflowY": "auto", "height": "100%", "flex": "1"}
    no_detail = {"display": "none"}

    if not selected_slug:
        return split_style, no_detail_table, no_detail, None, no_update

    table_style = {"overflowX": "auto", "overflowY": "auto", "height": "100%", "flex": "0 0 52%", "borderRight": "1px solid #1a1a1a"}
    detail_style = {"flex": "0 0 48%", "overflowY": "auto", "height": "100%", "backgroundColor": "#080808", "padding": "0"}

    # Build the panel content for this school
    title_children, panel_children = update_median_panel(selected_slug, user_lsat, user_gpa, user_state)

    detail_content = html.Div([
        # Header
        html.Div([
            html.Div(title_children, style={"flex": "1"}),
            html.Span("✕", id='cmp-detail-close', n_clicks=0, style={
                "cursor": "pointer", "color": "#555", "fontSize": "20px",
                "padding": "4px 8px", "lineHeight": "1",
            }),
        ], style={"display": "flex", "alignItems": "flex-start", "justifyContent": "space-between",
                  "padding": "14px 16px 10px", "borderBottom": "1px solid #1a1a1a"}),

        # Panel content
        html.Div(panel_children, style={"padding": "12px 16px", "overflowY": "auto"}),
    ])

    return split_style, table_style, detail_style, detail_content, selected_slug


@app.callback(
    Output('cmp-selected-school', 'data', allow_duplicate=True),
    Input('cmp-detail-close', 'n_clicks'),
    prevent_initial_call=True,
)
def close_cmp_detail(n):
    return None


@app.callback(
    Output('cmp-table', 'children'),
    Output('cmp-stats', 'children'),
    Input('user-lsat', 'value'),
    Input('user-gpa', 'value'),
    Input('user-state', 'value'),
    Input('cmp-sort-store', 'data'),
)
def update_cmp_table(user_lsat, user_gpa, user_state, sort_by):
    try:
        u_lsat = float(user_lsat) if user_lsat else None
        u_gpa  = float(user_gpa)  if user_gpa  else None
    except: u_lsat = u_gpa = None

    state = (user_state or '').upper().strip() or None
    if state and state not in STATE_TO_PUBLIC_SCHOOLS: state = None

    JD = 31
    rows = []
    for slug, name in full_slug_map.items():
        si = school_info.get(slug, {})
        if not si: continue
        rank = rankings.get(slug, {}).get('rank', 999)
        is_instate = bool(state and slug in STATE_TO_PUBLIC_SCHOOLS.get(state, []))
        has_is_oos = (si.get('tuition') and si.get('tuition_res') and si.get('tuition') != si.get('tuition_res')) or \
                     (si.get('credit_res') and si.get('credit_oos') and si.get('credit_res') != si.get('credit_oos'))

        if is_instate:
            t = si.get('tuition_res') or (si.get('credit_res',0)*JD)
            f = si.get('fees_res') or si.get('fees') or 0
        else:
            t = si.get('tuition') or si.get('tuition_res') or ((si.get('credit_oos') or si.get('credit_res') or 0)*JD)
            f = si.get('fees') or si.get('fees_res') or 0
        living = si.get('living') or 0
        tf3  = (t+f)*3 if t else None
        coa3 = (t+f+living)*3 if t else None

        g = grant_data.get(slug, {})
        aba_med = g.get('p50')
        p25 = p50 = p75 = None
        if u_lsat and u_gpa:
            p25, p50, p75 = _predict(slug, u_lsat, u_gpa, is_instate)

        schol = p50 if p50 else aba_med
        net_tf  = (tf3  - schol) if (tf3  and schol) else tf3
        net_coa = (coa3 - schol) if (coa3 and schol) else coa3

        sort_key = {'rank': rank or 999, 'tf': tf3 or 9e9, 'net_tf': net_tf or 9e9, 'net_coa': net_coa or 9e9}

        rows.append({
            'slug': slug,
            'rank': rank, 'rank_str': f"#{rank}" if rank and rank < 999 else "NR",
            'name': name, 'state': si.get('city',''), 'type': ('IS' if is_instate else 'OOS') if has_is_oos else '—',
            'tf3': tf3, 'coa3': coa3, 'aba_med': aba_med, 'pred': p50,
            'net_tf': net_tf, 'net_coa': net_coa,
            '_sort': sort_key.get(sort_by or 'rank', rank or 999),
        })

    rows.sort(key=lambda r: r['_sort'])

    def fmt(v):
        return f"${v:,.0f}" if v and v < 9e8 else "—"

    def cell(val, color="#888", bold=False):
        return html.Td(val, style={"padding": "7px 12px", "color": color, "whiteSpace": "nowrap",
                                    "fontWeight": "600" if bold else "400", "fontSize": "12px"})

    tbody = []
    for i, r in enumerate(rows):
        tc = "#7fff7f" if r['type'] == 'IS' else ("#c8a96e" if r['type'] == 'OOS' else "#333")
        bg = "#0a0a0a" if i % 2 == 0 else "#000"
        tbody.append(html.Tr([
            cell(r['rank_str'], "#444"),
            html.Td(r['name'], style={"padding": "7px 12px", "color": "#d0c8b8", "fontWeight": "500",
                                       "fontSize": "12px", "whiteSpace": "nowrap", "cursor": "pointer"}),
            cell(r['state'], "#444"),
            cell(r['type'], tc),
            cell(fmt(r['tf3'])),
            cell(fmt(r['coa3'])),
            cell(fmt(r['aba_med']), "#c8a96e"),
            cell(fmt(r['pred']), "#ff69b4"),
            cell(fmt(r['net_tf']), "#f4c430", bold=True),
            cell(fmt(r['net_coa']), "#aaa"),
        ], id={'type': 'cmp-row', 'index': r['slug']},
           style={"background": bg, "borderBottom": "1px solid #111", "cursor": "pointer"},
           n_clicks=0))

    th_style = {"padding": "8px 12px", "fontSize": "9px", "letterSpacing": "0.1em",
                 "color": "#444", "borderBottom": "1px solid #1e1e1e", "fontWeight": "500",
                 "whiteSpace": "nowrap", "textAlign": "left"}
    table = html.Table([
        html.Thead(html.Tr([html.Th(h, style=th_style) for h in [
            'Rank','School','State','IS/OOS','T+F (3yr)','COA (3yr)','ABA Aid','Pred. Aid','Net T+F','Net COA'
        ]])),
        html.Tbody(tbody),
    ], style={"width": "100%", "borderCollapse": "collapse", "marginTop": "12px"})

    with_data = sum(1 for r in rows if r['net_tf'])
    stats = f"{len(rows)} schools  ·  {with_data} with cost data"
    if u_lsat and u_gpa: stats += f"  ·  Predictions for LSAT {int(u_lsat)}, GPA {u_gpa}"
    if state: stats += f"  ·  {state} resident rates applied"

    return table, stats



@app.callback(
    Output('overlay-toggles', 'options'),
    Input('school-dropdown', 'data'),
    prevent_initial_call=True,
)
def update_overlay_options(selected_slug):
    def opt(label, value, color="#aaa"):
        return {'label': html.Span(label, style={"color": color, "fontSize": "11px", "marginLeft": "4px", "marginRight": "12px"}), 'value': value}

    base = [
        opt('Median LSAT',       'lsat_med'),
        opt('Median GPA',        'gpa_med'),
        opt('25/75 LSAT',        'lsat_pct'),
        opt('25/75 GPA',         'gpa_pct'),
        opt('Tuition+Fees (3yr)','coa'),
        opt('50% T+F (3yr)',     'half_coa'),
        opt('Your Profile',      'you'),
        opt('Scholarship Only',  'hide_no_schol', '#c8a96e'),
        opt('Color by Median',   'color_median'),
        opt('Surface View',      'surface_view'),
    ]

    if selected_slug:
        si = school_info.get(selected_slug, {})
        t_oos = si.get('tuition'); t_is = si.get('tuition_res')
        cr_oos = si.get('credit_oos'); cr_is = si.get('credit_res')
        has_diff = (t_oos and t_is and t_oos != t_is) or (cr_oos and cr_is and cr_oos != cr_is)
        if has_diff:
            base.insert(6, opt('IS T+F (3yr)', 'is_coa'))

    return base


@app.callback(
    Output('starred-schools', 'data'),
    Input({'type': 'star-btn', 'index': ALL}, 'n_clicks'),
    State('starred-schools', 'data'),
    prevent_initial_call=True,
)
def toggle_star(n_clicks_list, starred):
    if not ctx.triggered_id:
        return starred
    # Only act on actual clicks (n_clicks > 0), ignore resets to 0 from re-renders
    triggered = ctx.triggered
    if not triggered or triggered[0]['value'] in (0, None):
        from dash.exceptions import PreventUpdate
        raise PreventUpdate
    slug = ctx.triggered_id['index']
    starred = list(starred or [])
    if slug in starred:
        starred.remove(slug)
    else:
        starred.append(slug)
    return starred

# --- CALLBACK: handle school list clicks ---

@app.callback(
    Output('school-dropdown', 'data'),
    Input({'type': 'school-item', 'index': ALL}, 'n_clicks'),
    prevent_initial_call=True
)
def select_school(n_clicks_list):
    if not ctx.triggered_id:
        return all_ranked_schools[0] if all_ranked_schools else None
    return ctx.triggered_id['index']

# Re-render school list
@app.callback(
    Output('school-list', 'children'),
    Input('school-dropdown', 'data'),
    Input('starred-schools', 'data'),
    Input('user-lsat', 'value'),
    Input('user-gpa', 'value'),
    Input('school-search', 'value'),
)
def update_school_list_highlight(selected_slug, starred, user_lsat, user_gpa, search):
    starred_set = set(starred or [])
    name_map    = full_slug_map
    rank_map    = full_medians
    search_q    = (search or '').strip().lower()

    try:
        u_lsat = float(user_lsat) if user_lsat is not None else None
        u_gpa  = float(user_gpa)  if user_gpa  is not None else None
    except (ValueError, TypeError):
        u_lsat, u_gpa = None, None

    has_profile = u_lsat is not None and u_gpa is not None

    # Filter by search
    def matches(s):
        if not search_q:
            return True
        name = name_map.get(s, s).lower()
        return search_q in name or search_q in s

    filtered = [s for s in all_ranked_schools if matches(s)]

    # Split into starred (top) and unstarred (bottom)
    starred_list   = [s for s in filtered if s in starred_set]
    unstarred_list = [s for s in filtered if s not in starred_set]

    items = []

    def make_item(s, i_starred):
        m = rank_map.get(s, {})
        rank_str = f"#{m['rank']}" if m.get('rank') else "NR"
        rank_style = {"color": "#444", "fontSize": "9px"} if not m.get('rank') else {}
        name = name_map.get(s) or slug_map.get(s) or s.replace('-', ' ').title().replace(' Law School','').replace(' School Of Law','').replace(' College Of Law','')

        if has_profile and 'lsat' in m and 'gpa' in m:
            above_lsat = u_lsat >= m['lsat']
            above_gpa  = u_gpa  >= m['gpa']
            if above_lsat and above_gpa:
                ind_node = html.Span("▲▲", style={"fontSize": "9px", "color": "#7fff7f", "marginLeft": "auto"})
            elif above_lsat and not above_gpa:
                ind_node = html.Span([html.Span("L↑", style={"color": "#7fff7f"}), html.Span(" G↓", style={"color": "#c85060"})], style={"fontSize": "9px", "marginLeft": "auto"})
            elif above_gpa and not above_lsat:
                ind_node = html.Span([html.Span("L↓", style={"color": "#c85060"}), html.Span(" G↑", style={"color": "#7fff7f"})], style={"fontSize": "9px", "marginLeft": "auto"})
            else:
                ind_node = html.Span("▼▼", style={"fontSize": "9px", "color": "#c85060", "marginLeft": "auto"})
        elif 'lsat' in m and 'gpa' in m:
            ind_node = html.Span(f"{m['lsat']} / {m['gpa']}", style={"fontSize": "9px", "color": "#444", "marginLeft": "auto"})
        else:
            ind_node = html.Span("")

        star_char = "★" if i_starred else "☆"
        star_color = "#c8a96e" if i_starred else "#333"

        return html.Div(
            [
                # Star button — separate from the clickable school area
                html.Span(
                    star_char,
                    id={'type': 'star-btn', 'index': s},
                    n_clicks=0,
                    style={
                        "color": star_color,
                        "fontSize": "11px",
                        "cursor": "pointer",
                        "marginRight": "4px",
                        "flexShrink": "0",
                        "userSelect": "none",
                        "padding": "4px 2px",
                    }
                ),
                # School clickable area — separate div
                html.Div(
                    [
                        html.Span(rank_str, className="school-rank", style=rank_style),
                        html.Span(name, style={"flex": "1"}),
                        ind_node,
                    ],
                    className="school-item" + (" active" if s == selected_slug else ""),
                    id={'type': 'school-item', 'index': s},
                    **{'data-slug': s},
                    style={"display": "flex", "alignItems": "center", "gap": "4px", "flex": "1"},
                ),
            ],
            style={"display": "flex", "alignItems": "center"},
        )

    # Starred section
    if starred_list and not search_q:
        items.append(html.Div("★ STARRED", style={
            "fontSize": "9px", "color": "#c8a96e", "letterSpacing": "0.1em",
            "padding": "4px 10px 2px", "fontWeight": "600",
        }))
        for s in starred_list:
            items.append(make_item(s, True))
        items.append(html.Div(style={"borderTop": "1px solid #1e1e1e", "margin": "6px 4px"}))
        items.append(html.Div("ALL SCHOOLS", style={
            "fontSize": "9px", "color": "#444", "letterSpacing": "0.1em",
            "padding": "4px 10px 2px",
        }))
    elif starred_list and search_q:
        for s in starred_list:
            items.append(make_item(s, True))

    for s in unstarred_list:
        items.append(make_item(s, False))

    return items

    try:
        u_lsat = float(user_lsat) if user_lsat is not None else None
        u_gpa  = float(user_gpa)  if user_gpa  is not None else None
    except (ValueError, TypeError):
        u_lsat, u_gpa = None, None

    has_profile = u_lsat is not None and u_gpa is not None


# --- CALLBACKS ---
@app.callback(
    Output('school-title', 'children'),
    Output('median-panel', 'children'),
    Input('school-dropdown', 'data'),
    Input('user-lsat', 'value'),
    Input('user-gpa', 'value'),
    Input('user-state', 'value'),
)
def update_median_panel(selected_slug, user_lsat, user_gpa, user_state):
    name = full_slug_map.get(selected_slug) or slug_map.get(selected_slug) or selected_slug.replace('-', ' ').title() if selected_slug else "Law School Scholarship Visualizer"
    rank_info = rankings.get(selected_slug, {}) if selected_slug else {}
    rank_str = ""
    delta_str = ""
    if rank_info and rank_info.get('rank') and rank_info['rank'] != 999:
        delta = rank_info.get('rank_delta', 0) or 0
        if delta > 0:
            delta_str = f" ▲{delta}"
        elif delta < 0:
            delta_str = f" ▼{abs(delta)}"
        rank_str = f"#{rank_info['rank']}"
    elif selected_slug:
        _rank = (full_medians.get(selected_slug) or {}).get('rank')
        if _rank and _rank != 999:
            rank_str = f"#{_rank}"

    location = school_info.get(selected_slug, {}).get("city", "") if selected_slug else ""

    # City names for well-known schools

    city_state = CITY_MAP.get(selected_slug, "") if selected_slug else ""
    # Fall back to location field if no city map entry
    if not city_state and location:
        city_state = location

    delta_span = None
    if delta_str:
        color = "#7fff7f" if delta_str.startswith(" ▲") else "#c85060"
        delta_span = html.Span(delta_str, style={"color": color, "fontSize": "14px"})

    subtitle_base = name + rank_str
    city_part = f"  ·  {city_state}" if city_state else ""

    # School name — large and prominent
    title_children = [
        html.Span(name, style={
            "fontFamily": "'DM Serif Display', serif",
            "fontSize": "26px",
            "color": "#f0e6d0",
            "letterSpacing": "0.01em",
            "fontWeight": "400",
            "display": "block",
            "marginBottom": "2px",
        })
    ]

    # Rank + delta + location on second line
    sub_children = []
    if rank_str:
        sub_children.append(html.Span(rank_str, style={
            "fontSize": "13px", "color": "#666", "fontWeight": "500",
        }))
        if delta_span:
            sub_children.append(delta_span)
    if city_part:
        city_text = city_part.strip().lstrip("·").strip()
        sep = "  ·  " if sub_children else ""
        sub_children.append(html.Span(sep + city_text, style={
            "fontSize": "12px", "color": "#444",
        }))

    if sub_children:
        title_children.append(html.Div(sub_children, style={"display": "flex", "alignItems": "center", "gap": "2px"}))

    if not selected_slug:
        return title_children, "No school selected"

    # Build m from medians (curated) or admissions_data (all ABA schools)
    if selected_slug in medians:
        m = medians[selected_slug]
    elif selected_slug in admissions_data:
        adm = admissions_data.get(selected_slug, {})
        rank_entry = rankings.get(selected_slug, {})
        m = {
            "lsat":    adm.get("lsat50"),
            "gpa":     adm.get("gpa50"),
            "lsat25":  adm.get("lsat25"),
            "lsat75":  adm.get("lsat75"),
            "gpa25":   adm.get("gpa25"),
            "gpa75":   adm.get("gpa75"),
            "rank":    rank_entry.get("rank") if rank_entry.get("rank") != 999 else None,
        }
        # Filter out schools with no usable data at all
        if not m["lsat"] and not m["gpa"]:
            return title_children, html.P("No admissions data available", style={"color": "#555", "fontSize": "13px"})
    else:
        return title_children, html.P("No data available", style={"color": "#555", "fontSize": "13px"})

    si = school_info.get(selected_slug)
    school_coa = None
    if si:
        _JD = 31
        _t = (si.get("tuition") or si.get("tuition_res") or
              (si.get("credit_oos") or si.get("credit_res") or 0) * _JD)
        _f = si.get("fees") or si.get("fees_res") or 0
        school_coa = (_t + _f + (si.get("living") or 0)) * 3

    def box(header, content_items, border_color="#1e1e1e"):
        return html.Div([
            html.Div(header, style={
                "fontSize": "11px", "letterSpacing": "0.12em",
                "color": "#555", "fontWeight": "500", "marginBottom": "10px"
            }),
            *content_items
        ], style={
            "backgroundColor": "#0f0f0f",
            "border": f"1px solid {border_color}",
            "borderRadius": "8px",
            "padding": "8px",
            "marginBottom": "6px",
            "width": "100%",
            "boxSizing": "border-box",
        })

    comparison = ""
    if user_lsat is not None and user_gpa is not None:
        try:
            above_lsat = float(user_lsat) >= m['lsat']
            above_gpa  = float(user_gpa)  >= m['gpa']
            if above_lsat and above_gpa:
                comparison = "🔥 Above both medians"
            elif above_lsat or above_gpa:
                comparison = "⚖️ Split medians"
            else:
                comparison = "📉 Below both medians"
        except (ValueError, TypeError):
            pass

    def stat_row(label, val, color="#ccc"):
        return html.Div([
            html.Span(label, style={"color": "#666", "fontSize": "15px", "minWidth": "90px", "display": "inline-block"}),
            html.Span(str(val), style={"color": color, "fontSize": "15px", "fontWeight": "500"}),
        ], style={"marginBottom": "6px"})

    rank_info = rankings.get(selected_slug, {})
    rank_display = []

    # --- ADMISSIONS BOX (prominent LSAT/GPA) ---
    def big_stat(label, val25, val50, val75, user_val=None, color="#f0e6d0"):
        # Determine where the user falls
        marker = None
        if user_val is not None:
            try:
                uv = float(user_val)
                v25 = float(val25) if val25 != '—' else None
                v50 = float(val50) if val50 != '—' else None
                v75 = float(val75) if val75 != '—' else None
                if v75 and uv >= v75:
                    pos_label = "▲ Above 75th"
                    pos_color = "#7fff7f"
                elif v50 and uv >= v50:
                    pos_label = "↑ Above median"
                    pos_color = "#c8a96e"
                elif v25 and uv >= v25:
                    pos_label = "↓ Below median"
                    pos_color = "#f4c430"
                else:
                    pos_label = "▼ Below 25th"
                    pos_color = "#c85060"
                marker = html.Span(pos_label, style={
                    "fontSize": "11px", "color": pos_color,
                    "marginLeft": "10px", "fontFamily": "'DM Sans', sans-serif",
                    "fontWeight": "500", "letterSpacing": "0.04em"
                })
            except (ValueError, TypeError):
                pass

        user_display = None
        if user_val is not None:
            try:
                uv = float(user_val)
                fmt = f"{uv:.2f}" if label == "GPA" else f"{int(uv)}"
                user_display = html.Span(f" ({fmt})", style={
                    "fontSize": "14px", "color": "#c8a96e88",
                    "fontFamily": "'DM Sans', sans-serif"
                })
            except (ValueError, TypeError):
                pass

        return html.Div([
            html.Div([
                html.Span(label, style={"fontSize": "11px", "color": "#555", "letterSpacing": "0.1em"}),
                marker or html.Span(""),
            ], style={"display": "flex", "alignItems": "center", "marginBottom": "2px"}),
            html.Div([
                html.Span(str(val25), style={"color": "#555", "fontSize": "14px"}),
                html.Span(f"  {val50}  ", style={"color": color, "fontSize": "22px", "fontWeight": "600", "fontFamily": "'DM Serif Display', serif"}),
                html.Span(str(val75), style={"color": "#555", "fontSize": "14px"}),
                user_display or html.Span(""),
            ]),
        ], style={"marginBottom": "14px"})

    try:
        u_lsat = float(user_lsat) if user_lsat is not None else None
        u_gpa  = float(user_gpa)  if user_gpa  is not None else None
    except (ValueError, TypeError):
        u_lsat, u_gpa = None, None

    adm = admissions_data.get(selected_slug)

    lsat_gpa_col = html.Div([
        html.Div("25th · median · 75th", style={"fontSize": "11px", "color": "#333", "marginBottom": "10px", "letterSpacing": "0.08em"}),
        big_stat("LSAT", m.get('lsat25','—'), m['lsat'], m.get('lsat75','—'), u_lsat, "#c8a96e"),
        big_stat("GPA",  m.get('gpa25','—'),  m['gpa'],  m.get('gpa75','—'),  u_gpa,  "#c8a96e"),
    ], style={"flex": "1", "minWidth": "0"})

    def adm_stat(label, val):
        return html.Div([
            html.Div(label, style={"fontSize": "10px", "color": "#555", "letterSpacing": "0.08em", "marginBottom": "1px"}),
            html.Div(str(val), style={"fontSize": "14px", "color": "#ccc", "fontWeight": "500"}),
        ], style={"marginBottom": "8px"})

    adm_col_items = []
    if adm:
        if adm.get("apps"):
            adm_col_items.append(adm_stat("APPS", f"{adm['apps']:,}"))
        if adm.get("accept_rate") is not None:
            adm_col_items.append(adm_stat("ACCEPT", f"{adm['accept_rate']:.1f}%"))
        if adm.get("offers"):
            adm_col_items.append(adm_stat("OFFERS", f"{adm['offers']:,}"))
        if adm.get("enrollees"):
            adm_col_items.append(adm_stat("ENROLLED", f"{adm['enrollees']:,}"))
        if adm.get("yield_rate") is not None:
            adm_col_items.append(adm_stat("YIELD", f"{adm['yield_rate']:.1f}%"))

    adm_col = html.Div(adm_col_items, style={
        "flex": "0 0 90px", "width": "90px",
        "borderLeft": "1px solid #1e1e1e",
        "paddingLeft": "10px",
    })

    admissions_items = [
        html.Div([lsat_gpa_col, adm_col], style={"display": "flex", "gap": "0", "alignItems": "flex-start"}),
    ]
    if rank_display:
        admissions_items.append(html.Div(style={"borderTop": "1px solid #1e1e1e", "margin": "6px 0"}))
        admissions_items.extend(rank_display)

    sections = []
    sections.append(box("ADMISSIONS  ·  ABA 509", admissions_items))

    # --- COST BOX ---
    if si is not None:
        df_school = df[df['school_slug'] == selected_slug]

        def fmt_dollar(v, suffix=""):
            return f"${v:,.0f}{suffix}" if v else "—"

        t_oos  = si.get("tuition") or 0
        f_oos  = si.get("fees")    or 0
        t_is   = si.get("tuition_res") or 0
        f_is   = si.get("fees_res")    or 0
        living = si.get("living")  or 0
        credit_res = si.get("credit_res")
        credit_oos = si.get("credit_oos")

        # Determine if user qualifies for in-state
        user_is_instate = bool(user_state and selected_slug in STATE_TO_PUBLIC_SCHOOLS.get(user_state, []))

        has_diff = (
            (t_is and t_oos and t_is != t_oos) or
            (credit_res and credit_oos and credit_res != credit_oos)
        )

        # Scholarship stats from scatter data
        schol_rows = df_school['scholarship'].dropna()
        median_schol = schol_rows.median() if len(schol_rows) else None
        max_schol    = schol_rows.max()    if len(schol_rows) else None

        JD_CREDITS = 31  # credits per year (typical law school year)

        def cost_section(label, tuition, fees, color, credit_res=None, credit_oos=None, highlighted=False):
            # Use per-credit × 31 credits/year as annual tuition estimate
            is_oos = (label == "OUT-OF-STATE")
            if not tuition and (credit_oos or credit_res):
                cr = credit_oos if (is_oos or not credit_res) else credit_res
                tuition = cr * JD_CREDITS if cr else None
            tf_3yr = (tuition + (fees or 0)) * 3 if tuition else None
            coa_3yr = (tf_3yr + living * 3) if tf_3yr else None
            items = [
                html.Div(label, style={"fontSize": "9px", "color": color,
                    "letterSpacing": "0.1em", "fontWeight": "600", "marginBottom": "4px"}),
            ]
            cr_note = ""
            if not (si.get("tuition") or si.get("tuition_res")) and (credit_oos or credit_res):
                cr_val = credit_oos if (is_oos or not credit_res) else credit_res
                cr_note = f"${cr_val:,}/cr × {JD_CREDITS}/yr"
            if tuition:
                label_str = f"Tuition ({cr_note})" if cr_note else "Tuition"
                items += [
                    stat_row(label_str, fmt_dollar(tuition, "/yr")),
                    stat_row("Fees",          fmt_dollar(fees or None, "/yr")),
                ]
            items += [
                stat_row("Living",       fmt_dollar(living or None, "/yr")),
            ]
            if tf_3yr:
                items += [
                    html.Div(style={"borderTop": "1px solid #1e1e1e", "margin": "5px 0"}),
                    stat_row("T+F (3yr)",    fmt_dollar(tf_3yr), "#c8a96e"),
                    stat_row("COA (3yr)",    fmt_dollar(coa_3yr)),
                ]
                if median_schol:
                    net_med = tf_3yr - median_schol
                    net_max = tf_3yr - max_schol if max_schol else None
                    items += [
                        html.Div(style={"borderTop": "1px solid #1e1e1e", "margin": "5px 0"}),
                        stat_row("Median aid", fmt_dollar(median_schol)),
                        stat_row("Net (med)",  fmt_dollar(net_med), "#f4c430"),
                        stat_row("Net (max)",  fmt_dollar(net_max), "#7fff7f"),
                    ]
            if highlighted == 'green':
                return html.Div(items, style={
                    "background": "rgba(127,255,127,0.07)",
                    "border": "1px solid rgba(127,255,127,0.25)",
                    "borderRadius": "6px",
                    "padding": "6px 6px 2px",
                    "marginBottom": "2px",
                })
            elif highlighted == 'red':
                return html.Div(items, style={
                    "background": "rgba(200,80,96,0.07)",
                    "border": "1px solid rgba(200,80,96,0.25)",
                    "borderRadius": "6px",
                    "padding": "6px 6px 2px",
                    "marginBottom": "2px",
                })
            return html.Div(items, style={"padding": "2px 0"})

        if has_diff:
            if user_state:
                oos_color     = "#c8a96e" if user_is_instate else "#7fff7f"
                is_color      = "#7fff7f" if user_is_instate else "#c8a96e"
                oos_highlight = None    if user_is_instate else 'green'
                is_highlight  = 'green' if user_is_instate else None
            else:
                oos_color = is_color = "#c8a96e"
                oos_highlight = is_highlight = None
            oos_section = cost_section("OUT-OF-STATE", t_oos, f_oos, oos_color, credit_res, credit_oos, highlighted=oos_highlight)
            is_section  = cost_section("IN-STATE",     t_is,  f_is,  is_color,  credit_res, credit_res, highlighted=is_highlight)
            cost_items = [oos_section, html.Div(style={"borderTop": "1px solid #2a2a2a", "margin": "8px 0"}), is_section]
        else:
            cost_items = [cost_section("COST  (2025–26)", t_oos or t_is, f_oos or f_is, "#555", credit_res, credit_oos)]

        sections.append(box("COST  (2025–26)", cost_items))
    # --- GRANT AID BOX ---
    g = grant_data.get(selected_slug)
    if g:
        total      = g["total_students"]
        no_aid_n   = total - (g["total_receiving"] or 0)
        no_aid_pct = round(100 * no_aid_n / total) if total else 0

        # Get predicted tier for highlighting
        pred_tier_for_grant = None
        if u_lsat and u_gpa:
            _, _, _, pred_tier_for_grant = get_prediction(selected_slug, u_lsat, u_gpa, user_state)

        bars = [
            ("Stipend",      "gt_full",      g["gt_full_n"],      g["gt_full_pct"],      "#c82060"),
            ("Full tuition", "full",          g["full_n"],          g["full_pct"],         "#c8921f"),
            ("½ – full",     "half_to_full", g["half_to_full_n"], g["half_to_full_pct"], "#27a87c"),
            ("< ½ tuition",  "lt_half",      g["lt_half_n"],      g["lt_half_pct"],      "#3a7fbf"),
            ("No aid",       None,            no_aid_n,             no_aid_pct,            "#555"),
        ]
        max_pct = max((b[3] or 0) for b in bars) or 1
        scale   = 60 / max_pct

        grant_items = [
            html.P(
                f"{g['total_receiving']} of {total} ({g['pct_receiving']}%)",
                style={"fontSize": "13px", "color": "#666", "margin": "0 0 8px"}
            )
        ]
        for label, tier_key, n, pct, fill_color in bars:
            n_str   = str(n) if n is not None else "—"
            pct_str = f"{pct}%" if pct is not None else "—"
            bar_w   = round((pct or 0) * scale, 1)
            is_predicted = (pred_tier_for_grant and tier_key == pred_tier_for_grant)
            row_style = {
                "background": "rgba(200,169,110,0.08)",
                "border": "1px solid #c8a96e55",
                "borderRadius": "4px",
                "padding": "2px 4px",
                "marginBottom": "2px",
            } if is_predicted else {}
            label_el = html.Span([
                label,
                html.Span(" ← predicted", style={"fontSize": "9px", "color": "#c8a96e", "marginLeft": "4px"}) if is_predicted else None,
            ], style={"fontSize": "14px"})
            grant_items.append(html.Div([
                html.Div([
                    label_el,
                    html.Span(f"{n_str} ({pct_str})", style={"fontSize": "13px"}),
                ], className="grant-bar-label"),
                html.Div(
                    html.Div(style={"width": f"{bar_w}%", "background": fill_color, "height": "100%", "borderRadius": "4px"}),
                    className="grant-bar-track", style={"background": "#1e1e1e"}
                )
            ], className="grant-bar-row", style=row_style))

        if any(g[k] for k in ["p25", "p50", "p75"]):
            grant_items.append(html.Div([
                html.Span("GRANT PERCENTILES", style={"fontSize": "9px", "color": "#444", "letterSpacing": "0.08em"}),
                html.Span("  1yr / 3yr", style={"fontSize": "9px", "color": "#333", "letterSpacing": "0.04em"}),
            ], style={"marginTop": "10px", "marginBottom": "4px"}))
            for label, key in [("75th", "p75"), ("50th", "p50"), ("25th", "p25")]:
                val = g[key]
                val3 = val * 3 if val else None
                grant_items.append(html.Div([
                    html.Span(label, style={"marginRight": "6px", "minWidth": "32px", "display": "inline-block"}),
                    html.Span(f"${val:,.0f}" if val else "—", style={"color": "#888"}),
                    html.Span("  /  ", style={"color": "#333"}),
                    html.Span(f"${val3:,.0f}" if val3 else "—", style={"color": "#c8a96e"}),
                ], className="percentile-row", style={"justifyContent": "flex-start"}))

        sections.append(box("GRANT AID  (ABA 509)", grant_items))

    # --- SCHOLARSHIP PREDICTION BOX ---
    try:
        u_lsat = float(user_lsat) if user_lsat else None
        u_gpa  = float(user_gpa)  if user_gpa  else None
    except:
        u_lsat, u_gpa = None, None

    if selected_slug:
        _si = school_info.get(selected_slug, {})
        _user_is_instate = bool(user_state and selected_slug in STATE_TO_PUBLIC_SCHOOLS.get(user_state, []))
        if _user_is_instate:
            _t = _si.get("tuition_res") or (_si.get("credit_res") or 0) * 31
            _f = _si.get("fees_res") or _si.get("fees") or 0
        else:
            _t = _si.get("tuition") or _si.get("tuition_res") or (_si.get("credit_oos") or _si.get("credit_res") or 0) * 31
            _f = _si.get("fees") or _si.get("fees_res") or 0
        _living  = _si.get("living") or 0
        _tf_3yr  = (_t + _f) * 3 if _t else None
        _coa_3yr = (_t + _f + _living) * 3 if _t else None

        # ABA grant data for p25/p75 display
        g = grant_data.get(selected_slug, {})
        aba_p25 = (g.get("p25") or 0) * 3 or None
        aba_p75 = (g.get("p75") or 0) * 3 or None
        aba_pct = g.get("pct_receiving")

        # Use shared prediction function (KNN or adjusted ABA)
        pred_med = pred_p25 = pred_p75 = None
        source_label = None
        pct_got = None
        p25_from_knn = p75_from_knn = False

        if u_lsat and u_gpa:
            pred_med, source_label, n_nearby, pred_tier = get_prediction(selected_slug, u_lsat, u_gpa, user_state)
            if pred_med:
                # Get p25/p75 from KNN if available
                _is_instate = bool(user_state and selected_slug in STATE_TO_PUBLIC_SCHOOLS.get(user_state, []))
                df_s = df[df['school_slug'] == selected_slug].dropna(subset=['lsat','gpa'])
                if user_state and 'is_in_state' in df_s.columns:
                    f = df_s[df_s['is_in_state'].astype(bool) == _is_instate]
                    if len(f) >= 5: df_s = f
                df_schol = df_s[df_s['scholarship'] > 0].copy()
                if len(df_schol) >= 5:
                    df_schol['_dist'] = (((df_schol['lsat']-u_lsat)/60)**2 + ((df_schol['gpa']-u_gpa)/2)**2)**0.5
                    n_near = int((df_schol['_dist'] <= ((7/60)**2+(0.35/2)**2)**0.5).sum())
                    if n_near >= 5:
                        k = max(30, int(len(df_schol)*0.2))
                        nb = df_schol.nsmallest(k, '_dist')
                        pred_p25 = nb['scholarship'].quantile(0.25)
                        pred_p75 = nb['scholarship'].quantile(0.75)
                # Track whether p25/p75 came from KNN or ABA fallback
                p25_from_knn = pred_p25 is not None
                p75_from_knn = pred_p75 is not None
                if not p25_from_knn: pred_p25 = aba_p25
                if not p75_from_knn: pred_p75 = aba_p75
                aba_pct_r = g.get('pct_receiving')
                pct_got = (aba_pct_r / 100) if aba_pct_r else None

        # Use prediction if available
        use_knn = pred_med is not None and source_label and 'lsd.law' in source_label
        _aba_p50 = (g.get("p50") or 0) * 3 or None
        show_p25 = pred_p25 if pred_med else aba_p25
        show_med = pred_med if pred_med else _aba_p50
        show_p75 = pred_p75 if pred_med else aba_p75

        # Scale p25/p75 proportionally whenever the median was adjusted
        # and p25/p75 came from ABA (not KNN)
        if show_med and _aba_p50 and _aba_p50 != show_med:
            scale = show_med / _aba_p50
            if not p25_from_knn and aba_p25: show_p25 = int(aba_p25 * scale)
            if not p75_from_knn and aba_p75: show_p75 = int(aba_p75 * scale)
            # Cap p75 at full tuition 3yr and p25 at 0
            _si2 = school_info.get(selected_slug, {})
            _t2 = _si2.get('tuition') or _si2.get('tuition_res') or (_si2.get('credit_oos') or _si2.get('credit_res') or 0) * 31
            _full_3yr = _t2 * 3 if _t2 else None
            if _full_3yr:
                if show_p75: show_p75 = min(show_p75, int(_full_3yr * 1.05))
            if show_p25: show_p25 = max(0, show_p25)
            # Ensure ordering: p25 <= median <= p75
            if show_p25 and show_p25 > show_med: show_p25 = int(show_med * 0.8)
            if show_p75 and show_p75 < show_med: show_p75 = int(show_med * 1.1)

        # Only show box if we have some prediction value
        if show_med is not None:
            def fmt_k(v):
                return f"${v:,.0f}" if v else "—"

            net_med_tf  = (_tf_3yr  - show_med) if (_tf_3yr  and show_med) else None
            net_p25_tf  = (_tf_3yr  - show_p75) if (_tf_3yr  and show_p75) else None
            net_med_coa = (_coa_3yr - show_med) if (_coa_3yr and show_med) else None

            pred_items = [
                html.Div(
                    source_label if use_knn else f"ABA 509 grant data ({aba_pct:.0f}% of students receive aid)" if aba_pct else "ABA 509 grant data",
                    style={"fontSize": "10px", "color": "#444", "marginBottom": "8px", "fontStyle": "italic"}
                ),
                html.Div([
                    html.Div([
                        html.Div("25th", style={"fontSize": "9px", "color": "#555", "textAlign": "center"}),
                        html.Div(fmt_k(show_p25), style={"fontSize": "15px", "color": "#ccc", "textAlign": "center", "fontWeight": "500"}),
                    ], style={"flex": "1"}),
                    html.Div([
                        html.Div("Median", style={"fontSize": "9px", "color": "#555", "textAlign": "center"}),
                        html.Div(fmt_k(show_med), style={"fontSize": "20px", "color": "#c8a96e", "textAlign": "center", "fontWeight": "700"}),
                    ], style={"flex": "1"}),
                    html.Div([
                        html.Div("75th", style={"fontSize": "9px", "color": "#555", "textAlign": "center"}),
                        html.Div(fmt_k(show_p75), style={"fontSize": "15px", "color": "#ccc", "textAlign": "center", "fontWeight": "500"}),
                    ], style={"flex": "1"}),
                ], style={"display": "flex", "gap": "4px", "marginBottom": "8px"}),
            ]

            if not use_knn and _aba_p50:
                df_check = df[df['school_slug'] == selected_slug]
                n_schol = int((df_check['scholarship'] > 0).sum()) if len(df_check) else 0
                if n_schol > 0 and u_lsat and u_gpa:
                    # Calculate how many were nearby
                    df_c = df_check.dropna(subset=['lsat','gpa','scholarship'])
                    df_c = df_c[df_c['scholarship'] > 0].copy()
                    df_c['_dist'] = (((df_c['lsat']-u_lsat)/60)**2 + ((df_c['gpa']-u_gpa)/2)**2)**0.5
                    n_near = int((df_c['_dist'] <= ((7/60)**2+(0.35/2)**2)**0.5).sum())
                    msg = f"⚠ Only {n_near} nearby applicants on lsd.law (need 5+) — using adjusted ABA estimate"
                elif n_schol > 0:
                    msg = f"⚠ Only {n_schol} scholarship reports on lsd.law — using adjusted ABA estimate"
                else:
                    msg = "⚠ No applicant data on lsd.law — showing school-reported ABA medians"
                pred_items.append(html.Div(msg, style={
                    "fontSize": "10px", "color": "#555", "fontStyle": "italic", "marginBottom": "6px"
                }))

            if _tf_3yr and show_med is not None:
                _t_3yr = (_si.get("tuition") or _si.get("tuition_res") or 0) * 3 if _si else 0
                _f_3yr = (_si.get("fees") or _si.get("fees_res") or 0) * 3 if _si else 0
                if not _t_3yr and not _f_3yr:
                    _t_3yr = _tf_3yr  # fallback if no split available
                net_tf  = max(0, _tf_3yr - show_med)
                aid_pct = min(100, show_med / _tf_3yr * 100)
                living_3yr = (_si.get("living") or 0) * 3 if _si else 0
                grand_total = net_tf + living_3yr

                def cost_line(label, value, color="#888", size="12px", bold=False, indent=False):
                    return html.Div([
                        html.Span(label, style={
                            "color": "#555", "fontSize": "11px",
                            "paddingLeft": "12px" if indent else "0",
                            "minWidth": "90px",
                        }),
                        html.Span(fmt_k(value), style={
                            "color": color, "fontSize": size,
                            "fontWeight": "600" if bold else "400",
                        }),
                    ], style={"display": "flex", "alignItems": "center", "marginBottom": "3px"})

                def divider():
                    return html.Div(style={"borderTop": "1px solid #2a2a2a", "margin": "5px 0"})

                pred_items += [
                    html.Div(style={"borderTop": "1px solid #1e1e1e", "margin": "8px 0 6px"}),
                    html.Div("YOUR COST", style={
                        "fontSize": "9px", "color": "#555", "letterSpacing": "0.1em", "marginBottom": "8px"
                    }),
                    cost_line("Tuition (3yr)", _t_3yr if _t_3yr else _tf_3yr),
                    cost_line("Fees (3yr)", _f_3yr if _f_3yr else None) if _f_3yr else None,
                    cost_line("– Aid", -show_med, "#c8a96e"),
                    divider(),
                    # Bar
                    html.Div([
                        html.Div(style={
                            "width": f"{aid_pct:.1f}%",
                            "background": "#c8a96e",
                            "height": "100%",
                            "borderRadius": "4px 0 0 4px" if aid_pct < 100 else "4px",
                        }),
                        html.Div(style={
                            "width": f"{100-aid_pct:.1f}%",
                            "background": "#2a2a2a",
                            "height": "100%",
                            "borderRadius": "0 4px 4px 0" if aid_pct < 100 else "4px",
                        }),
                    ], style={"display": "flex", "height": "6px", "borderRadius": "4px",
                              "overflow": "hidden", "marginBottom": "5px"}),
                    cost_line("Total", net_tf, "#f4c430", "13px", bold=True),
                ]
                if _f_3yr is None:
                    pred_items = [x for x in pred_items if x is not None]

                if living_3yr:
                    pred_items += [
                        divider(),
                        cost_line("+ Living (3yr)", living_3yr, "#888"),
                        divider(),
                        cost_line("Grand Total", grand_total, "#f0e6d0", "14px", bold=True),
                    ]

                # Remove any None items
                pred_items = [x for x in pred_items if x is not None]

            sections.append(box("SCHOLARSHIP PREDICTION", pred_items))
    o = outcomes.get(selected_slug)
    if o:
        def out_row(label, val, color="#ccc"):
            return html.Div([
                html.Span(label, style={"color": "#666", "fontSize": "15px", "minWidth": "110px", "display": "inline-block"}),
                html.Span(str(val), style={"color": color, "fontSize": "15px", "fontWeight": "500"}),
            ], style={"marginBottom": "6px"})

        biglaw_color = "#c8a96e" if o["biglaw_pct"] >= 0.50 else "#ccc"
        clerk_color  = "#27a87c" if o["fed_clerk_pct"] >= 0.10 else "#ccc"

        grads = o["grads"]

        def pct(n):
            return f"  ({n/grads*100:.0f}%)" if grads else ""

        outcomes_items = [
            out_row("Graduates",       o["grads"]),
            html.Div(style={"borderTop": "1px solid #1e1e1e", "margin": "6px 0"}),
            out_row("BigLaw (501+)",   f"{o['biglaw_n']}{pct(o['biglaw_n'])}", biglaw_color),
            out_row("Mid (251–500)",   f"{o['mid_n']}{pct(o['mid_n'])}"),
            out_row("Small (101–250)", f"{o['small_n']}{pct(o['small_n'])}"),
            html.Div(style={"borderTop": "1px solid #1e1e1e", "margin": "6px 0"}),
            out_row("Fed clerkship",   f"{o['fed_clerk_n']}{pct(o['fed_clerk_n'])}", clerk_color),
            out_row("State clerkship", f"{o['state_clerk_n']}{pct(o['state_clerk_n'])}"),
        ]
        sections.append(box("EMPLOYMENT  ·  ABA 2024", outcomes_items))

    # Split sections into two columns
    left_cols  = sections[::2]   # indexes 0, 2, 4 ...
    right_cols = sections[1::2]  # indexes 1, 3, 5 ...

    return title_children, html.Div([
        html.Div(left_cols,  style={"flex": "1", "minWidth": "0"}),
        html.Div(right_cols, style={"flex": "1", "minWidth": "0"}),
    ], style={"display": "flex", "gap": "6px", "alignItems": "flex-start", "width": "100%"})


@app.callback(
    Output('reset-view-store', 'data'),
    Input('reset-view-btn', 'n_clicks'),
    prevent_initial_call=True,
)
def update_reset_store(n):
    return n or 0


@app.callback(
    Output('plot', 'figure'),
    Input('school-dropdown', 'data'),
    Input('user-lsat', 'value'),
    Input('user-gpa', 'value'),
    Input('overlay-toggles', 'value'),
    Input('reset-view-store', 'data'),
    Input('user-state', 'value'),
    Input('lsat-range', 'value'),
    Input('gpa-range', 'value'),
    Input('schol-range', 'value'),
)
def update_graph(selected_slug, user_lsat, user_gpa, overlays, _reset, user_state, lsat_range, gpa_range, schol_range):
    overlays = overlays or []
    if not selected_slug:
        return px.scatter_3d()

    df_school = df[df['school_slug'] == selected_slug].copy()

    if df_school.empty:
        return px.scatter_3d()

    # Fill missing scholarship with 0 so all applicants show as dots
    df_school['scholarship'] = df_school['scholarship'].fillna(0)
    # Drop rows still missing lsat or gpa
    df_school = df_school.dropna(subset=['lsat', 'gpa'])

    # Apply axis range filters
    lsat_min, lsat_max = (lsat_range or [120, 180])
    gpa_min,  gpa_max  = (gpa_range  or [2.0, 4.0])
    schol_min, schol_max = ((s * 1000) for s in (schol_range or [0, 400]))
    df_school = df_school[
        (df_school['lsat'] >= lsat_min) & (df_school['lsat'] <= lsat_max) &
        (df_school['gpa']  >= gpa_min)  & (df_school['gpa']  <= gpa_max)  &
        (df_school['scholarship'] >= schol_min) & (df_school['scholarship'] <= schol_max)
    ]

    # Filter to scholarship recipients only if toggle is on
    if 'hide_no_schol' in overlays:
        df_school = df_school[df_school['scholarship'] > 0]

    if df_school.empty:
        return px.scatter_3d()

    # --- SURFACE VIEW ---
    if 'surface_view' in overlays:
        import numpy as np
        df_surf = df_school.dropna(subset=['lsat', 'gpa', 'scholarship'])
        df_surf = df_surf[df_surf['scholarship'] > 0]
        if len(df_surf) >= 10:
            n = 35
            xi = np.linspace(df_surf['lsat'].min(), df_surf['lsat'].max(), n)
            yi = np.linspace(df_surf['gpa'].min(),  df_surf['gpa'].max(),  n)
            Zi = np.full((n, n), np.nan)
            counts = np.zeros((n, n))

            # Bin each data point into the nearest grid cell
            for _, row in df_surf.iterrows():
                ix = int(np.round((row['lsat'] - xi[0]) / (xi[-1] - xi[0]) * (n-1))) if xi[-1] > xi[0] else 0
                iy = int(np.round((row['gpa']  - yi[0]) / (yi[-1] - yi[0]) * (n-1))) if yi[-1] > yi[0] else 0
                ix, iy = max(0, min(n-1, ix)), max(0, min(n-1, iy))
                if np.isnan(Zi[iy, ix]):
                    Zi[iy, ix] = row['scholarship']
                    counts[iy, ix] = 1
                else:
                    Zi[iy, ix] = (Zi[iy, ix] * counts[iy, ix] + row['scholarship']) / (counts[iy, ix] + 1)
                    counts[iy, ix] += 1

            # For empty cells, use weighted average of nearby cells
            # Cells far from any data (sparse high-end regions) get the local max
            # to avoid artificial dips
            global_max = np.nanmax(Zi)
            for iy in range(n):
                for ix in range(n):
                    if np.isnan(Zi[iy, ix]):
                        neighbors = []
                        weights = []
                        for radius in range(1, 6):
                            for dy in range(-radius, radius+1):
                                for dx in range(-radius, radius+1):
                                    if abs(dy) != radius and abs(dx) != radius:
                                        continue
                                    ny, nx = iy+dy, ix+dx
                                    if 0 <= ny < n and 0 <= nx < n and not np.isnan(Zi[ny, nx]):
                                        dist = (dy**2 + dx**2) ** 0.5
                                        neighbors.append(Zi[ny, nx])
                                        weights.append(1.0 / dist)
                            if neighbors:
                                break
                        if neighbors:
                            # Use max-biased interpolation for cells with no nearby data
                            # This prevents dipping in above-median regions
                            weighted_avg = np.average(neighbors, weights=weights)
                            local_max = max(neighbors)
                            # Blend toward max in high-LSAT/GPA corner
                            lsat_pct = ix / (n-1)
                            gpa_pct  = iy / (n-1)
                            corner_bias = (lsat_pct + gpa_pct) / 2
                            Zi[iy, ix] = weighted_avg * (1 - corner_bias * 0.4) + local_max * (corner_bias * 0.4)
                        else:
                            Zi[iy, ix] = 0

            fig = go.Figure(data=[go.Surface(
                x=xi, y=yi, z=Zi,
                colorscale='Viridis',
                opacity=0.9,
                showscale=True,
                colorbar=dict(
                    title="Aid ($)", tickfont=dict(color='#aaa', size=9),
                    title_font=dict(color='#aaa', size=9), len=0.5, thickness=12,
                ),
            )])
        else:
            fig = go.Figure()

    # --- COLOR MODE (skip in surface view) ---
    m_data = medians.get(selected_slug) or full_medians.get(selected_slug) or {}
    med_lsat = m_data.get('lsat')
    med_gpa  = m_data.get('gpa')

    if 'surface_view' not in overlays:
        if 'color_median' in overlays and med_lsat and med_gpa:
            df_school = df_school.copy()
            df_school['_median_score'] = (
                (df_school['lsat'] >= med_lsat).astype(int) +
                (df_school['gpa']  >= med_gpa).astype(int)
            ).astype(str)
            fig = px.scatter_3d(
                df_school, x='lsat', y='gpa', z='scholarship',
                color='_median_score',
                color_discrete_map={'0': '#e05060', '1': '#f4c430', '2': '#7fff7f'},
                category_orders={'_median_score': ['0', '1', '2']},
                opacity=0.75,
            )
            fig.update_layout(legend=dict(title="vs Medians", itemsizing="constant"))
            name_map = {'0': 'Below both', '1': 'One above', '2': 'Above both'}
            for trace in fig.data:
                if hasattr(trace, 'name') and trace.name in name_map:
                    trace.name = name_map[trace.name]
        else:
            fig = px.scatter_3d(
                df_school, x='lsat', y='gpa', z='scholarship',
                color='scholarship',
                opacity=0.7,
            )

    z_min = df['scholarship'].min()
    z_max = df['scholarship'].max()
    x_min, x_max = df['lsat'].min(), df['lsat'].max()
    y_min, y_max = df['gpa'].min(), df['gpa'].max()

    # --- USER LINE ---
    if 'you' in overlays and user_lsat is not None and user_gpa is not None:
        try:
            _u_lsat = float(user_lsat)
            _u_gpa  = float(user_gpa)
        except (TypeError, ValueError):
            _u_lsat = _u_gpa = None

        if 'color_median' in overlays and med_lsat and med_gpa and _u_lsat and _u_gpa:
            above_lsat = _u_lsat >= med_lsat
            above_gpa  = _u_gpa  >= med_gpa
            if above_lsat and above_gpa:
                you_color = '#7fff7f'
            elif above_lsat or above_gpa:
                you_color = '#f4c430'
            else:
                you_color = '#e05060'
        else:
            you_color = 'white'

        fig.add_scatter3d(
            x=[user_lsat, user_lsat],
            y=[user_gpa, user_gpa],
            z=[z_min, z_max],
            mode='lines+markers',
            line=dict(color=you_color, width=6),
            marker=dict(color=you_color, size=3),
            name='You'
        )

    # --- MEDIAN / PERCENTILE WALLS ---
    m = medians.get(selected_slug) or full_medians.get(selected_slug)
    if m and m.get('lsat') and m.get('gpa'):

        if 'lsat_med' in overlays:
            fig.add_surface(
                x=[[m['lsat'], m['lsat']], [m['lsat'], m['lsat']]],
                y=[[y_min, y_max], [y_min, y_max]],
                z=[[z_min, z_min], [z_max, z_max]],
                opacity=0.35, colorscale=[[0,'gray'],[1,'gray']],
                showscale=False, name='LSAT median'
            )

        if 'gpa_med' in overlays:
            fig.add_surface(
                x=[[x_min, x_max], [x_min, x_max]],
                y=[[m['gpa'], m['gpa']], [m['gpa'], m['gpa']]],
                z=[[z_min, z_min], [z_max, z_max]],
                opacity=0.35, colorscale=[[0,'gray'],[1,'gray']],
                showscale=False, name='GPA median'
            )

        if 'lsat_pct' in overlays and 'lsat25' in m:
            for lsat_val, label in [(m['lsat25'], 'LSAT 25th'), (m['lsat75'], 'LSAT 75th')]:
                fig.add_surface(
                    x=[[lsat_val, lsat_val], [lsat_val, lsat_val]],
                    y=[[y_min, y_max], [y_min, y_max]],
                    z=[[z_min, z_min], [z_max, z_max]],
                    opacity=0.18, colorscale=[[0,'#3a7fbf'],[1,'#3a7fbf']],
                    showscale=False, name=label
                )

        if 'gpa_pct' in overlays and 'gpa25' in m:
            for gpa_val, label in [(m['gpa25'], 'GPA 25th'), (m['gpa75'], 'GPA 75th')]:
                fig.add_surface(
                    x=[[x_min, x_max], [x_min, x_max]],
                    y=[[gpa_val, gpa_val], [gpa_val, gpa_val]],
                    z=[[z_min, z_min], [z_max, z_max]],
                    opacity=0.18, colorscale=[[0,'#27a87c'],[1,'#27a87c']],
                    showscale=False, name=label
                )

    # --- COA PLANE ---
    if 'coa' in overlays and selected_slug in school_info:
        si = school_info.get(selected_slug, {})
        cost = ((si.get("tuition") or si.get("tuition_res") or (si.get("credit_oos") or si.get("credit_res") or 0) * 31) + (si.get("fees") or si.get("fees_res") or 0)) * 3
        clamped_cost = max(z_min, min(cost, z_max))
        fig.add_surface(
            x=[[x_min, x_max], [x_min, x_max]],
            y=[[y_min, y_min], [y_max, y_max]],
            z=[[clamped_cost, clamped_cost], [clamped_cost, clamped_cost]],
            opacity=0.12, colorscale=[[0,'yellow'],[1,'yellow']],
            showscale=False, name='Tuition+Fees (3yr)'
        )

    if 'half_coa' in overlays and selected_slug in school_info:
        si = school_info.get(selected_slug, {})
        half_cost = ((si.get("tuition") or si.get("tuition_res") or (si.get("credit_oos") or si.get("credit_res") or 0) * 31) + (si.get("fees") or si.get("fees_res") or 0)) * 3 * 0.5
        clamped_half = max(z_min, min(half_cost, z_max))
        fig.add_surface(
            x=[[x_min, x_max], [x_min, x_max]],
            y=[[y_min, y_min], [y_max, y_max]],
            z=[[clamped_half, clamped_half], [clamped_half, clamped_half]],
            opacity=0.15, colorscale=[[0,'#00ccff'],[1,'#00ccff']],
            showscale=False, name='50% Tuition+Fees (3yr)'
        )

    if 'is_coa' in overlays and selected_slug in school_info:
        si = school_info.get(selected_slug, {})
        t_is = si.get("tuition_res") or (si.get("credit_res") or 0) * 31
        f_is = si.get("fees_res") or si.get("fees") or 0
        if t_is:
            is_cost = (t_is + f_is) * 3
            clamped_is = max(z_min, min(is_cost, z_max))
            fig.add_surface(
                x=[[x_min, x_max], [x_min, x_max]],
                y=[[y_min, y_min], [y_max, y_max]],
                z=[[clamped_is, clamped_is], [clamped_is, clamped_is]],
                opacity=0.15, colorscale=[[0,'#7fff7f'],[1,'#7fff7f']],
                showscale=False, name='IS T+F (3yr)'
            )

    # --- PREDICTED SCHOLARSHIP DOT ---
    if user_lsat is not None and user_gpa is not None:
        dot_pred, dot_label, _, _dot_tier = get_prediction(selected_slug, user_lsat, user_gpa, user_state)
        if dot_pred is not None:
            clamped_pred = max(z_min, min(dot_pred, z_max))
            # Color dot based on median position
            if med_lsat and med_gpa:
                try:
                    above_lsat = float(user_lsat) >= med_lsat
                    above_gpa  = float(user_gpa)  >= med_gpa
                except: above_lsat = above_gpa = False
                if above_lsat and above_gpa:     dot_color = '#7fff7f'
                elif above_lsat or above_gpa:    dot_color = '#f4c430'
                else:                             dot_color = '#e05060'
            else:
                dot_color = '#ff69b4'
            fig.add_scatter3d(
                x=[user_lsat],
                y=[user_gpa],
                z=[clamped_pred],
                mode='markers',
                marker=dict(size=10, color=dot_color, symbol='diamond', opacity=1.0,
                            line=dict(color='white', width=1)),
                name=f'~${dot_pred:,.0f} predicted'
            )

    fig.update_layout(
        paper_bgcolor="#000000",
        plot_bgcolor="#000000",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.05,
            xanchor="center",
            x=0.5,
            font=dict(size=10, color="#aaa"),
            bgcolor="rgba(0,0,0,0)",
        ),
        scene=dict(
            aspectmode="cube",
            camera=dict(eye=dict(x=1.8, y=1.8, z=1.2)),
            bgcolor="#000000",
            xaxis=dict(backgroundcolor="#000000", gridcolor="#222", zerolinecolor="#222",
                       range=[lsat_min, lsat_max]),
            yaxis=dict(backgroundcolor="#000000", gridcolor="#222", zerolinecolor="#222",
                       range=[gpa_min, gpa_max]),
            zaxis=dict(backgroundcolor="#000000", gridcolor="#222", zerolinecolor="#222",
                       range=[schol_min, schol_max]),
        )
    )

    return fig


# --- RUN ---
if __name__ == '__main__':
    try:
        app.run(debug=True)
    except Exception as e:
        import traceback
        print("\n" + "="*60)
        print("CRASH:")
        print("="*60)
        traceback.print_exc()
        input("\nPress Enter to close.")
