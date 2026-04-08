import re
import pandas as pd

INPUT_FILE  = r"C:\Users\enric\OneDrive\Desktop\NewJob\DataCleanProject_excelRbi\rawSales.xlsx"
OUTPUT_XLSX = r"C:\Users\enric\OneDrive\Desktop\NewJob\DataCleanProject_excelRbi\cleaned_sales_data_python.xlsx"
OUTPUT_CSV  = r"C:\Users\enric\OneDrive\Desktop\NewJob\DataCleanProject_excelRbi\cleaned_sales_data_python.csv"

# ── 0. LOAD RAW DATA ─────────────────────────────────────────
print("── Loading raw data...")

df = pd.read_excel(INPUT_FILE, sheet_name="raw_sales_data", dtype=str)
df.columns = df.columns.str.strip()   # rimuove spazi dai nomi colonna
print(f"   Raw rows loaded: {len(df)}")

# ── 1. REMOVE BLANK ROWS ─────────────────────────────────────
print("── Step 1: Removing blank rows...")

df = df.dropna(subset=["Order ID"])
df = df[df["Order ID"].str.strip() != ""]
print(f"   Rows after removing blanks: {len(df)}")

# ── 2. REMOVE DUPLICATES ─────────────────────────────────────
print("── Step 2: Removing duplicate Order IDs...")

df = df.drop_duplicates(subset=["Order ID"], keep="first")
print(f"   Rows after deduplication: {len(df)}")

# ── 3. RENAME COLUMNS ────────────────────────────────────────
df = df.rename(columns={
    "Order ID":        "order_id",
    "order date":      "order_date",
    "CUSTOMER_NAME":   "customer",
    "CUSTOMER_Email":  "email",
    "product":         "product",
    "Category":        "category",
    "Quantity":        "quantity",
    "Unit Price":      "unit_price",
    "Tot Sale":        "total_sale",
    "Region":          "region",
    "Sales Rep":       "sales_rep",
})
df = df.drop(columns=["  Notes  "], errors="ignore")

# ── 4. CLEAN TEXT FIELDS ─────────────────────────────────────
print("── Step 4: Cleaning text fields...")

def title_clean(s):
    if pd.isna(s):
        return s
    return " ".join(s.strip().split()).title()   # strip + squeeze spaces + title case

df["customer"]  = df["customer"].apply(title_clean)
df["email"]     = df["email"].str.strip().str.lower()
df["product"]   = df["product"].apply(title_clean)
df["region"]    = df["region"].apply(title_clean)
df["sales_rep"] = df["sales_rep"].apply(title_clean)

# ── 5. FIX CATEGORY TYPOS ────────────────────────────────────
print("── Step 5: Fixing category typos...")

CATEGORY_FIXES = {
    r"electronisc":    "Electronics",
    r"furnitures?":    "Furniture",
    r"sofware":        "Software",
    r"office\s+supplies": "Office Supplies",
}

def fix_category(s):
    if pd.isna(s):
        return s
    s = s.strip()
    for pattern, replacement in CATEGORY_FIXES.items():
        if re.fullmatch(pattern, s, flags=re.IGNORECASE):
            return replacement
    return s.title()

df["category"] = df["category"].apply(fix_category)
print(f"   Categories after fix: {sorted(df['category'].unique())}")

# ── 6. STANDARDISE DATES ─────────────────────────────────────
print("── Step 6: Parsing dates...")

MONTH_ABBR = {
    "Jan": "January", "Feb": "February", "Mar": "March",
    "Apr": "April",   "Jun": "June",     "Jul": "July",
    "Aug": "August",  "Sep": "September","Oct": "October",
    "Nov": "November","Dec": "December"
}

def expand_month(s):
    for abbr, full in MONTH_ABBR.items():
        if s.startswith(abbr + " "):
            return full + s[len(abbr):]
    return s

def parse_date(s):
    if pd.isna(s) or str(s).strip() == "":
        return pd.NaT
    s = str(s).strip()

    # Mese testuale
    if re.search(r"[A-Za-z]", s):
        s = expand_month(s)
        for fmt in ("%B %d, %Y", "%B %d %Y", "%B %Y"):
            try:
                return pd.to_datetime(s, format=fmt)
            except ValueError:
                continue
        return pd.NaT

    # ISO: inizia con 4 cifre
    if re.match(r"^\d{4}", s):
        for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
            try:
                return pd.to_datetime(s, format=fmt)
            except ValueError:
                continue
        return pd.NaT

    # Slash: se secondo blocco > 12 è formato US (m/d/Y)
    if "/" in s:
        parts = s.split("/")
        if len(parts) == 3 and int(parts[1]) > 12:
            try:
                return pd.to_datetime(s, format="%m/%d/%Y")
            except ValueError:
                pass
        for fmt in ("%d/%m/%Y", "%m/%d/%Y"):
            try:
                return pd.to_datetime(s, format=fmt)
            except ValueError:
                continue
        return pd.NaT

    # Trattino: stessa logica
    if "-" in s:
        parts = s.split("-")
        if len(parts) == 3 and int(parts[1]) > 12:
            try:
                return pd.to_datetime(s, format="%m-%d-%Y")
            except ValueError:
                pass
        for fmt in ("%d-%m-%Y", "%m-%d-%Y"):
            try:
                return pd.to_datetime(s, format=fmt)
            except ValueError:
                continue
        return pd.NaT

    return pd.NaT

df["order_date"] = df["order_date"].apply(parse_date)
parsed_ok = df["order_date"].notna().sum()
print(f"   Dates parsed successfully: {parsed_ok} / {len(df)}")

# ── 7. CLEAN NUMERIC FIELDS ──────────────────────────────────
print("── Step 7: Cleaning Unit Price and Total Sale...")

def strip_to_number(s):
    if pd.isna(s) or str(s).strip() in ("", "N/A"):
        return None
    s = str(s)
    s = re.sub(r"[$£€]", "", s)
    s = re.sub(r"(?i)usd|eur|gbp", "", s)
    s = re.sub(r"[^\d.]", "", s)
    try:
        return float(s)
    except ValueError:
        return None

df["quantity"]   = pd.to_numeric(df["quantity"], errors="coerce").astype("Int64")
df["unit_price"] = df["unit_price"].apply(strip_to_number)
df["total_sale"] = df["total_sale"].apply(strip_to_number)

# Fallback: ricalcola total_sale se mancante
mask = df["total_sale"].isna() & df["unit_price"].notna() & df["quantity"].notna()
df.loc[mask, "total_sale"] = df.loc[mask, "unit_price"] * df.loc[mask, "quantity"]

# ── 8. FINAL VALIDATION ──────────────────────────────────────
print("── Step 8: Validation...")
print(f"   Missing dates:      {df['order_date'].isna().sum()}")
print(f"   Missing unit price: {df['unit_price'].isna().sum()}")
print(f"   Missing totals:     {df['total_sale'].isna().sum()}")

# ── 9. FINAL COLUMN ORDER ────────────────────────────────────
df = df[["order_id", "order_date", "customer", "email",
         "product", "category", "quantity", "unit_price",
         "total_sale", "region", "sales_rep"]]

df = df.sort_values("order_date").reset_index(drop=True)

# ── 10. EXPORT ───────────────────────────────────────────────
print("── Exporting...")
df.to_csv(OUTPUT_CSV, index=False)
df.to_excel(OUTPUT_XLSX, index=False, sheet_name="cleaned_sales_data")

print(f"\n✓ Done! Cleaned dataset: {len(df)} rows × {len(df.columns)} columns")
print(f"  Output files:")
print(f"    → {OUTPUT_CSV}")
print(f"    → {OUTPUT_XLSX}")

# ── QUICK SUMMARY ────────────────────────────────────────────
print("\n── Summary stats:")
print(f"   Total revenue (cleaned): ${df['total_sale'].sum():,.2f}")
print(f"   Date range: {df['order_date'].min().date()} → {df['order_date'].max().date()}")
print(f"   Unique customers: {df['customer'].nunique()}")
print(f"   Unique products:  {df['product'].nunique()}")