rm(list=ls())
library(tidyverse)
library(readxl)
library(writexl)
library(lubridate)
library(janitor)
setwd("C:/Users/enric/OneDrive/Desktop/NewJob/DataCleanProject_excelRbi")

# ----- LOAD RAW DATA -----------------------------------------
raw <- read_excel("rawSales.xlsx",
  sheet = "raw_sales_data", col_types = "text")
cat(sprintf("Raw rows loaded: %d\n", nrow(raw)))
head(raw)

# ----- REMOVE BLANK ROWS -------------------------------------
df <- raw %>% filter(!is.na(`Order ID`) & `Order ID` != "")
cat(sprintf("Rows after removing blanks: %d\n", nrow(df)))

# ----- REMOVE DUPLICATES -------------------------------------
df <- df %>% distinct(`Order ID`, .keep_all = TRUE)
cat(sprintf("   Rows after deduplication: %d\n", nrow(df)))

# ----- STANDARDISE COLUMN NAMES ------------------------------
df <- df %>% clean_names() #everything snake_case
# Rename to clean final names
df <- df %>%
  rename(
    order_id    = order_id,
    order_date  = order_date,
    customer    = customer_name,
    email       = customer_email,
    product     = product,
    category    = category,
    quantity    = quantity,
    unit_price  = unit_price,
    total_sale  = tot_sale,
    region      = region,
    sales_rep   = sales_rep
  ) %>%
  select(-notes)  # drop Notes column — not needed for analysis

# ----- CLEAN TEXT FIELDS -------------------------------------
title_case_trim <- function(x) str_to_title(str_squish(x))

df <- df %>%
  mutate(
    customer  = title_case_trim(customer),
    email     = str_to_lower(str_squish(email)),
    product   = title_case_trim(product),
    region    = title_case_trim(region),
    sales_rep = title_case_trim(sales_rep)
  )

# ----- FIX CATEGORY TYPOS ------------------------------
df %>% select(category) %>% table()
df <- df %>%
  mutate(
    category = str_to_title(str_squish(category)),
    category = case_when(
      str_detect(category, regex("electronisc", ignore_case = TRUE)) ~ "Electronics",
      str_detect(category, regex("^furnitures?$", ignore_case = TRUE)) ~ "Furniture",
      str_detect(category, regex("sofware", ignore_case = TRUE)) ~ "Software",
      TRUE ~ category
    )
  )

print(sort(unique(df$category)))

# ----- STANDARDISE DATES --------------------------------
expand_month <- function(x) {
  abbr <- c("Jan","Feb","Mar","Apr","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
  full <- c("January","February","March","April","June","July","August","September","October","November","December")
  for (i in seq_along(abbr)) {
    x <- gsub(paste0("^", abbr[i], " "), paste0(full[i], " "), x)
  }
  return(x)
}

parse_messy_date <- function(x) {
  if (is.na(x) || trimws(x) == "") return(as.Date(NA))
  x <- trimws(x)
  # month as text: "May 09 2024", "January 8, 2024", ecc.
  if (grepl("[A-Za-z]", x)) {
    x <- expand_month(x)
    result <- suppressWarnings(mdy(x))
    if (!is.na(result)) return(result)
    result <- suppressWarnings(dmy(x))
    if (!is.na(result)) return(result)
    return(as.Date(NA))
  }
  # ISO: starts with 4 digit - 2024-..., 2024/...
  if (grepl("^\\d{4}", x)) {
    result <- suppressWarnings(ymd(x))
    if (!is.na(result)) return(result)
    return(as.Date(NA))
  }
  # Slash: second block > 12 -> US (m/d/Y)
  if (grepl("/", x)) {
    parts <- strsplit(x, "/")[[1]]
    if (length(parts) == 3 && as.integer(parts[2]) > 12) {
      result <- suppressWarnings(mdy(x))
      if (!is.na(result)) return(result)
    }
    # otherwise EU, then US
    result <- suppressWarnings(dmy(x))
    if (!is.na(result)) return(result)
    result <- suppressWarnings(mdy(x))
    if (!is.na(result)) return(result)
    return(as.Date(NA))
  }
  # with "-": same
  if (grepl("-", x)) {
    parts <- strsplit(x, "-")[[1]]
    if (length(parts) == 3 && as.integer(parts[2]) > 12) {
      result <- suppressWarnings(mdy(x))
      if (!is.na(result)) return(result)
    }
    result <- suppressWarnings(dmy(x))
    if (!is.na(result)) return(result)
    result <- suppressWarnings(mdy(x))
    if (!is.na(result)) return(result)
    return(as.Date(NA))
  }
  return(as.Date(NA))
}

df <- df %>% mutate(order_date = as.Date(sapply(order_date, parse_messy_date), origin = "1970-01-01"))

cat(sprintf("   Dates parsed successfully: %d / %d\n",
            sum(!is.na(df$order_date)), nrow(df)))

# ----- CLEAN NUMERIC FIELDS ----------------------------------
strip_to_number <- function(x) {
  x %>%
    str_remove_all("[$£€]") %>%         # remove currency symbols
    str_remove_all("(?i)usd|eur|gbp") %>%  # remove currency codes
    str_remove_all("[^0-9.]") %>%        # keep only digits and dot
    str_squish() %>%
    as.numeric()
}

df <- df %>%
  mutate(
    quantity   = as.integer(quantity),
    unit_price = strip_to_number(unit_price),
    total_sale = strip_to_number(total_sale),
    total_sale = if_else(is.na(total_sale), quantity * unit_price, total_sale)
  )

# ----- FINAL VALIDATION --------------------------------------
issues <- list(
  missing_dates  = sum(is.na(df$order_date)),
  missing_price  = sum(is.na(df$unit_price)),
  missing_total  = sum(is.na(df$total_sale)),
  negative_qty   = sum(df$quantity <= 0, na.rm = TRUE)
)

cat(sprintf("   Missing dates:      %d\n", issues$missing_dates))
cat(sprintf("   Missing unit price: %d\n", issues$missing_price))
cat(sprintf("   Missing totals:     %d\n", issues$missing_total))
cat(sprintf("   Invalid quantities: %d\n", issues$negative_qty))

# ----- FINAL COLUMN ORDER & TYPES ----------------------------
df <- df %>%
  select(order_id, order_date, customer, email,
         product, category, quantity, unit_price, total_sale,
         region, sales_rep) %>%
  arrange(order_date)

# ----- EXPORT ------------------------------------------------
write_csv(df,  "cleaned_sales_data_R.csv")
write_xlsx(df, "cleaned_sales_data_R.xlsx")

# ----- QUICK SUMMARY -----------------------------------------

cat("── Summary stats:\n")
cat(sprintf("   Total revenue (cleaned): $%s\n",
            format(sum(df$total_sale, na.rm = TRUE), big.mark = ",", nsmall = 2)))
cat(sprintf("   Date range: %s → %s\n",
            min(df$order_date, na.rm = TRUE),
            max(df$order_date, na.rm = TRUE)))
cat(sprintf("   Unique customers: %d\n", n_distinct(df$customer)))
cat(sprintf("   Unique products:  %d\n", n_distinct(df$product)))
