---
title: "WIIH_2026_DOCUMENTATION_RZM"
author: "Roger Zamora"
date: "2026-10-19"
output:
  word_document: default
  pdf_document: default
  html_document: default
---

```{r setup, include=FALSE}
knitr::opts_chunk$set(echo = TRUE)
```


# I loaded the required packages, defined the input and output paths and publication details, and imported the original database without modifying it.

```{r setup_and_import, message=FALSE}

# ============================================================
# WIIH 2026 — MASTER ECONOMY PROFILES
#
# Final output:
# One Excel workbook with an Index and 29 economy worksheets.
# Each profile will contain:
# - Profile settings and display instructions
# - Four key indicator cards
# - Seven automatically generated key insights
# - Country-specific and cross-country figure data
# - Titles, subtitles, colours, sources and notes
#
# CHUNK 1: SETUP AND SOURCE DATA IMPORT
# Applies to: all economies.
# ============================================================


# ------------------------------------------------------------
# 1. Load packages used by this workflow
# ------------------------------------------------------------

library(readxl)    # Read the source Excel files
library(dplyr)     # Prepare data and calculate indicators
library(tidyr)     # Reshape asset and time-series data
library(purrr)     # Build profiles across the selected economies
library(stringr)   # Construct and format profile text
library(tibble)    # Create structured profile tables
library(openxlsx)  # Format and export the final Excel workbook


# ------------------------------------------------------------
# 2. Define the project location and file paths
# ------------------------------------------------------------

project_dir <- "C:/Users/zamoramatu/OneDrive - WIPO/Documents"

input_file <- file.path(
  project_dir,
  "intangible_mock2026.xlsx"
)

output_file <- file.path(
  project_dir,
  "WIIH_2026_Intangible_Economy_Profiles_MASTER_FINAL.xlsx"
)


# ------------------------------------------------------------
# 3. Define the publication information
# Used later in worksheet titles, profile settings and sources.
# ------------------------------------------------------------

report_year <- 2026L

report_title <- paste(
  "World Intangible Investment Highlights",
  report_year
)

source_text <- paste(
  "Source: WIPO-LBS Global INTAN-Invest Database,",
  "July 2026."
)


# ------------------------------------------------------------
# 4. Import the original database
# Check the path first so a missing file produces a clear error.
# ------------------------------------------------------------

if (!file.exists(input_file)) {
  stop(
    "Source database not found: ",
    input_file,
    call. = FALSE
  )
}

data_intan_raw <- readxl::read_excel(
  path = input_file,
  sheet = 1
)


# ------------------------------------------------------------
# 5. Inspect the imported structure
#
# Keep data_intan_raw unchanged throughout the workflow.
# The next chunk will create the working dataset and establish
# economy identifiers and the required country-specific rules.
# ------------------------------------------------------------

dim(data_intan_raw)
names(data_intan_raw)

```



# I defined the 29 economies, their names, ordering and calculation exceptions. I standardized identifiers, applied Japan’s code correction, and added checks for valid years, economy coverage and duplicate observations.
```{r economies_and_specificities, message=FALSE}

# ============================================================
# CHUNK 2: ECONOMY DEFINITIONS AND COUNTRY-SPECIFIC RULES
#
# Purpose:
# Establish one consistent set of identifiers, names and
# exceptions for the entire master workbook.
#
# No indicators or key insights are calculated in this chunk.
# ============================================================


# ------------------------------------------------------------
# 1. Define the 29 economies in the reference workbook order
#
# Keep the database's codes:
# EL = Greece; UK = United Kingdom;
# IND = India; PHL = Philippines.
#
# EU-22 and other aggregates are outside the profile scope.
# ------------------------------------------------------------

economy_lookup <- tibble::tribble(
  ~geo_code, ~economy_name,
  "BR",     "Brazil",
  "BG",     "Bulgaria",
  "CA",     "Canada",
  "HR",     "Croatia",
  "CZ",     "Czech Republic",
  "DK",     "Denmark",
  "EE",     "Estonia",
  "FI",     "Finland",
  "FR",     "France",
  "DE",     "Germany",
  "EL",     "Greece",
  "HU",     "Hungary",
  "IND",    "India",
  "IT",     "Italy",
  "JP",     "Japan",
  "LV",     "Latvia",
  "LT",     "Lithuania",
  "LU",     "Luxembourg",
  "NL",     "Netherlands",
  "PHL",    "Philippines",
  "PL",     "Poland",
  "PT",     "Portugal",
  "RO",     "Romania",
  "SK",     "Slovakia",
  "SI",     "Slovenia",
  "ES",     "Spain",
  "SE",     "Sweden",
  "UK",     "United Kingdom",
  "US",     "United States of America"
)


# ------------------------------------------------------------
# 2. Define names for different parts of the profiles
#
# ALL ECONOMIES:
# economy_name is used for worksheet names and profile headers.
#
# US / UK:
# Use "the US" and "the UK" within key insights.
#
# CZ / NL / PHL:
# Include "the" in prose and use the appropriate possessive.
#
# Keep separate sentence-opening possessives so that:
# "the Netherlands'" becomes "The Netherlands'".
# ------------------------------------------------------------

economy_lookup <- economy_lookup %>%
  mutate(
    display_order = row_number(),

    prose_name = case_when(
      geo_code == "US"  ~ "the US",
      geo_code == "UK"  ~ "the UK",
      geo_code == "CZ"  ~ "the Czech Republic",
      geo_code == "NL"  ~ "the Netherlands",
      geo_code == "PHL" ~ "the Philippines",
      TRUE             ~ economy_name
    ),

    possessive_name = case_when(
      geo_code %in% c("NL", "PHL") ~ paste0(prose_name, "\u2019"),
      TRUE                        ~ paste0(prose_name, "\u2019s")
    ),

    possessive_sentence = stringr::str_replace(
      possessive_name,
      "^the ",
      "The "
    ),

    # Match the figure titles in the reference workbook:
    # US uses "(US)"; UK uses "(United Kingdom)".
    figure_name = if_else(
      geo_code == "US",
      "US",
      economy_name
    ),

    # The separate profile-settings label uses US / UK.
    figure_label = paste0(
      "(",
      if_else(geo_code %in% c("US", "UK"), geo_code, economy_name),
      ")"
    )
  )

geo_list <- economy_lookup$geo_code


# ------------------------------------------------------------
# 3. Record calculation and presentation exceptions
#
# NA year settings mean:
# Select the appropriate available period from the data.
# They do not mean that the indicator itself is unavailable.
# ------------------------------------------------------------

economy_rules <- economy_lookup %>%
  transmute(
    geo_code,

    # CANADA:
    # Use GDP_Adj because VAadj is unavailable.
    #
    # INDIA:
    # Use the separate formal-sector intensity file for:
    # Figure 1, Figure 9, KPI card 3 and key insight 4.
    #
    # OTHER ECONOMIES:
    # Use the adjusted value-added denominator VAadj.
    intensity_source = case_when(
      geo_code == "CA"  ~ "GDP_Adj",
      geo_code == "IND" ~ "formal_sector_file",
      TRUE             ~ "VAadj"
    ),

    # BRAZIL:
    # Figures 2 and 3 use the comparable 2011–2021 period.
    # Its headline intangible CAGR and global comparison will
    # still use the latest available 10-year intangible window.
    comparison_start_year = if_else(
      geo_code == "BR", 2011L, NA_integer_
    ),
    comparison_end_year = if_else(
      geo_code == "BR", 2021L, NA_integer_
    ),

    # CANADA / NETHERLANDS:
    # Preserve the 2013–2023 real asset-growth window used in
    # this edition. Other economies use their qualifying data.
    asset_growth_start_year = if_else(
      geo_code %in% c("CA", "NL"), 2013L, NA_integer_
    ),
    asset_growth_end_year = if_else(
      geo_code %in% c("CA", "NL"), 2023L, NA_integer_
    ),

    # CANADA:
    # Use 2023 for nominal asset composition, matching the
    # reference profile and the earlier production exception.
    composition_year = if_else(
      geo_code == "CA", 2023L, NA_integer_
    ),

    # INDIA / NETHERLANDS / POLAND:
    # Exclude Other IP products from the asset-growth chart
    # and fastest-growing-asset ranking, matching the master.
    #
    # This exclusion does NOT apply to asset composition:
    # Figure 5 retains all seven categories when available.
    exclude_oipp_growth = geo_code %in% c("IND", "NL", "PL")
  )

# INDIA:
# Store the supplemental file path here.
# Its contents will be imported and checked in the intensity
# chunk; no generic-denominator fallback will be assumed.
india_intensity_file <- file.path(
  project_dir,
  "IND_2026_formal_IA&TA_intensity.xlsx"
)


# ------------------------------------------------------------
# 4. Create the working dataset
#
# Preserve data_intan_raw as the untouched source.
# Keep the source country and region columns unchanged.
# ------------------------------------------------------------

required_id_columns <- c("geo_code", "country", "region", "year")

missing_id_columns <- setdiff(
  required_id_columns,
  names(data_intan_raw)
)

if (length(missing_id_columns) > 0L) {
  stop(
    "Required source columns are missing: ",
    paste(missing_id_columns, collapse = ", "),
    call. = FALSE
  )
}

data_intan_inflow <- data_intan_raw %>%
  mutate(
    geo_code = na_if(
      toupper(trimws(as.character(geo_code))),
      ""
    ),

    # JAPAN:
    # The supplied original code identifies Japan as country 18.
    # Apply its JP correction once so all indicators use it.
    geo_code = case_when(
      trimws(as.character(country)) == "18" ~ "JP",
      TRUE                                ~ geo_code
    )
  ) %>%
  filter(geo_code %in% geo_list)


# ------------------------------------------------------------
# 5. Validate years without silently changing invalid entries
#
# ALL SELECTED ECONOMIES:
# Missing years remain NA; nonnumeric or fractional years stop
# the workflow before they can affect sorting or growth periods.
# ------------------------------------------------------------

year_text <- trimws(as.character(data_intan_inflow$year))
year_text[!is.na(year_text) & year_text == ""] <- NA_character_

year_numeric <- suppressWarnings(as.numeric(year_text))

invalid_year <- !is.na(year_text) & (
  is.na(year_numeric) |
    !is.finite(year_numeric) |
    year_numeric != trunc(year_numeric)
)

if (any(invalid_year)) {
  stop(
    "Invalid year entries: ",
    paste(unique(year_text[invalid_year]), collapse = ", "),
    call. = FALSE
  )
}

data_intan_inflow$year <- year_numeric


# ------------------------------------------------------------
# 6. Check economy coverage and observation uniqueness
#
# ALL SELECTED ECONOMIES:
# The output requires 29 profiles.
# Each economy-year must identify one source observation.
# ------------------------------------------------------------

missing_economies <- setdiff(
  geo_list,
  unique(data_intan_inflow$geo_code)
)

if (length(missing_economies) > 0L) {
  stop(
    "Expected economy codes are missing: ",
    paste(missing_economies, collapse = ", "),
    ". Check the source identifiers.",
    call. = FALSE
  )
}

duplicate_country_years <- data_intan_inflow %>%
  filter(!is.na(year)) %>%
  count(geo_code, year, name = "n_rows") %>%
  filter(n_rows > 1L)

if (nrow(duplicate_country_years) > 0L) {
  print(duplicate_country_years)

  stop(
    "Multiple source observations exist for an economy-year. ",
    "Resolve these before calculating indicators.",
    call. = FALSE
  )
}

data_intan_inflow <- data_intan_inflow %>%
  arrange(match(geo_code, geo_list), year)

# Display the economy registry for inspection.
economy_lookup %>%
  select(geo_code, economy_name, prose_name, possessive_name)

```



# I selected the latest available intangible and tangible investment levels at current PPP prices, retaining each series’ actual observation year. I also defined the numeric-validation helper used in subsequent calculations.
```{r investment_levels, message=FALSE}

# ============================================================
# CHUNK 3 — LATEST INVESTMENT LEVELS
#
# Inputs:
#   I_Intang_PPP  : intangible investment at current PPP prices
#   I_TangNRes_PPP: tangible investment at current PPP prices
#
# Output:
#   investment_levels — one row per economy, with the latest
#   available value and its actual observation year.
# ============================================================


# ------------------------------------------------------------
# 1. Reusable helper for reading numeric source columns
#
# Applies to every economy.
# Accepts numeric cells and numeric text.
# Preserves missing values; never replaces them with zero.
# Stops on ambiguous or invalid entries instead of silently
# introducing missing values into calculations or insights.
# ------------------------------------------------------------

to_numeric_checked <- function(x, column_name) {

  if (inherits(x, c("Date", "POSIXt", "difftime"))) {
    stop(
      "Expected investment values, but found dates/times in: ",
      column_name,
      call. = FALSE
    )
  }

  if (is.numeric(x)) {

    result <- as.double(x)

  } else if (is.logical(x) && all(is.na(x))) {

    # readxl can import a completely empty column as logical.
    result <- rep(NA_real_, length(x))

  } else if (is.character(x)) {

    text_values <- trimws(x)
    text_values[!is.na(text_values) & text_values == ""] <- NA_character_

    result <- suppressWarnings(as.numeric(text_values))

    invalid <- !is.na(text_values) & is.na(result)

    if (any(invalid)) {
      stop(
        "Non-numeric entries in ", column_name, ": ",
        paste(
          head(unique(text_values[invalid]), 5L),
          collapse = ", "
        ),
        ". Check these source cells.",
        call. = FALSE
      )
    }

  } else {

    stop(
      "Unsupported data type in: ", column_name,
      call. = FALSE
    )

  }

  if (any(!is.na(result) & !is.finite(result))) {
    stop(
      "Infinite values found in: ", column_name,
      call. = FALSE
    )
  }

  result[is.na(result)] <- NA_real_

  result
}


# ------------------------------------------------------------
# 2. Check and prepare the two investment columns
#
# Only the working dataset is updated.
# data_intan_raw remains unchanged.
# ------------------------------------------------------------

level_variables <- c(
  "I_Intang_PPP",
  "I_TangNRes_PPP"
)

missing_level_columns <- setdiff(
  level_variables,
  names(data_intan_inflow)
)

if (length(missing_level_columns) > 0L) {
  stop(
    "Required investment-level columns are missing: ",
    paste(missing_level_columns, collapse = ", "),
    call. = FALSE
  )
}

data_intan_inflow <- data_intan_inflow %>%
  mutate(
    across(
      all_of(level_variables),
      ~ to_numeric_checked(.x, cur_column())
    )
  )


# ------------------------------------------------------------
# 3. Select the latest available value for each investment type
#
# All economies:
# Select intangible and tangible observations independently.
# Their latest years do not have to match.
#
# Brazil (BR):
# This preserves the distinction when tangible investment
# ends earlier than intangible investment.
#
# Canada (CA), India (IND), Japan (JP), Philippines (PHL):
# Earlier latest years are determined from source availability,
# rather than hard-coded to the latest year of other economies.
#
# Duplicate economy-year rows were checked in Chunk 2.
# ------------------------------------------------------------

latest_levels_long <- data_intan_inflow %>%
  select(
    geo_code,
    year,
    all_of(level_variables)
  ) %>%
  pivot_longer(
    cols = all_of(level_variables),
    names_to = "series",
    values_to = "level_raw"
  ) %>%
  filter(
    !is.na(year),
    !is.na(level_raw)
  ) %>%
  group_by(geo_code, series) %>%
  slice_max(
    order_by = year,
    n = 1,
    with_ties = FALSE
  ) %>%
  ungroup()


# ------------------------------------------------------------
# 4. Build the economy-level table
#
# Starting from economy_lookup retains all 29 profiles.
# An economy without a valid observation keeps NA for that
# value and year instead of disappearing from the output.
#
# United States (US):
# Keep source values unchanged here. The trillion display
# will be handled explicitly in the formatting chunk.
#
# All economies:
# Preserve source units and full precision at this stage.
# No unit conversion is inferred from the size of a value.
# ------------------------------------------------------------

investment_levels <- economy_lookup %>%
  select(
    geo_code,
    economy_name,
    display_order
  ) %>%
  left_join(
    latest_levels_long %>%
      filter(series == "I_Intang_PPP") %>%
      transmute(
        geo_code,
        latest_year_intangible = year,
        intangible_raw = level_raw
      ),
    by = "geo_code"
  ) %>%
  left_join(
    latest_levels_long %>%
      filter(series == "I_TangNRes_PPP") %>%
      transmute(
        geo_code,
        latest_year_tangible = year,
        tangible_raw = level_raw
      ),
    by = "geo_code"
  ) %>%
  arrange(display_order)

investment_levels
```


# I calculated year-on-year growth in constant-price investment using observations exactly one year apart. I selected the latest valid growth rate separately for intangible and tangible investment.
```{r investment_yoy, message=FALSE}

# ============================================================
# CHUNK 4 — LATEST REAL YEAR-ON-YEAR GROWTH
#
# Inputs:
#   Iq_Intang : intangible investment at constant prices
#   Iq_TangNRes: tangible investment at constant prices
#
# Outputs:
#   real_investment_long   — reusable real investment series
#   investment_yoy_history— annual growth calculations
#   investment_yoy        — latest valid growth per economy
#
# Growth remains a decimal: 0.028 means 2.8%.
# Formatting and narrative wording are applied later.
# ============================================================


# ------------------------------------------------------------
# 1. Check and prepare the real investment columns
#
# Applies to every economy.
# Reuses the numeric validation helper from Chunk 3.
# ------------------------------------------------------------

real_investment_variables <- c(
  "Iq_Intang",
  "Iq_TangNRes"
)

missing_real_columns <- setdiff(
  real_investment_variables,
  names(data_intan_inflow)
)

if (length(missing_real_columns) > 0L) {
  stop(
    "Required real investment columns are missing: ",
    paste(missing_real_columns, collapse = ", "),
    call. = FALSE
  )
}

data_intan_inflow <- data_intan_inflow %>%
  mutate(
    across(
      all_of(real_investment_variables),
      ~ to_numeric_checked(.x, cur_column())
    )
  )


# ------------------------------------------------------------
# 2. Prepare a reusable table for real investment calculations
#
# Keep missing investment observations in the timeline.
# This table will also support the CAGR and index chunks.
# ------------------------------------------------------------

real_investment_long <- data_intan_inflow %>%
  select(
    geo_code,
    year,
    all_of(real_investment_variables)
  ) %>%
  pivot_longer(
    cols = all_of(real_investment_variables),
    names_to = "series",
    values_to = "real_value"
  ) %>%
  filter(!is.na(year)) %>%
  arrange(geo_code, series, year)


# ------------------------------------------------------------
# 3. Calculate genuine annual growth
#
# All economies:
# Both observations must exist, and their years must differ
# by exactly one. A gap such as 2021–2023 is not annual growth.
#
# Missing values and zero denominators produce NA.
# Negative growth is retained for contraction wording later.
# A valid zero end value produces -100% growth.
# ------------------------------------------------------------

investment_yoy_history <- real_investment_long %>%
  group_by(geo_code, series) %>%
  arrange(year, .by_group = TRUE) %>%
  mutate(
    start_year = lag(year),
    start_value = lag(real_value),

    yoy_real = if_else(
      !is.na(start_year) &
        year - start_year == 1 &
        !is.na(start_value) &
        start_value != 0 &
        !is.na(real_value),

      (real_value / start_value) - 1,

      NA_real_
    ),

    yoy_real = if_else(
      is.finite(yoy_real),
      yoy_real,
      NA_real_
    )
  ) %>%
  ungroup() %>%
  select(
    geo_code,
    series,
    start_year,
    end_year = year,
    start_value,
    end_value = real_value,
    yoy_real
  )


# ------------------------------------------------------------
# 4. Select the latest valid annual comparison independently
#
# Brazil (BR):
# Tangible and intangible investment can end in different
# years. Preserve both periods for accurate insight wording.
#
# All economies:
# Derive the available periods from the data.
# The latest real-growth year may also differ from the latest
# PPP investment-level year calculated in Chunk 3.
# ------------------------------------------------------------

latest_yoy_long <- investment_yoy_history %>%
  filter(!is.na(yoy_real)) %>%
  group_by(geo_code, series) %>%
  slice_max(
    order_by = end_year,
    n = 1,
    with_ties = FALSE
  ) %>%
  ungroup()


# ------------------------------------------------------------
# 5. Build the economy-level output
#
# Retain all 29 economies.
# Missing growth stays NA; it is never reported as 0%.
# Keep full precision for later rounding and text generation.
# ------------------------------------------------------------

investment_yoy <- economy_lookup %>%
  select(
    geo_code,
    economy_name,
    display_order
  ) %>%
  left_join(
    latest_yoy_long %>%
      filter(series == "Iq_Intang") %>%
      transmute(
        geo_code,
        yoy_start_year_intangible = start_year,
        yoy_end_year_intangible = end_year,
        intangible_yoy_real = yoy_real
      ),
    by = "geo_code"
  ) %>%
  left_join(
    latest_yoy_long %>%
      filter(series == "Iq_TangNRes") %>%
      transmute(
        geo_code,
        yoy_start_year_tangible = start_year,
        yoy_end_year_tangible = end_year,
        tangible_yoy_real = yoy_real
      ),
    by = "geo_code"
  ) %>%
  arrange(display_order)

investment_yoy

```


# I calculated real CAGR using each investment series’ latest observation and its value exactly ten years earlier. I preserved separate periods for intangible and tangible investment and left invalid calculations missing.
```{r investment_cagr, message=FALSE}

# ============================================================
# CHUNK 5 — LATEST 10-YEAR REAL INVESTMENT GROWTH
#
# Input:
#   real_investment_long, prepared in Chunk 4.
#
# Outputs:
#   investment_cagr_long — endpoints and calculations by series
#   investment_cagr      — one row per economy
#
# CAGR = (end_value / start_value)^(1 / 10) - 1
#
# Results remain unrounded decimals.
# ============================================================


# ------------------------------------------------------------
# 1. Identify the latest observation for each real series
#
# All economies:
# Anchor the window to the latest non-missing observation.
# The required starting year is exactly ten years earlier.
#
# Intangible and tangible investment are handled separately,
# because their latest available years can differ.
#
# Canada (CA) and Netherlands (NL):
# Their fixed 2013–2023 asset-growth rule applies to individual
# assets. Aggregate investment growth uses the dates derived
# here from Iq_Intang and Iq_TangNRes.
# ------------------------------------------------------------

latest_real_endpoints <- real_investment_long %>%
  filter(!is.na(real_value)) %>%
  group_by(geo_code, series) %>%
  slice_max(
    order_by = year,
    n = 1,
    with_ties = FALSE
  ) %>%
  ungroup() %>%
  transmute(
    geo_code,
    series,
    start_year = year - 10,
    end_year = year,
    end_value = real_value
  )


# ------------------------------------------------------------
# 2. Retrieve the exact starting observation
#
# Retain both series for all 29 economies.
#
# If the required starting observation is missing, keep NA.
# Do not substitute a nearby year or shorten the period.
#
# Intermediate observations are not required for CAGR:
# the calculation uses the two endpoints.
# ------------------------------------------------------------

investment_cagr_long <- tidyr::expand_grid(
  geo_code = geo_list,
  series = real_investment_variables
) %>%
  left_join(
    latest_real_endpoints,
    by = c("geo_code", "series")
  ) %>%
  left_join(
    real_investment_long %>%
      transmute(
        geo_code,
        series,
        start_year = year,
        start_value = real_value
      ),
    by = c("geo_code", "series", "start_year")
  )


# ------------------------------------------------------------
# 3. Calculate CAGR using valid endpoints
#
# All economies:
# Both endpoints must be strictly positive, following the
# original CAGR rule. Missing or invalid endpoints yield NA.
#
# Mask invalid ratios before taking the fractional power.
# This avoids calculating growth from zero/negative bases.
#
# Negative CAGR remains valid when positive investment
# levels decrease over the decade.
# ------------------------------------------------------------

investment_cagr_long <- investment_cagr_long %>%
  mutate(
    n_years = end_year - start_year,

    valid_endpoints =
      is.finite(start_value) &
      start_value > 0 &
      is.finite(end_value) &
      end_value > 0 &
      !is.na(n_years) &
      n_years == 10,

    endpoint_ratio = if_else(
      valid_endpoints,
      end_value / start_value,
      NA_real_
    ),

    cagr_real = endpoint_ratio^(1 / 10) - 1,

    cagr_real = if_else(
      is.finite(cagr_real),
      cagr_real,
      NA_real_
    )
  ) %>%
  select(
    geo_code,
    series,
    start_year,
    end_year,
    start_value,
    end_value,
    n_years,
    cagr_real
  )


# ------------------------------------------------------------
# 4. Build the economy-level output
#
# Preserve separate intangible and tangible periods so that
# the third key insight can state different windows accurately.
#
# Brazil (BR):
# This table supplies the latest available decade for the
# headline KPI, key insight and global CAGR comparison.
#
# Brazil's fixed 2011–2021 window for the country comparison
# chart (Figure 3) is a separate calculation, to be built
# using the comparison rule already defined in Chunk 2.
# ------------------------------------------------------------

investment_cagr <- economy_lookup %>%
  select(
    geo_code,
    economy_name,
    display_order
  ) %>%
  left_join(
    investment_cagr_long %>%
      filter(series == "Iq_Intang") %>%
      transmute(
        geo_code,
        start_year_intangible = start_year,
        end_year_intangible = end_year,
        intangible_cagr_real = cagr_real
      ),
    by = "geo_code"
  ) %>%
  left_join(
    investment_cagr_long %>%
      filter(series == "Iq_TangNRes") %>%
      transmute(
        geo_code,
        start_year_tangible = start_year,
        end_year_tangible = end_year,
        tangible_cagr_real = cagr_real
      ),
    by = "geo_code"
  ) %>%
  arrange(display_order)

investment_cagr

```


# I constructed investment indices with the selected starting year equal to 100 and calculated growth over the same comparison window for Figures 2–3. I applied Brazil’s 2011–2021 window and preserved missing observations as gaps.
```{r}

# ============================================================
# CHUNK 6 — COMPARISON PERIODS AND INDEXED REAL INVESTMENT
#
# Inputs:
#   real_investment_long — prepared in Chunk 4
#   economy_rules       — prepared in Chunk 2
#
# Outputs:
#   comparison_windows
#   comparison_endpoints
#   indexed_real_series
#   investment_comparison_cagr
#
# Used by:
#   Figure 2 — Indexed real intangible/tangible investment
#   Figure 3 — Real CAGR over the comparison period
#
# All calculations retain full precision.
# Display rounding is applied later.
# ============================================================


# ------------------------------------------------------------
# 1. Check the required upstream objects
# ------------------------------------------------------------

required_comparison_objects <- c(
  "real_investment_long",
  "real_investment_variables",
  "economy_lookup",
  "economy_rules",
  "geo_list"
)

missing_comparison_objects <- required_comparison_objects[
  !vapply(
    required_comparison_objects,
    exists,
    logical(1),
    inherits = TRUE
  )
]

if (length(missing_comparison_objects) > 0L) {
  stop(
    "Run the preceding chunks first. Missing objects: ",
    paste(missing_comparison_objects, collapse = ", "),
    call. = FALSE
  )
}

# A prescribed comparison period must specify both endpoints.
# Brazil currently has this exception: 2011–2021.

if (any(xor(
  is.na(economy_rules$comparison_start_year),
  is.na(economy_rules$comparison_end_year)
))) {
  stop(
    "Each prescribed comparison period must specify both ",
    "comparison_start_year and comparison_end_year.",
    call. = FALSE
  )
}


# ------------------------------------------------------------
# 2. Select one comparison window per economy
#
# General rule:
# Use the latest year with a finite, strictly positive value
# in either real investment series. Start ten years earlier.
# This gives an eleven-calendar-year timeline.
#
# Brazil (BR):
# Use its prescribed 2011–2021 comparison window from Chunk 2.
# This applies to Figures 2 and 3 only.
#
# Canada (CA) and Netherlands (NL):
# Their prescribed asset-growth periods apply to individual
# assets in Figure 4, not to these aggregate series.
#
# All economies:
# Missing observations within the selected window remain
# missing. Do not shift the window separately for each series.
# ------------------------------------------------------------

comparison_latest_years <- real_investment_long %>%
  filter(
    series %in% real_investment_variables,
    is.finite(year),
    is.finite(real_value),
    real_value > 0
  ) %>%
  group_by(geo_code) %>%
  summarise(
    latest_available_year = max(year),
    .groups = "drop"
  )

comparison_windows <- economy_lookup %>%
  select(geo_code) %>%
  left_join(
    economy_rules %>%
      select(
        geo_code,
        comparison_start_year,
        comparison_end_year
      ),
    by = "geo_code"
  ) %>%
  left_join(
    comparison_latest_years,
    by = "geo_code"
  ) %>%
  transmute(
    geo_code,

    start_year = coalesce(
      comparison_start_year,
      latest_available_year - 10
    ),

    end_year = coalesce(
      comparison_end_year,
      latest_available_year
    )
  ) %>%
  arrange(match(geo_code, geo_list))

# Economies without a selectable period retain NA dates.
# Every defined period must span exactly ten years.

invalid_comparison_windows <- comparison_windows %>%
  filter(
    !is.na(start_year) | !is.na(end_year)
  ) %>%
  filter(
    !is.finite(start_year) |
      !is.finite(end_year) |
      end_year - start_year != 10
  )

if (nrow(invalid_comparison_windows) > 0L) {
  stop(
    "Invalid comparison period for: ",
    paste(
      invalid_comparison_windows$geo_code,
      collapse = ", "
    ),
    ". Each defined window must span exactly ten years.",
    call. = FALSE
  )
}


# ------------------------------------------------------------
# 3. Retrieve exact endpoints for both investment series
#
# All economies:
# Retain both series, including unavailable results.
# Retrieve observations at the selected start and end years.
# Do not substitute nearby years.
#
# Brazil (BR):
# Both series use the same 2011 and 2021 endpoints.
#
# These comparison endpoints do not replace the independent
# headline CAGR endpoints calculated in Chunk 5.
# ------------------------------------------------------------

comparison_endpoints <- tidyr::expand_grid(
  geo_code = geo_list,
  series = real_investment_variables
) %>%
  left_join(
    comparison_windows,
    by = "geo_code"
  ) %>%
  left_join(
    real_investment_long %>%
      transmute(
        geo_code,
        series,
        start_year = year,
        start_value = real_value
      ),
    by = c("geo_code", "series", "start_year")
  ) %>%
  left_join(
    real_investment_long %>%
      transmute(
        geo_code,
        series,
        end_year = year,
        end_value = real_value
      ),
    by = c("geo_code", "series", "end_year")
  ) %>%
  mutate(
    n_years = end_year - start_year,

    valid_endpoints =
      is.finite(start_value) &
      start_value > 0 &
      is.finite(end_value) &
      end_value > 0 &
      !is.na(n_years) &
      n_years == 10,

    endpoint_ratio = if_else(
      valid_endpoints,
      end_value / start_value,
      NA_real_
    ),

    cagr_real = endpoint_ratio^(1 / 10) - 1,

    cagr_real = if_else(
      is.finite(cagr_real),
      cagr_real,
      NA_real_
    )
  ) %>%
  select(
    geo_code,
    series,
    start_year,
    end_year,
    start_value,
    end_value,
    n_years,
    cagr_real
  ) %>%
  arrange(
    match(geo_code, geo_list),
    match(series, real_investment_variables)
  )


# ------------------------------------------------------------
# 4. Build the complete annual index timeline
#
# All economies:
# Create eleven calendar years for each selected window
# and for each investment series.
#
# Each series uses its own investment level in the common
# starting year as its base:
#
#   index_100 = 100 * real_value / base_value
#
# Require finite, strictly positive observations, consistent
# with the index eligibility rule.
#
# If the exact base-year value is unavailable or invalid:
# keep that series' index unavailable throughout the window.
# Do not silently rebase to a later year.
#
# Missing calendar years and missing source values stay NA.
# No interpolation is performed.
#
# India (IND):
# The manual intensity override has no effect here.
# These indices use constant-price investment levels.
# ------------------------------------------------------------

indexed_real_series <- tidyr::expand_grid(
  geo_code = geo_list,
  series = real_investment_variables,
  year_offset = 0:10
) %>%
  left_join(
    comparison_windows,
    by = "geo_code"
  ) %>%
  filter(
    is.finite(start_year),
    is.finite(end_year)
  ) %>%
  mutate(
    year = start_year + year_offset,
    base_year = start_year
  ) %>%
  left_join(
    real_investment_long %>%
      select(
        geo_code,
        series,
        year,
        real_value
      ),
    by = c("geo_code", "series", "year")
  ) %>%
  left_join(
    comparison_endpoints %>%
      transmute(
        geo_code,
        series,
        base_year = start_year,
        base_value = start_value
      ),
    by = c("geo_code", "series", "base_year")
  ) %>%
  mutate(
    index_100 = if_else(
      is.finite(base_value) &
        base_value > 0 &
        is.finite(real_value) &
        real_value > 0,

      100 * (real_value / base_value),

      NA_real_
    ),

    index_100 = if_else(
      is.finite(index_100),
      index_100,
      NA_real_
    )
  ) %>%
  select(
    geo_code,
    series,
    year,
    base_year,
    base_value,
    real_value,
    index_100
  ) %>%
  arrange(
    match(geo_code, geo_list),
    match(series, real_investment_variables),
    year
  )


# ------------------------------------------------------------
# 5. Prepare the comparison CAGR table for Figure 3
#
# All economies:
# Retain the selected period and unrounded decimal CAGR.
# A value of 0.05 means 5% annual growth.
#
# Brazil (BR):
# Figure 3 uses the common 2011–2021 comparison period.
# The headline KPI, key insight and global Figure 10 continue
# to use investment_cagr from Chunk 5.
#
# Missing comparison growth remains NA, never zero.
# ------------------------------------------------------------

investment_comparison_cagr <- comparison_endpoints %>%
  left_join(
    economy_lookup %>%
      select(geo_code, economy_name),
    by = "geo_code"
  ) %>%
  select(
    geo_code,
    economy_name,
    series,
    start_year,
    end_year,
    cagr_real
  ) %>%
  arrange(
    match(geo_code, geo_list),
    match(series, real_investment_variables)
  )


# Inspect the selected periods and comparison growth.
comparison_windows
investment_comparison_cagr

```



# I calculated historical and latest investment intensities using adjusted value added, except for Canada, where I used adjusted GDP. I enforced India’s external 10% intangible intensity for its latest available intangible-investment year.
```{r investment_intensity, message=FALSE}

# ============================================================
# CHUNK 7 — INVESTMENT INTENSITY
#
# Outputs:
#   intensity_history    — annual series
#   investment_intensity — latest values, years and sources
#
# Supports Figure 1, Figure 9, KPI 3 and key insight 4.
# Values are percentages: 10 means 10%.
# ============================================================


# ------------------------------------------------------------
# 1. Define India's approved manual exception
#
# India (IND):
# Enforce the externally supplied intangible intensity of 10%.
# Apply it to the latest year with available I_Intang data.
#
# This exception remains active until this setting is updated.
# Earlier observations are calculated from the database.
# ------------------------------------------------------------

india_intangible_intensity_override <- 10

economy_rules <- economy_rules %>%
  mutate(
    intensity_source = case_when(
      geo_code == "CA"  ~ "GDP_Adj",
      geo_code == "IND" ~ "VAadj_with_manual_latest_intangible",
      TRUE             ~ "VAadj"
    )
  )


# ------------------------------------------------------------
# 2. Validate and prepare the database inputs
#
# Canada (CA):
# Use GDP_Adj as the denominator.
#
# Other economies, including India's historical calculations:
# Use VAadj, following the original formula.
# ------------------------------------------------------------

intensity_variables <- c(
  "I_Intang",
  "I_TangNRes"
)

intensity_required_columns <- c(
  intensity_variables,
  "VAadj",
  "GDP_Adj"
)

missing_intensity_columns <- setdiff(
  intensity_required_columns,
  names(data_intan_inflow)
)

if (length(missing_intensity_columns) > 0L) {
  stop(
    "Required intensity columns are missing: ",
    paste(missing_intensity_columns, collapse = ", "),
    call. = FALSE
  )
}

intensity_input <- data_intan_inflow %>%
  select(
    geo_code,
    year,
    all_of(intensity_required_columns)
  ) %>%
  filter(!is.na(year)) %>%
  mutate(
    across(
      all_of(intensity_required_columns),
      ~ to_numeric_checked(.x, cur_column())
    ),

    denominator = if_else(
      geo_code == "CA",
      GDP_Adj,
      VAadj
    )
  )


# ------------------------------------------------------------
# 3. Identify the year for India's manual exception
#
# India (IND):
# Derive the year from its latest available intangible
# investment observation, independently of other economies.
#
# If no observation year can be established, keep the year NA.
# ------------------------------------------------------------

india_available_years <- intensity_input %>%
  filter(
    geo_code == "IND",
    is.finite(I_Intang)
  ) %>%
  pull(year)

india_override_year <- if (length(india_available_years) > 0L) {
  max(india_available_years)
} else {
  NA_real_
}


# ------------------------------------------------------------
# 4. Calculate intensities and apply India's exception
#
# General calculation:
#   Intensity = investment / denominator * 100.
#
# Missing inputs or a zero denominator produce NA.
# An observed investment value of zero remains a valid 0%.
#
# India (IND):
# Replace only the latest intangible intensity with 10%.
# Its tangible intensity continues to use the database.
# Record the manual origin alongside the observation.
# ------------------------------------------------------------

intensity_calculations <- intensity_input %>%
  pivot_longer(
    cols = all_of(intensity_variables),
    names_to = "series",
    values_to = "investment_value"
  ) %>%
  mutate(
    intensity_pct_calculated = 100 * if_else(
      is.finite(investment_value) &
        is.finite(denominator) &
        denominator != 0,

      investment_value / denominator,

      NA_real_
    ),

    manual_override =
      !is.na(india_override_year) &
      geo_code == "IND" &
      series == "I_Intang" &
      year == india_override_year,

    intensity_pct = if_else(
      manual_override,
      india_intangible_intensity_override,
      intensity_pct_calculated
    ),

    value_source = case_when(
      manual_override  ~ "Manual override",
      geo_code == "CA" ~ "GDP_Adj",
      TRUE            ~ "VAadj"
    )
  )


# ------------------------------------------------------------
# 5. Build the annual timeline
#
# All economies:
# Derive the range from their available intensity observations.
# Preserve missing calendar years and series values as NA.
#
# Brazil (BR):
# Preserve missing tangible observations after its series ends.
#
# Poland (PL):
# Preserve missing tangible observations before its series starts.
# ------------------------------------------------------------

intensity_history <- intensity_calculations %>%
  filter(is.finite(intensity_pct)) %>%
  select(
    geo_code,
    year,
    series,
    intensity_pct,
    value_source
  )

if (nrow(intensity_history) > 0L) {

  intensity_history <- intensity_history %>%
    group_by(geo_code) %>%
    tidyr::complete(
      year = seq.int(min(year), max(year)),
      series = intensity_variables
    ) %>%
    ungroup()

}

intensity_history <- intensity_history %>%
  arrange(
    match(geo_code, geo_list),
    year,
    match(series, intensity_variables)
  )


# ------------------------------------------------------------
# 6. Select the latest valid intensity for each series
#
# Keep separate intangible and tangible observation years.
# This preserves Brazil's different periods for insight wording.
# ------------------------------------------------------------

latest_intensity_long <- intensity_history %>%
  filter(!is.na(intensity_pct)) %>%
  group_by(geo_code, series) %>%
  slice_max(
    order_by = year,
    n = 1,
    with_ties = FALSE
  ) %>%
  ungroup()


# ------------------------------------------------------------
# 7. Build the summary for all 29 economies
#
# India (IND):
# Explicitly retain the approved 10% in the summary.
# Its year remains NA if the database provides no usable year.
#
# Source columns document the origin of each latest value.
# Values remain numeric and unrounded.
# ------------------------------------------------------------

investment_intensity <- economy_lookup %>%
  select(
    geo_code,
    economy_name,
    display_order
  ) %>%
  left_join(
    economy_rules %>%
      select(geo_code, intensity_source),
    by = "geo_code"
  ) %>%
  left_join(
    latest_intensity_long %>%
      filter(series == "I_Intang") %>%
      transmute(
        geo_code,
        intensity_year_intangible = year,
        intangible_intensity_pct = intensity_pct,
        intangible_intensity_source = value_source
      ),
    by = "geo_code"
  ) %>%
  left_join(
    latest_intensity_long %>%
      filter(series == "I_TangNRes") %>%
      transmute(
        geo_code,
        intensity_year_tangible = year,
        tangible_intensity_pct = intensity_pct,
        tangible_intensity_source = value_source
      ),
    by = "geo_code"
  ) %>%
  mutate(
    intangible_intensity_pct = if_else(
      geo_code == "IND",
      india_intangible_intensity_override,
      intangible_intensity_pct
    ),

    intensity_year_intangible = if_else(
      geo_code == "IND",
      india_override_year,
      intensity_year_intangible
    ),

    intangible_intensity_source = if_else(
      geo_code == "IND",
      "Manual override",
      intangible_intensity_source
    )
  ) %>%
  arrange(display_order)

investment_intensity

```


# I calculated ten-year real growth by intangible asset and ranked eligible assets to identify the fastest-growing asset and the top three. I applied Canada’s and the Netherlands’ fixed periods and excluded Other IP products from growth comparisons for India, the Netherlands and Poland.
```{r asset_growth, message=FALSE}

# ============================================================
# CHUNK 8 — REAL GROWTH BY INTANGIBLE ASSET TYPE
#
# Outputs:
#   asset_lookup         — shared asset definitions
#   asset_cagr_real      — growth, periods and endpoint values
#   asset_growth_ranked  — valid assets ranked within economies
#   asset_growth_top3    — assets used in key insight 5
#   fastest_growing_asset— asset used in KPI 4
#
# Use constant-price Iq_* variables.
# Growth remains a decimal: 0.06 represents 6%.
# ============================================================


# ------------------------------------------------------------
# 1. Define the seven assets
#
# Keep the full dictionary for subsequent composition tables.
# Country-specific growth exclusions are applied separately.
# ------------------------------------------------------------

asset_lookup <- tibble::tribble(
  ~asset_code, ~asset_name,                 ~real_variable, ~nominal_variable,
  "RD",        "Research & development",    "Iq_RD",        "I_RD",
  "Soft_DB",   "Software & databases",      "Iq_Soft_DB",   "I_Soft_DB",
  "OrgCap",    "Organizational capital",    "Iq_OrgCap",    "I_OrgCap",
  "Brand",     "Brands",                    "Iq_Brand",     "I_Brand",
  "Design",    "Design",                    "Iq_Design",    "I_Design",
  "NFP",       "New financial products",    "Iq_NFP",       "I_NFP",
  "OIPP",      "Other IP products",         "Iq_OIPP",      "I_OIPP"
) %>%
  mutate(asset_order = row_number())

asset_vars_real <- asset_lookup$real_variable


# ------------------------------------------------------------
# 2. Validate the real asset inputs
# ------------------------------------------------------------

missing_asset_columns <- setdiff(
  asset_vars_real,
  names(data_intan_inflow)
)

if (length(missing_asset_columns) > 0L) {
  stop(
    "Required real asset columns are missing: ",
    paste(missing_asset_columns, collapse = ", "),
    call. = FALSE
  )
}

asset_real_long <- data_intan_inflow %>%
  select(geo_code, year, all_of(asset_vars_real)) %>%
  filter(!is.na(year)) %>%
  mutate(
    across(
      all_of(asset_vars_real),
      ~ to_numeric_checked(.x, cur_column())
    )
  ) %>%
  pivot_longer(
    cols = all_of(asset_vars_real),
    names_to = "real_variable",
    values_to = "value"
  ) %>%
  left_join(
    asset_lookup %>% select(real_variable, asset_code),
    by = "real_variable"
  ) %>%
  select(geo_code, year, asset_code, value)


# ------------------------------------------------------------
# 3. Define the assets eligible for growth comparisons
#
# India (IND), Netherlands (NL), Poland (PL):
# Exclude OIPP from the growth chart and fastest-asset ranking,
# following the profile-specific rules defined in Chunk 2.
#
# These exclusions apply to growth only.
# All seven assets remain available for nominal composition.
# ------------------------------------------------------------

asset_growth_requirements <- tidyr::expand_grid(
  geo_code = geo_list,
  asset_code = asset_lookup$asset_code
) %>%
  left_join(
    economy_rules %>% select(geo_code, exclude_oipp_growth),
    by = "geo_code"
  ) %>%
  filter(!(asset_code == "OIPP" & exclude_oipp_growth)) %>%
  select(geo_code, asset_code)

asset_required_counts <- asset_growth_requirements %>%
  count(geo_code, name = "n_required_assets")


# ------------------------------------------------------------
# 4. Identify valid pairs exactly ten years apart
#
# All economies:
# Both endpoints must be finite and strictly positive.
# Intermediate observations are not required for CAGR.
# ------------------------------------------------------------

asset_valid_values <- asset_real_long %>%
  semi_join(
    asset_growth_requirements,
    by = c("geo_code", "asset_code")
  ) %>%
  filter(is.finite(value), value > 0)

asset_candidate_pairs <- asset_valid_values %>%
  transmute(
    geo_code,
    asset_code,
    start_year = year,
    end_year = year + 10,
    start_value = value
  ) %>%
  inner_join(
    asset_valid_values %>%
      transmute(
        geo_code,
        asset_code,
        end_year = year,
        end_value = value
      ),
    by = c("geo_code", "asset_code", "end_year")
  )


# ------------------------------------------------------------
# 5. Select one comparison window per economy
#
# General rule:
# Select the latest ten-year window with valid endpoints
# for every eligible asset, ensuring comparable growth rates.
#
# Canada (CA) and Netherlands (NL):
# Respect their fixed window settings from Chunk 2.
#
# Netherlands (NL):
# Retain available assets within its prescribed window.
# Unavailable asset growth remains NA.
#
# Latvia (LV) and Philippines (PHL):
# Earlier periods are selected automatically from availability.
# ------------------------------------------------------------

asset_growth_windows <- asset_candidate_pairs %>%
  group_by(geo_code, start_year, end_year) %>%
  summarise(
    n_available_assets = n_distinct(asset_code),
    .groups = "drop"
  ) %>%
  left_join(asset_required_counts, by = "geo_code") %>%
  left_join(
    economy_rules %>%
      select(
        geo_code,
        asset_growth_start_year,
        asset_growth_end_year
      ),
    by = "geo_code"
  ) %>%
  filter(
    n_available_assets == n_required_assets | geo_code == "NL",

    is.na(asset_growth_start_year) |
      start_year == asset_growth_start_year,

    is.na(asset_growth_end_year) |
      end_year == asset_growth_end_year
  ) %>%
  group_by(geo_code) %>%
  slice_max(
    order_by = end_year,
    n = 1,
    with_ties = FALSE
  ) %>%
  ungroup() %>%
  select(geo_code, start_year, end_year)


# ------------------------------------------------------------
# 6. Calculate growth and retain the supporting observations
#
# Starting from the requirements table retains every economy
# and eligible asset, including those with unavailable growth.
#
# Netherlands (NL):
# Missing assets retain the selected comparison period and NA.
#
# Japan (JP), and any other economy with declining assets:
# Preserve negative growth for the chart and later formatting.
# ------------------------------------------------------------

asset_cagr_real <- asset_growth_requirements %>%
  left_join(asset_growth_windows, by = "geo_code") %>%
  left_join(
    asset_candidate_pairs,
    by = c("geo_code", "asset_code", "start_year", "end_year")
  ) %>%
  left_join(asset_lookup, by = "asset_code") %>%
  left_join(
    economy_lookup %>% select(geo_code, economy_name),
    by = "geo_code"
  ) %>%
  mutate(
    n_years = end_year - start_year,

    endpoint_ratio = if_else(
      is.finite(start_value) &
        start_value > 0 &
        is.finite(end_value) &
        end_value > 0 &
        !is.na(n_years) &
        n_years == 10,

      end_value / start_value,

      NA_real_
    ),

    cagr_real = endpoint_ratio^(1 / 10) - 1,

    cagr_real = if_else(
      is.finite(cagr_real),
      cagr_real,
      NA_real_
    )
  ) %>%
  select(
    geo_code,
    economy_name,
    asset_code,
    asset_name,
    real_variable,
    asset_order,
    start_year,
    end_year,
    start_value,
    end_value,
    n_years,
    cagr_real
  ) %>%
  arrange(match(geo_code, geo_list), asset_order)


# ------------------------------------------------------------
# 7. Rank assets using full numerical precision
#
# All economies:
# Rank before applying display rounding.
# Resolve exact numerical ties using the dictionary order.
# Exclude unavailable growth from rankings.
#
# Poland (PL):
# OIPP is already excluded, so it cannot become the headline
# fastest-growing asset or appear among the three leaders.
# ------------------------------------------------------------

asset_growth_ranked <- asset_cagr_real %>%
  filter(is.finite(cagr_real)) %>%
  group_by(geo_code) %>%
  arrange(desc(cagr_real), asset_order, .by_group = TRUE) %>%
  mutate(asset_rank = row_number()) %>%
  ungroup() %>%
  arrange(match(geo_code, geo_list), asset_rank)

asset_growth_top3 <- asset_growth_ranked %>%
  filter(asset_rank <= 3L)

fastest_growing_asset <- asset_growth_ranked %>%
  filter(asset_rank == 1L)

asset_growth_ranked

```



# I calculated the seven assets’ normalized current-price shares, summing to 100% before rounding, and identified the three largest. I selected the latest complete year, with Canada fixed to 2023.
```{r asset_composition, message=FALSE}

# ============================================================
# CHUNK 9 — INTANGIBLE ASSET COMPOSITION
#
# Outputs:
#   asset_composition_years
#   asset_composition
#   asset_composition_ranked
#   asset_composition_top3
#
# Use current-price I_* variables.
# share     : fraction of total composition
# share_pct : percentage of total composition
# ============================================================


# ------------------------------------------------------------
# 1. Validate the nominal asset inputs
#
# Use the complete asset dictionary from Chunk 8.
#
# India (IND), Netherlands (NL), Poland (PL):
# Include OIPP in composition, even though it is excluded
# from their asset-growth comparisons.
# ------------------------------------------------------------

asset_vars_nominal <- asset_lookup$nominal_variable

composition_required_columns <- c(
  "I_Intang",
  asset_vars_nominal
)

missing_composition_columns <- setdiff(
  composition_required_columns,
  names(data_intan_inflow)
)

if (length(missing_composition_columns) > 0L) {
  stop(
    "Required composition columns are missing: ",
    paste(missing_composition_columns, collapse = ", "),
    call. = FALSE
  )
}

asset_nominal_long <- data_intan_inflow %>%
  select(
    geo_code,
    year,
    all_of(composition_required_columns)
  ) %>%
  filter(!is.na(year)) %>%
  mutate(
    across(
      all_of(composition_required_columns),
      ~ to_numeric_checked(.x, cur_column())
    )
  ) %>%
  pivot_longer(
    cols = all_of(asset_vars_nominal),
    names_to = "nominal_variable",
    values_to = "asset_value"
  ) %>%
  left_join(
    asset_lookup %>% select(nominal_variable, asset_code),
    by = "nominal_variable"
  ) %>%
  select(
    geo_code,
    year,
    asset_code,
    I_Intang,
    asset_value
  )


# ------------------------------------------------------------
# 2. Identify years with complete composition data
#
# All economies:
# Require finite observations for all seven assets and
# a finite, strictly positive I_Intang value.
#
# Observed zero asset values are retained.
# Require a positive asset sum for normalization.
# ------------------------------------------------------------

composition_candidates <- asset_nominal_long %>%
  filter(
    is.finite(I_Intang),
    I_Intang > 0,
    is.finite(asset_value)
  ) %>%
  group_by(geo_code, year) %>%
  summarise(
    n_available_assets = n_distinct(asset_code),
    total_intangible = first(I_Intang),
    asset_total = sum(asset_value),
    .groups = "drop"
  ) %>%
  filter(
    n_available_assets == nrow(asset_lookup),
    is.finite(asset_total),
    asset_total > 0
  )


# ------------------------------------------------------------
# 3. Select the composition year
#
# General rule:
# Use the latest year with complete valid composition data.
#
# Canada (CA):
# Respect its prescribed composition year in economy_rules.
# That year must satisfy the same completeness requirements.
#
# Latvia (LV) and Philippines (PHL):
# Earlier years are selected automatically from availability.
#
# Composition dates are determined independently of aggregate
# investment dates and real asset-growth windows.
# ------------------------------------------------------------

asset_composition_years <- composition_candidates %>%
  left_join(
    economy_rules %>%
      select(
        geo_code,
        required_composition_year = composition_year
      ),
    by = "geo_code"
  ) %>%
  filter(
    is.na(required_composition_year) |
      year == required_composition_year
  ) %>%
  group_by(geo_code) %>%
  slice_max(
    order_by = year,
    n = 1,
    with_ties = FALSE
  ) %>%
  ungroup() %>%
  transmute(
    geo_code,
    composition_year = year,
    total_intangible,
    asset_total
  )


# ------------------------------------------------------------
# 4. Calculate raw and normalized shares
#
# Original raw share:
#   asset_value / I_Intang
#
# Normalizing these raw shares across all seven assets is
# algebraically equivalent to:
#   asset_value / sum(all seven asset values)
#
# Keep both measures for traceability.
# Normalized shares sum to one before display rounding.
#
# Starting with all economy-asset combinations preserves
# missing results as NA when no eligible year is available.
# ------------------------------------------------------------

asset_composition <- tidyr::expand_grid(
  geo_code = geo_list,
  asset_code = asset_lookup$asset_code
) %>%
  left_join(
    asset_composition_years,
    by = "geo_code"
  ) %>%
  left_join(
    asset_nominal_long %>%
      transmute(
        geo_code,
        asset_code,
        composition_year = year,
        asset_value
      ),
    by = c("geo_code", "asset_code", "composition_year")
  ) %>%
  left_join(
    asset_lookup,
    by = "asset_code"
  ) %>%
  left_join(
    economy_lookup %>% select(geo_code, economy_name),
    by = "geo_code"
  ) %>%
  mutate(
    share_raw = asset_value / total_intangible,
    share = asset_value / asset_total,
    share_pct = share * 100
  ) %>%
  select(
    geo_code,
    economy_name,
    asset_code,
    asset_name,
    nominal_variable,
    asset_order,
    composition_year,
    asset_value,
    total_intangible,
    asset_total,
    share_raw,
    share,
    share_pct
  ) %>%
  arrange(
    match(geo_code, geo_list),
    asset_order
  )


# ------------------------------------------------------------
# 5. Rank the composition and extract its three largest assets
#
# All economies:
# Rank using unrounded normalized shares.
# Resolve exact numerical ties using the dictionary order.
#
# Figure 5 will use all seven available asset shares.
# Key insight 6 will use the three largest shares.
#
# Percentage rounding and narrative wording are applied later.
# ------------------------------------------------------------

asset_composition_ranked <- asset_composition %>%
  filter(is.finite(share)) %>%
  group_by(geo_code) %>%
  arrange(desc(share), asset_order, .by_group = TRUE) %>%
  mutate(asset_rank = row_number()) %>%
  ungroup() %>%
  arrange(
    match(geo_code, geo_list),
    asset_rank
  )

asset_composition_top3 <- asset_composition_ranked %>%
  filter(asset_rank <= 3L)

asset_composition_ranked

```



# I calculated the shares of intangible investment included in and excluded from national accounts using their combined total. I selected the latest valid year and prepared the two components for Figure 6.
```{r measured_unmeasured, message=FALSE}

# ============================================================
# CHUNK 10 — MEASURED VS. UNMEASURED INTANGIBLE INVESTMENT
#
# All economies:
# - Measured investment:   I_NatAcc
# - Unmeasured investment: I_NonNatAcc
# - Denominator: I_NatAcc + I_NonNatAcc, in the same year.
# - Select the latest year with both components available
#   and a finite, strictly positive combined total.
#
# Economy specificities:
# - No forced observation year applies to this indicator.
# - Canada's GDP_Adj denominator applies only to intensity.
# - India's manual 10% override applies only to intensity.
#
# Outputs:
# - measured_unmeasured: one row per economy.
# - measured_unmeasured_long: two rows per economy, with
#   Unmeasured first, matching the profile's doughnut chart.
#
# Percentages remain numeric and unrounded.
# ============================================================


# ------------------------------------------------------------
# 1. Check the required source columns
# ------------------------------------------------------------

measurement_variables <- c("I_NatAcc", "I_NonNatAcc")

missing_measurement_columns <- setdiff(
  measurement_variables,
  names(data_intan_inflow)
)

if (length(missing_measurement_columns) > 0L) {
  stop(
    "Missing columns for measured/unmeasured investment: ",
    paste(missing_measurement_columns, collapse = ", "),
    call. = FALSE
  )
}


# ------------------------------------------------------------
# 2. Prepare and validate the source values
#
# Use the numeric validation helper defined earlier.
# Missing values remain NA; they are never replaced with zero.
# ------------------------------------------------------------

measurement_input <- data_intan_inflow %>%
  select(
    geo_code,
    year,
    all_of(measurement_variables)
  ) %>%
  filter(!is.na(year)) %>%
  mutate(
    across(
      all_of(measurement_variables),
      ~ to_numeric_checked(.x, cur_column())
    ),
    total_intangible = I_NatAcc + I_NonNatAcc
  )


# ------------------------------------------------------------
# 3. Calculate percentages and select the latest valid year
#
# Both components must refer to the same economy and year.
# An observed zero is valid if the combined total is positive.
#
# Following the original calculation, the denominator is the
# sum of these two components, rather than I_Intang.
# ------------------------------------------------------------

measured_unmeasured_latest <- measurement_input %>%
  filter(
    is.finite(I_NatAcc),
    is.finite(I_NonNatAcc),
    is.finite(total_intangible),
    total_intangible > 0
  ) %>%
  mutate(
    measured_pct   = (I_NatAcc / total_intangible) * 100,
    unmeasured_pct = (I_NonNatAcc / total_intangible) * 100
  ) %>%
  filter(
    is.finite(measured_pct),
    is.finite(unmeasured_pct)
  ) %>%
  group_by(geo_code) %>%
  slice_max(
    order_by = year,
    n = 1,
    with_ties = FALSE
  ) %>%
  ungroup() %>%
  transmute(
    geo_code,
    split_year = year,
    I_NatAcc,
    I_NonNatAcc,
    total_intangible,
    measured_pct,
    unmeasured_pct
  )


# ------------------------------------------------------------
# 4. Retain every economy in the master output
#
# Economy names and ordering come from economy_lookup.
# If an economy has no valid observation, retain its row with
# NA values so that it does not disappear from the workflow.
#
# Keep split_year separately: this indicator may use a
# different year from levels, intensity or asset composition.
# ------------------------------------------------------------

measured_unmeasured <- economy_lookup %>%
  select(
    geo_code,
    economy_name,
    display_order
  ) %>%
  left_join(
    measured_unmeasured_latest,
    by = "geo_code"
  ) %>%
  arrange(display_order) %>%
  select(-display_order)


# ------------------------------------------------------------
# 5. Prepare Figure 6's two-component structure
#
# All economies:
# - Unmeasured appears first.
# - Measured appears second.
#
# share_pct is already expressed in percentage units:
# a value of 60 means 60%, not 0.60.
# Formatting and chart colours will be applied later.
# ------------------------------------------------------------

measured_unmeasured_long <- measured_unmeasured %>%
  select(
    geo_code,
    economy_name,
    split_year,
    unmeasured_pct,
    measured_pct
  ) %>%
  pivot_longer(
    cols = c(unmeasured_pct, measured_pct),
    names_to = "component",
    values_to = "share_pct"
  ) %>%
  mutate(
    component = case_when(
      component == "unmeasured_pct" ~ "Unmeasured",
      component == "measured_pct"   ~ "Measured"
    ),
    component_order = if_else(
      component == "Unmeasured",
      1L,
      2L
    )
  ) %>%
  arrange(
    match(geo_code, geo_list),
    component_order
  ) %>%
  select(-component_order)


measured_unmeasured

```



# I ranked economies by intangible investment levels, intensity and real ten-year growth using unrounded values. I created the underlying tables for Figures 7–10, separating the top five economies from the remaining economies for investment levels.
```{r global_comparisons, message=FALSE}

# ============================================================
# CHUNK 11 — GLOBAL CONTEXT / CROSS-ECONOMY COMPARISONS
#
# Figure 7:  Top 5 economies by intangible investment levels.
# Figure 8:  Remaining economies by intangible investment levels.
# Figure 9:  Latest intangible investment intensity.
# Figure 10: Latest 10-year real CAGR of intangible investment.
#
# All economies:
# - Use the latest eligible observations calculated earlier.
# - Keep the actual observation years and growth periods.
# - Rank using full precision, before display rounding.
# - Use the official display names from economy_lookup.
# - Keep missing results in the underlying global tables.
#   Only available results enter the figure tables.
#
# Economy specificities:
# - India (IND): preserve the 10% intensity override from Chunk 7.
# - Canada (CA): preserve intensity calculated using GDP_Adj.
# - Brazil (BR): Figure 10 uses the latest headline CAGR from
#   Chunk 5. Its fixed 2011–2021 comparison window applies to
#   Figures 2 and 3, not to this global comparison.
#
# No unit conversion or display rounding is performed here.
# ============================================================


# ------------------------------------------------------------
# 1. Global intangible investment levels
#
# Use the latest intangible level independently of the
# availability or observation year of tangible investment.
#
# Exact ties are ordered using the established economy order.
# level_position is a display position, not a tied rank.
# ------------------------------------------------------------

global_investment_levels <- economy_lookup %>%
  select(
    geo_code,
    economy_name,
    display_order
  ) %>%
  left_join(
    investment_levels %>%
      select(
        geo_code,
        latest_year_intangible,
        intangible_raw
      ),
    by = "geo_code"
  ) %>%
  arrange(
    desc(intangible_raw),
    display_order
  ) %>%
  mutate(
    level_position = if_else(
      is.finite(intangible_raw),
      row_number(),
      NA_integer_
    )
  ) %>%
  select(
    geo_code,
    economy_name,
    level_position,
    latest_year_intangible,
    intangible_raw
  )


# ------------------------------------------------------------
# 2. Figures 7 and 8 — divide the same ordered level table
#
# Select the top five dynamically each time the code runs.
# No economy is permanently assigned to either panel.
#
# intangible_raw retains the source PPP units.
# Billions/trillions and text labels will be handled later.
# ------------------------------------------------------------

figure7_top5_levels <- global_investment_levels %>%
  filter(
    !is.na(level_position),
    level_position <= 5L
  )

figure8_remaining_levels <- global_investment_levels %>%
  filter(
    !is.na(level_position),
    level_position > 5L
  )


# ------------------------------------------------------------
# 3. Global intangible investment intensity
#
# Reuse Chunk 7's finalized results so that the global chart,
# KPI card and key insight use the same value and year.
#
# India: retain the manual 10% value and its recorded source.
# Canada: retain the GDP_Adj-based value and recorded source.
#
# Intensity is already in percentage units:
# 10 means 10%, so do not multiply it by 100 again.
# ------------------------------------------------------------

global_intangible_intensity <- economy_lookup %>%
  select(
    geo_code,
    economy_name,
    display_order
  ) %>%
  left_join(
    investment_intensity %>%
      select(
        geo_code,
        intensity_year_intangible,
        intangible_intensity_pct,
        intangible_intensity_source
      ),
    by = "geo_code"
  ) %>%
  arrange(
    desc(intangible_intensity_pct),
    display_order
  ) %>%
  mutate(
    intensity_position = if_else(
      is.finite(intangible_intensity_pct),
      row_number(),
      NA_integer_
    )
  ) %>%
  select(
    geo_code,
    economy_name,
    intensity_position,
    intensity_year_intangible,
    intangible_intensity_pct,
    intangible_intensity_source
  )

figure9_intensity_all <- global_intangible_intensity %>%
  filter(!is.na(intensity_position))


# ------------------------------------------------------------
# 4. Global real CAGR of intangible investment
#
# Use the headline CAGR from Chunk 5.
# Each economy retains its own start and end years.
#
# Brazil: do not substitute the comparison CAGR from Chunk 6.
#
# Preserve negative growth rates and rank them normally.
# Convert decimal growth to percentage units exactly once,
# while retaining the original decimal value.
# ------------------------------------------------------------

global_intangible_cagr <- economy_lookup %>%
  select(
    geo_code,
    economy_name,
    display_order
  ) %>%
  left_join(
    investment_cagr %>%
      select(
        geo_code,
        start_year_intangible,
        end_year_intangible,
        intangible_cagr_real
      ),
    by = "geo_code"
  ) %>%
  mutate(
    intangible_cagr_pct = intangible_cagr_real * 100
  ) %>%
  arrange(
    desc(intangible_cagr_real),
    display_order
  ) %>%
  mutate(
    cagr_position = if_else(
      is.finite(intangible_cagr_real),
      row_number(),
      NA_integer_
    )
  ) %>%
  select(
    geo_code,
    economy_name,
    cagr_position,
    start_year_intangible,
    end_year_intangible,
    intangible_cagr_real,
    intangible_cagr_pct
  )

figure10_cagr_all <- global_intangible_cagr %>%
  filter(!is.na(cagr_position))


# ------------------------------------------------------------
# 5. Inspect the four shared figure tables
#
# These values will be reused across all economy profiles.
# The economy highlighted in each profile will be assigned
# later when its presentation settings are assembled.
# ------------------------------------------------------------

list(
  Figure_7  = figure7_top5_levels,
  Figure_8  = figure8_remaining_levels,
  Figure_9  = figure9_intensity_all,
  Figure_10 = figure10_cagr_all
)

```


# I defined consistent unit conversions, rounding, percentage labels and prose formatting, including displaying PPP investment levels in billions or trillions of USD. I kept calculation values separate from display values and accommodated Japan’s small negative growth label.

```{r profile_number_formatting, message=FALSE}

# ============================================================
# CHUNK 12 — SHARED NUMBER FORMATTING
#
# All economies:
# - Keep calculation and ranking values at full precision.
# - Apply rounding only when preparing display values.
# - Use decimal half-up rounding: 34.5 becomes 35.
# - For negative values, exact halves round away from zero.
# - Display at most one decimal, removing a trailing ".0".
# - Composition percentages in insights use whole numbers.
# - Use "percent" in prose and "%" in cards/chart labels.
# - Use "USD", never the dollar symbol.
#
# Economy specificities:
# - India: the accepted intensity of 10 displays as "10%"
#   or "10 percent", without changing its value.
# - Canada and Czech Republic: composition values ending
#   in .5 round upwards for the whole-number insight.
# - Japan: a small negative asset growth rate can retain
#   the label "-0.0%" when requested for its chart.
#
# These helpers do not overwrite the calculation tables.
# ============================================================


# ------------------------------------------------------------
# 1. Declare the units of the source PPP level variables
#
# The supplied profile instructions specify millions of USD.
# This is an explicit input convention, not an inference
# from the size of the values.
#
# If a future source workbook changes its units, update this
# setting before running the remaining chunks.
# ------------------------------------------------------------

ppp_input_unit <- "million_USD"

ppp_billions_multipliers <- c(
  USD          = 1e-9,
  thousand_USD = 1e-6,
  million_USD  = 1e-3,
  billion_USD  = 1,
  trillion_USD = 1e3
)

if (
  length(ppp_input_unit) != 1L ||
  is.na(ppp_input_unit) ||
  !ppp_input_unit %in% names(ppp_billions_multipliers)
) {
  stop(
    "ppp_input_unit must be one of: ",
    paste(names(ppp_billions_multipliers), collapse = ", "),
    call. = FALSE
  )
}

ppp_to_billions <- function(x) {
  x * unname(ppp_billions_multipliers[[ppp_input_unit]])
}


# ------------------------------------------------------------
# 2. Decimal half-up rounding
#
# Convert each value to a decimal representation using
# 15 significant digits, consistent with Excel's numeric
# precision. This conversion is only for final display;
# calculation and ranking values remain untouched.
#
# Round by inspecting the first discarded decimal digit.
# Increment retained digits as text to avoid binary
# arithmetic at the rounding boundary.
#
# Supports the two precisions required by the profiles:
# - digits = 1: levels, growth, intensity and chart shares.
# - digits = 0: composition shares in key insights.
#
# Missing/non-finite values return NA, never a zero.
# ------------------------------------------------------------

round_half_up_text <- function(x, digits = 1L) {

  if (!is.numeric(x)) {
    stop("Formatting requires numeric input.", call. = FALSE)
  }

  if (
    length(digits) != 1L ||
    is.na(digits) ||
    !digits %in% c(0L, 1L)
  ) {
    stop("digits must be 0 or 1.", call. = FALSE)
  }

  vapply(
    x,
    function(value) {

      if (!is.finite(value)) {
        return(NA_character_)
      }

      decimal_text <- format(
        abs(value),
        digits = 15,
        scientific = FALSE,
        trim = TRUE,
        big.mark = "",
        decimal.mark = "."
      )

      parts <- strsplit(
        decimal_text,
        split = ".",
        fixed = TRUE
      )[[1L]]

      whole <- parts[1L]

      fraction <- if (length(parts) > 1L) {
        parts[2L]
      } else {
        ""
      }

      # Pad so that the first discarded digit always exists.
      fraction <- paste0(
        fraction,
        strrep("0", digits + 1L)
      )

      retained <- paste0(
        whole,
        if (digits == 0L) "" else substr(fraction, 1L, digits)
      )

      retained_digits <- strsplit(
        retained,
        split = "",
        fixed = TRUE
      )[[1L]]

      discarded_digit <- as.integer(
        substr(fraction, digits + 1L, digits + 1L)
      )

      carry <- discarded_digit >= 5L

      # Decimal carry, including cases such as 9.95 -> 10.0.
      for (j in rev(seq_along(retained_digits))) {

        if (!carry) break

        next_digit <- as.integer(retained_digits[j]) + 1L

        retained_digits[j] <- as.character(next_digit %% 10L)
        carry <- next_digit >= 10L
      }

      if (carry) {
        retained_digits <- c("1", retained_digits)
      }

      result <- paste0(retained_digits, collapse = "")

      if (digits == 1L) {
        n <- nchar(result)

        result <- paste0(
          substr(result, 1L, n - 1L),
          ".",
          substr(result, n, n)
        )
      }

      if (value < 0) {
        result <- paste0("-", result)
      }

      result
    },
    character(1),
    USE.NAMES = FALSE
  )
}


# Numeric counterpart for rounded Excel data cells.
# Apply directly to the original value, not an already
# rounded label, to avoid successive rounding.

round_profile <- function(x, digits = 1L) {
  as.numeric(round_half_up_text(x, digits = digits))
}


# ------------------------------------------------------------
# 3. General number labels
#
# Remove ".0" after rounding.
#
# Japan's small negative asset CAGR:
# keep_negative_zero = TRUE preserves "-0.0" when a negative
# value rounds to zero at one decimal place.
# The underlying unrounded value remains available.
# ------------------------------------------------------------

format_profile_number <- function(
  x,
  digits = 1L,
  keep_negative_zero = FALSE
) {

  result <- round_half_up_text(x, digits = digits)

  negative_zero <- (
    is.finite(x) &
    x < 0 &
    as.numeric(result) == 0
  )

  negative_zero[is.na(negative_zero)] <- FALSE

  if (!keep_negative_zero) {
    result[negative_zero] <- sub(
      "^-",
      "",
      result[negative_zero]
    )
  }

  strip_decimal <- rep(TRUE, length(result))

  if (keep_negative_zero) {
    strip_decimal[negative_zero] <- FALSE
  }

  result[strip_decimal] <- sub(
    "\\.0$",
    "",
    result[strip_decimal]
  )

  result
}


# ------------------------------------------------------------
# 4. Percentage labels
#
# Input must already be in percentage units:
# 10 means 10%, not 0.10.
#
# Use digits = 0 for composition percentages in insights.
# All other percentage labels normally use digits = 1.
# ------------------------------------------------------------

format_profile_percent <- function(
  pct,
  digits = 1L,
  prose = FALSE,
  keep_negative_zero = FALSE
) {

  number <- format_profile_number(
    pct,
    digits = digits,
    keep_negative_zero = keep_negative_zero
  )

  suffix <- if (prose) " percent" else "%"

  ifelse(
    is.na(number),
    NA_character_,
    paste0(number, suffix)
  )
}


# Growth calculations are stored as decimal rates.
# This wrapper converts them to percentage units once.
#
# Do not use it for intensity or share_pct, which are
# already expressed in percentage units.

format_profile_growth <- function(
  rate,
  prose = FALSE,
  keep_negative_zero = FALSE
) {

  format_profile_percent(
    pct = rate * 100,
    digits = 1L,
    prose = prose,
    keep_negative_zero = keep_negative_zero
  )
}


# ------------------------------------------------------------
# 5. Investment level labels
#
# Convert source PPP levels to billions first.
#
# Use trillions when the unrounded magnitude reaches
# 1,000 billion. This produces the US trillions label and
# automatically handles other economies reaching that
# threshold in future editions.
#
# Styles:
# - card:  "USD 4.9T" / "USD 312.3B"
# - prose: "USD 4.9 trillion" / "USD 312.3 billion"
# - axis:  "4.9T" / "312.3B"
#
# Global chart numeric values remain in billions even when
# the accompanying text label uses trillions.
# ------------------------------------------------------------

format_profile_level <- function(
  level_raw,
  style = c("card", "prose", "axis")
) {

  style <- match.arg(style)

  billions <- ppp_to_billions(level_raw)
  use_trillions <- abs(billions) >= 1000

  display_value <- ifelse(
    use_trillions,
    billions / 1000,
    billions
  )

  number <- format_profile_number(display_value)

  short_unit <- ifelse(use_trillions, "T", "B")
  long_unit  <- ifelse(use_trillions, "trillion", "billion")

  result <- switch(
    style,
    card  = paste0("USD ", number, short_unit),
    prose = paste("USD", number, long_unit),
    axis  = paste0(number, short_unit)
  )

  result[is.na(number)] <- NA_character_

  result
}


# ------------------------------------------------------------
# 6. Prepare level fields for profile assembly
#
# Preserve the source values and observation years.
# Add numeric values in billions and separate text labels.
#
# Numeric billions remain unrounded here so that later
# calculations never depend on formatted text.
# ------------------------------------------------------------

investment_levels_display <- investment_levels %>%
  mutate(
    intangible_billions = ppp_to_billions(intangible_raw),
    tangible_billions   = ppp_to_billions(tangible_raw),

    intangible_level_card = format_profile_level(
      intangible_raw,
      style = "card"
    ),

    intangible_level_prose = format_profile_level(
      intangible_raw,
      style = "prose"
    ),

    intangible_level_axis = format_profile_level(
      intangible_raw,
      style = "axis"
    ),

    tangible_level_card = format_profile_level(
      tangible_raw,
      style = "card"
    ),

    tangible_level_prose = format_profile_level(
      tangible_raw,
      style = "prose"
    )
  )


investment_levels_display %>%
  select(
    geo_code,
    economy_name,
    latest_year_intangible,
    intangible_billions,
    intangible_level_card,
    intangible_level_prose
  )

```




# I built the four headline cards covering intangible investment, ten-year growth, intangible intensity and the fastest-growing asset. I assigned their values, labels, explanatory text, colours and supporting observation periods.
```{r profile_kpi_cards, message=FALSE}

# ============================================================
# CHUNK 13 — BUILD THE FOUR KPI CARDS
#
# Card 1: Intangible investment level.
#         Subtext: latest real year-on-year growth.
#
# Card 2: Latest 10-year real intangible investment CAGR.
#
# Card 3: Latest intangible investment intensity.
#
# Card 4: Fastest-growing eligible intangible asset.
#         The asset name itself is the card label.
#
# Economy specificities:
# - India: preserve Chunk 7's accepted 10% intensity.
# - Canada: preserve the GDP_Adj-based intensity.
# - Brazil: Card 2 uses the headline CAGR from Chunk 5,
#   independently of the comparison window in Figures 2–3.
# - India, Netherlands and Poland: Card 4 inherits the OIPP
#   exclusions already applied in Chunk 8.
# - Canada and Netherlands: Card 4 inherits their configured
#   asset-growth windows from Chunk 8.
#
# Negative growth:
# - Reference YoY contractions in Denmark, Luxembourg and
#   Portugal are shown in red, with contraction wording.
# - The reference negative Netherlands CAGR is shown in red.
# - These behaviours follow the calculated sign, rather
#   than being permanently assigned to those economies.
# ============================================================


# ------------------------------------------------------------
# 1. Shared card colours and missing-value label
#
# Descriptive colour strings match the reference workbook.
# Background and border colours are stored separately.
# ------------------------------------------------------------

kpi_colours <- c(
  text     = "Black #1c1c1c",
  muted    = "Grey #78716c",
  positive = "Green #059669",
  negative = "Red #dc2626"
)

kpi_missing_label <- "Not available"


# ------------------------------------------------------------
# 2. Growth colour and YoY subtext helpers
#
# Determine direction using the unrounded rate.
# A contraction uses its absolute magnitude in the subtext:
# "3.2% year-on-year contraction".
#
# A true zero is neutral; missing growth is never shown as 0%.
# ------------------------------------------------------------

kpi_growth_colour <- function(rate) {
  case_when(
    !is.finite(rate) ~ kpi_colours[["muted"]],
    rate < 0        ~ kpi_colours[["negative"]],
    rate > 0        ~ kpi_colours[["positive"]],
    TRUE            ~ kpi_colours[["muted"]]
  )
}

kpi_yoy_subtext <- function(rate) {
  case_when(
    !is.finite(rate) ~ "Year-on-year growth unavailable",

    rate < 0 ~ paste0(
      format_profile_growth(abs(rate)),
      " year-on-year contraction"
    ),

    rate > 0 ~ paste0(
      format_profile_growth(rate),
      " year-on-year growth"
    ),

    TRUE ~ "No year-on-year change"
  )
}


# ------------------------------------------------------------
# 3. Assemble the inputs, one row per economy
#
# Each indicator keeps its own observation year or period.
# In particular, Card 1's investment level and YoY subtext
# are not assumed to have the same latest year.
#
# Reuse the previously selected fastest asset. Do not rank
# rounded display values or reintroduce excluded assets.
# ------------------------------------------------------------

kpi_inputs <- economy_lookup %>%
  select(
    geo_code,
    economy_name,
    display_order
  ) %>%
  left_join(
    investment_levels_display %>%
      select(
        geo_code,
        latest_year_intangible,
        intangible_billions,
        intangible_level_card
      ),
    by = "geo_code"
  ) %>%
  left_join(
    investment_yoy %>%
      select(
        geo_code,
        yoy_start_year_intangible,
        yoy_end_year_intangible,
        intangible_yoy_real
      ),
    by = "geo_code"
  ) %>%
  left_join(
    investment_cagr %>%
      select(
        geo_code,
        start_year_intangible,
        end_year_intangible,
        intangible_cagr_real
      ),
    by = "geo_code"
  ) %>%
  left_join(
    investment_intensity %>%
      select(
        geo_code,
        intensity_year_intangible,
        intangible_intensity_pct,
        intangible_intensity_source
      ),
    by = "geo_code"
  ) %>%
  left_join(
    fastest_growing_asset %>%
      transmute(
        geo_code,
        fastest_asset_code = asset_code,
        fastest_asset_name = asset_name,
        fastest_asset_variable = real_variable,
        fastest_asset_start_year = start_year,
        fastest_asset_end_year = end_year,
        fastest_asset_cagr = cagr_real
      ),
    by = "geo_code"
  ) %>%
  arrange(display_order)


# ------------------------------------------------------------
# 4. Card 1 — Intangible investment
#
# The display label may use billions or trillions.
# The accompanying numeric value is consistently in billions.
#
# Store the YoY period separately from the level year.
# ------------------------------------------------------------

kpi_card1 <- kpi_inputs %>%
  transmute(
    geo_code,
    economy_name,
    kpi_id = 1L,

    label = "Intangible investment",

    value = coalesce(
      intangible_level_card,
      kpi_missing_label
    ),

    value_colour = if_else(
      is.finite(intangible_billions),
      kpi_colours[["text"]],
      kpi_colours[["muted"]]
    ),

    subtext = kpi_yoy_subtext(intangible_yoy_real),
    subtext_colour = kpi_growth_colour(intangible_yoy_real),

    value_numeric = intangible_billions,
    value_unit = "Billion USD (PPP, current prices)",
    value_source = "I_Intang_PPP",
    value_start_year = latest_year_intangible,
    value_end_year = latest_year_intangible,

    subtext_value_numeric = intangible_yoy_real * 100,
    subtext_unit = "Percent",
    subtext_source = "Iq_Intang",
    subtext_start_year = yoy_start_year_intangible,
    subtext_end_year = yoy_end_year_intangible
  )


# ------------------------------------------------------------
# 5. Card 2 — Growth in intangibles
#
# Retain the minus sign for a negative CAGR.
# Keep the exact start/end years behind the "10-year CAGR"
# label, including Brazil's independently selected period.
# ------------------------------------------------------------

kpi_card2 <- kpi_inputs %>%
  transmute(
    geo_code,
    economy_name,
    kpi_id = 2L,

    label = "Growth in intangibles",

    value = coalesce(
      format_profile_growth(
        intangible_cagr_real,
        keep_negative_zero = TRUE
      ),
      kpi_missing_label
    ),

    value_colour = kpi_growth_colour(intangible_cagr_real),
    subtext = "10-year CAGR",
    subtext_colour = kpi_colours[["muted"]],

    value_numeric = intangible_cagr_real * 100,
    value_unit = "Percent per year",
    value_source = "Iq_Intang",
    value_start_year = start_year_intangible,
    value_end_year = end_year_intangible
  )


# ------------------------------------------------------------
# 6. Card 3 — Intangible intensity
#
# Reuse the finalized intensity and its recorded source.
# India therefore displays 10%, and Canada retains GDP_Adj.
#
# The value is already in percentage units.
# ------------------------------------------------------------

kpi_card3 <- kpi_inputs %>%
  transmute(
    geo_code,
    economy_name,
    kpi_id = 3L,

    label = "Intangible intensity",

    value = coalesce(
      format_profile_percent(intangible_intensity_pct),
      kpi_missing_label
    ),

    value_colour = if_else(
      is.finite(intangible_intensity_pct),
      kpi_colours[["text"]],
      kpi_colours[["muted"]]
    ),

    subtext = "Share of GDP",
    subtext_colour = kpi_colours[["muted"]],

    value_numeric = intangible_intensity_pct,
    value_unit = "Percent",
    value_source = intangible_intensity_source,
    value_start_year = intensity_year_intangible,
    value_end_year = intensity_year_intangible
  )


# ------------------------------------------------------------
# 7. Card 4 — Fastest-growing eligible asset
#
# Match the reference: asset name as label, CAGR as value.
# Positive values use black for this card.
#
# Future editions:
# If the highest eligible CAGR is zero or negative, use
# "Highest asset CAGR" instead of implying positive growth.
# A negative value is displayed in red.
# ------------------------------------------------------------

kpi_card4 <- kpi_inputs %>%
  transmute(
    geo_code,
    economy_name,
    kpi_id = 4L,

    label = coalesce(
      fastest_asset_name,
      "Fastest-growing asset"
    ),

    value = coalesce(
      format_profile_growth(
        fastest_asset_cagr,
        keep_negative_zero = TRUE
      ),
      kpi_missing_label
    ),

    value_colour = case_when(
      !is.finite(fastest_asset_cagr) ~ kpi_colours[["muted"]],
      fastest_asset_cagr < 0        ~ kpi_colours[["negative"]],
      TRUE                         ~ kpi_colours[["text"]]
    ),

    subtext = case_when(
      !is.finite(fastest_asset_cagr) ~
        "Asset growth unavailable",

      fastest_asset_cagr > 0 ~
        "Fastest-growing asset type\n10-year CAGR",

      TRUE ~
        "Highest asset CAGR\n10-year CAGR"
    ),

    subtext_colour = kpi_colours[["muted"]],

    value_numeric = fastest_asset_cagr * 100,
    value_unit = "Percent per year",
    value_source = fastest_asset_variable,
    value_start_year = fastest_asset_start_year,
    value_end_year = fastest_asset_end_year,

    asset_code = fastest_asset_code
  )


# ------------------------------------------------------------
# 8. Combine the four cards for every economy
#
# Visual fields are ready for the master Excel.
# Numeric values and period/source fields remain attached
# for the later profile assembly.
#
# bind_rows fills fields that do not apply to a card with NA.
# ------------------------------------------------------------

kpi_cards <- bind_rows(
  kpi_card1,
  kpi_card2,
  kpi_card3,
  kpi_card4
) %>%
  mutate(
    card_background = "#fff7ed",
    card_border = "#fed7aa"
  ) %>%
  arrange(
    match(geo_code, geo_list),
    kpi_id
  )


# Preview the fields displayed in the reference KPI section.

kpi_cards %>%
  select(
    geo_code,
    kpi_id,
    label,
    value,
    value_colour,
    subtext,
    subtext_colour
  )

```




# I generated seven narrative insights per economy from the calculated indicators and asset rankings. I accounted for economy names, observation periods, contractions, abbreviations and rounding.
```{r profile_key_insights, message=FALSE}

# ============================================================
# CHUNK 14 — GENERATE SEVEN KEY INSIGHTS PER ECONOMY
#
# All economies:
# - Generate text from the calculated results.
# - Use full-precision rankings and Chunk 12's formatting.
# - Keep each indicator's actual observation year/window.
# - Describe negative growth as a contraction.
# - Write "percent" in prose.
# - Reset R&D/IP abbreviation tracking for each economy.
# - Keep seven insights even when an indicator is unavailable.
#
# Economy specificities:
# - India: inherit the approved 10% intangible intensity.
# - Canada: inherit its GDP_Adj-based intensity.
# - Brazil: explicitly identify different tangible periods,
#   including its intensity year when it differs.
# - India, Netherlands and Poland: inherit the asset-growth
#   exclusions from Chunk 8; composition still uses all assets.
# - Canada and Netherlands: inherit configured asset periods.
# - Latvia and Philippines: use their selected composition
#   years and asset-growth periods, without hardcoding dates.
# - Poland: equal displayed rates for the second and third
#   assets can be expressed as "both at ...".
# ============================================================


# ------------------------------------------------------------
# 1. Assemble the inputs for insights 1–4 and 7
#
# Asset insights 5 and 6 use the ranked tables from
# Chunks 8 and 9 directly.
# ------------------------------------------------------------

insight_inputs <- economy_lookup %>%
  select(
    geo_code, economy_name, display_order,
    prose_name, possessive_name, possessive_sentence
  ) %>%
  left_join(
    investment_levels_display %>%
      select(
        geo_code,
        latest_year_intangible,
        intangible_level_prose
      ),
    by = "geo_code"
  ) %>%
  left_join(
    investment_yoy %>%
      select(
        geo_code,
        yoy_start_year_intangible,
        yoy_end_year_intangible,
        intangible_yoy_real,
        yoy_start_year_tangible,
        yoy_end_year_tangible,
        tangible_yoy_real
      ),
    by = "geo_code"
  ) %>%
  left_join(
    investment_cagr %>%
      select(
        geo_code,
        start_year_intangible,
        end_year_intangible,
        intangible_cagr_real,
        start_year_tangible,
        end_year_tangible,
        tangible_cagr_real
      ),
    by = "geo_code"
  ) %>%
  left_join(
    investment_intensity %>%
      select(
        geo_code,
        intensity_year_intangible,
        intangible_intensity_pct,
        intangible_intensity_source,
        intensity_year_tangible,
        tangible_intensity_pct
      ),
    by = "geo_code"
  ) %>%
  left_join(
    measured_unmeasured %>%
      select(geo_code, split_year, unmeasured_pct),
    by = "geo_code"
  ) %>%
  arrange(display_order)


# ------------------------------------------------------------
# 2. Identify the latest headline decade in this dataset
#
# Earlier economy-specific windows receive:
# "Over the past decade for which data are available ...".
#
# Derive the reference window from valid calculated results,
# so the wording updates when a newer dataset is supplied.
# ------------------------------------------------------------

available_cagr_end_years <- investment_cagr %>%
  filter(
    is.finite(intangible_cagr_real),
    is.finite(end_year_intangible)
  ) %>%
  pull(end_year_intangible)

standard_cagr_end_year <- if (length(available_cagr_end_years)) {
  max(available_cagr_end_years)
} else {
  NA_integer_
}

standard_cagr_start_year <- standard_cagr_end_year - 10L


# ------------------------------------------------------------
# 3. Helpers for periods and growth wording
# ------------------------------------------------------------

period_text <- function(start_year, end_year) {
  paste0(
    as.integer(start_year),
    "\u2013",
    as.integer(end_year)
  )
}

same_period <- function(start1, end1, start2, end2) {
  all(is.finite(c(start1, end1, start2, end2))) &&
    isTRUE(start1 == start2) &&
    isTRUE(end1 == end2)
}

has_growth <- function(rate, start_year, end_year, years) {
  all(is.finite(c(rate, start_year, end_year))) &&
    isTRUE(end_year - start_year == years)
}

observation_year_text <- function(year) {
  if (is.finite(year)) {
    paste0(" in ", as.integer(year))
  } else {
    " (observation year unavailable)"
  }
}

growth_clause <- function(rate, annual = FALSE) {

  magnitude <- format_profile_growth(abs(rate), prose = TRUE)

  if (annual) {
    if (rate > 0) {
      return(paste0(
        "grew at a real rate of ", magnitude, " per year"
      ))
    }

    if (rate < 0) {
      return(paste0(
        "contracted at a real rate of ", magnitude, " per year"
      ))
    }

    return("was unchanged in real terms")
  }

  if (rate > 0) return(paste0("grew at ", magnitude))
  if (rate < 0) return(paste0("contracted by ", magnitude))

  "was unchanged"
}

cagr_comparison_text <- function(rate) {
  if (rate < 0) {
    return(paste0(
      "a contraction of ",
      format_profile_growth(abs(rate), prose = TRUE)
    ))
  }

  if (rate == 0) return("no real growth")

  format_profile_growth(rate, prose = TRUE)
}


# ------------------------------------------------------------
# 4. Helper for the second and third ranked assets
#
# Display ties do not change the underlying ranking.
# For growth, identical displayed values can use "both at".
# Composition retains the name-and-percentage construction.
# ------------------------------------------------------------

following_assets_text <- function(
  asset_names,
  values,
  combine_equal = FALSE
) {

  if (length(asset_names) == 0L) return("")

  if (
    combine_equal &&
    length(asset_names) == 2L &&
    identical(values[[1L]], values[[2L]])
  ) {
    return(paste0(
      ", followed by ",
      asset_names[[1L]], " and ", asset_names[[2L]],
      ", both at ", values[[1L]]
    ))
  }

  items <- paste0(asset_names, " (", values, ")")

  paste0(
    ", followed by ",
    paste(items, collapse = " and ")
  )
}


# ------------------------------------------------------------
# 5. Build one economy's seven insights
#
# The abbreviation state exists only inside this function.
# First mention: research & development (R&D).
# Later mentions: R&D.
# The same rule applies to other intellectual property
# (IP) products and subsequent "IP products" mentions.
# ------------------------------------------------------------

build_economy_insights <- function(r) {

  country <- r$prose_name
  possessive <- r$possessive_name
  seen_assets <- character()

  asset_prose <- function(code) {

    first_mention <- !code %in% seen_assets

    result <- if (code == "RD") {
      if (first_mention) "research & development (R&D)" else "R&D"
    } else if (code == "OIPP") {
      if (first_mention) {
        "other intellectual property (IP) products"
      } else {
        "IP products"
      }
    } else {
      tolower(
        asset_lookup$asset_name[
          match(code, asset_lookup$asset_code)
        ]
      )
    }

    seen_assets <<- union(seen_assets, code)

    result
  }


  # ----------------------------------------------------------
  # Insight 1 — Latest intangible investment level
  # ----------------------------------------------------------

  insight1 <- if (
    !is.na(r$intangible_level_prose) &&
    is.finite(r$latest_year_intangible)
  ) {
    paste0(
      "Intangible investment in ", country, " reached ",
      r$intangible_level_prose, " in ",
      as.integer(r$latest_year_intangible), "."
    )
  } else {
    paste0(
      "Intangible investment level data for ", country,
      " are not available in this dataset."
    )
  }


  # ----------------------------------------------------------
  # Insight 2 — Latest real YoY comparison
  #
  # Brazil: retain the actual tangible YoY period when it
  # differs from the intangible period.
  # ----------------------------------------------------------

  intang_yoy_ok <- has_growth(
    r$intangible_yoy_real,
    r$yoy_start_year_intangible,
    r$yoy_end_year_intangible,
    years = 1L
  )

  tang_yoy_ok <- has_growth(
    r$tangible_yoy_real,
    r$yoy_start_year_tangible,
    r$yoy_end_year_tangible,
    years = 1L
  )

  if (intang_yoy_ok) {

    insight2 <- paste0(
      "In the last year (",
      period_text(
        r$yoy_start_year_intangible,
        r$yoy_end_year_intangible
      ),
      "), intangible investment ",
      growth_clause(r$intangible_yoy_real)
    )

    if (tang_yoy_ok) {

      tangible_period <- if (same_period(
        r$yoy_start_year_intangible,
        r$yoy_end_year_intangible,
        r$yoy_start_year_tangible,
        r$yoy_end_year_tangible
      )) {
        ""
      } else {
        paste0(
          " in the most recent year available (",
          period_text(
            r$yoy_start_year_tangible,
            r$yoy_end_year_tangible
          ),
          ")"
        )
      }

      insight2 <- paste0(
        insight2, ", while tangible investment ",
        growth_clause(r$tangible_yoy_real),
        tangible_period, "."
      )

    } else {
      insight2 <- paste0(
        insight2,
        "; real year-on-year growth for tangible investment ",
        "is not available in this dataset."
      )
    }

  } else {

    insight2 <- paste0(
      "Real year-on-year growth data for intangible investment in ",
      country, " are not available in this dataset."
    )

    if (tang_yoy_ok) {
      insight2 <- paste0(
        insight2, " Tangible investment ",
        growth_clause(r$tangible_yoy_real),
        " in the most recent year available (",
        period_text(
          r$yoy_start_year_tangible,
          r$yoy_end_year_tangible
        ),
        ")."
      )
    }
  }


  # ----------------------------------------------------------
  # Insight 3 — Latest 10-year real CAGR comparison
  #
  # Use headline CAGRs from Chunk 5.
  # State the tangible period when the two windows differ.
  # ----------------------------------------------------------

  intang_cagr_ok <- has_growth(
    r$intangible_cagr_real,
    r$start_year_intangible,
    r$end_year_intangible,
    years = 10L
  )

  tang_cagr_ok <- has_growth(
    r$tangible_cagr_real,
    r$start_year_tangible,
    r$end_year_tangible,
    years = 10L
  )

  if (intang_cagr_ok) {

    decade_intro <- if (same_period(
      r$start_year_intangible,
      r$end_year_intangible,
      standard_cagr_start_year,
      standard_cagr_end_year
    )) {
      "Over the past decade"
    } else {
      "Over the past decade for which data are available"
    }

    insight3 <- paste0(
      decade_intro, " (",
      period_text(
        r$start_year_intangible,
        r$end_year_intangible
      ),
      "), intangible investment in ", country, " ",
      growth_clause(r$intangible_cagr_real, annual = TRUE)
    )

    if (tang_cagr_ok) {

      tangible_period <- if (same_period(
        r$start_year_intangible,
        r$end_year_intangible,
        r$start_year_tangible,
        r$end_year_tangible
      )) {
        ""
      } else {
        paste0(
          " over ",
          period_text(
            r$start_year_tangible,
            r$end_year_tangible
          ),
          ", the most recent decade for which tangible ",
          "investment data are available"
        )
      }

      insight3 <- paste0(
        insight3, ", compared to ",
        cagr_comparison_text(r$tangible_cagr_real),
        " recorded for tangible investment",
        tangible_period, "."
      )

    } else {
      insight3 <- paste0(
        insight3,
        "; a 10-year real CAGR for tangible investment ",
        "is not available in this dataset."
      )
    }

  } else {

    insight3 <- paste0(
      "The latest 10-year real CAGR of intangible investment in ",
      country, " is not available in this dataset."
    )

    if (tang_cagr_ok) {
      insight3 <- paste0(
        insight3, " Tangible investment ",
        growth_clause(r$tangible_cagr_real, annual = TRUE),
        " over ",
        period_text(
          r$start_year_tangible,
          r$end_year_tangible
        ),
        "."
      )
    }
  }


  # ----------------------------------------------------------
  # Insight 4 — Latest investment intensity
  #
  # India: reuse its finalized 10%, without recalculation.
  # Canada: reuse its GDP_Adj-based value.
  # Brazil: explicitly show the earlier tangible year.
  # ----------------------------------------------------------

  if (is.finite(r$intangible_intensity_pct)) {

    insight4 <- paste0(
      "Intangible investment in ", country, " makes up ",
      format_profile_percent(
        r$intangible_intensity_pct,
        prose = TRUE
      ),
      " of GDP",
      observation_year_text(r$intensity_year_intangible)
    )

    if (is.finite(r$tangible_intensity_pct)) {

      same_intensity_year <- same_period(
        r$intensity_year_intangible,
        r$intensity_year_intangible,
        r$intensity_year_tangible,
        r$intensity_year_tangible
      )

      insight4 <- paste0(
        insight4, ", while tangible investment makes up ",
        format_profile_percent(
          r$tangible_intensity_pct,
          prose = TRUE
        ),
        if (same_intensity_year) {
          ""
        } else {
          observation_year_text(r$intensity_year_tangible)
        },
        "."
      )

    } else {
      insight4 <- paste0(
        insight4,
        "; tangible investment intensity data are not ",
        "available in this dataset."
      )
    }

  } else {
    insight4 <- paste0(
      "Intangible investment intensity data for ", country,
      " are not available in this dataset."
    )
  }


  # ----------------------------------------------------------
  # Insight 5 — Top three assets by real CAGR
  #
  # Inherit eligibility, periods and ranking from Chunk 8.
  # If no eligible asset grows, describe the highest CAGR
  # without incorrectly claiming positive growth.
  # ----------------------------------------------------------

  growth_assets <- asset_growth_top3 %>%
    filter(geo_code == r$geo_code) %>%
    arrange(asset_rank)

  if (nrow(growth_assets) > 0L) {

    growth_names <- vapply(
      growth_assets$asset_code,
      asset_prose,
      character(1),
      USE.NAMES = FALSE
    )

    growth_labels <- format_profile_growth(
      growth_assets$cagr_real,
      prose = TRUE,
      keep_negative_zero = TRUE
    )

    leading_rate <- growth_assets$cagr_real[1L]

    asset_period <- period_text(
      growth_assets$start_year[1L],
      growth_assets$end_year[1L]
    )

    if (leading_rate > 0) {

      tied_first <- nrow(growth_assets) > 1L &&
        leading_rate == growth_assets$cagr_real[2L]

      opening <- if (tied_first) {
        paste0(
          "One of ", possessive,
          " fastest-growing intangible asset types was "
        )
      } else {
        paste0(
          r$possessive_sentence,
          " fastest-growing intangible asset type was "
        )
      }

      insight5 <- paste0(
        opening, growth_names[1L],
        ", growing at ", growth_labels[1L],
        " between ", asset_period, " in real terms"
      )

    } else if (leading_rate < 0) {

      insight5 <- paste0(
        r$possessive_sentence,
        " smallest real contraction among the eligible ",
        "intangible asset types was in ", growth_names[1L],
        ", declining at ",
        format_profile_growth(abs(leading_rate), prose = TRUE),
        " per year between ", asset_period
      )

    } else {
      insight5 <- paste0(
        r$possessive_sentence,
        " highest intangible asset CAGR was recorded for ",
        growth_names[1L],
        ", with no real growth between ", asset_period
      )
    }

    insight5 <- paste0(
      insight5,
      following_assets_text(
        growth_names[-1L],
        growth_labels[-1L],
        combine_equal = TRUE
      ),
      "."
    )

  } else {
    insight5 <- paste0(
      "Real growth data by intangible asset type for ",
      country, " are not available in this dataset."
    )
  }


  # ----------------------------------------------------------
  # Insight 6 — Top three nominal composition shares
  #
  # Use all seven asset categories through Chunk 9's ranking.
  # Round original share_pct values directly to whole numbers.
  #
  # Abbreviations continue from Insight 5 within this economy.
  # ----------------------------------------------------------

  composition_assets <- asset_composition_top3 %>%
    filter(geo_code == r$geo_code) %>%
    arrange(asset_rank)

  if (nrow(composition_assets) > 0L) {

    composition_names <- vapply(
      composition_assets$asset_code,
      asset_prose,
      character(1),
      USE.NAMES = FALSE
    )

    composition_labels <- format_profile_percent(
      composition_assets$share_pct,
      digits = 0L,
      prose = TRUE
    )

    tied_largest <- nrow(composition_assets) > 1L &&
      composition_assets$share[1L] ==
        composition_assets$share[2L]

    insight6 <- paste0(
      "Investment in ", composition_names[1L],
      if (tied_largest) {
        " makes up one of the largest shares ("
      } else {
        " makes up the largest share ("
      },
      composition_labels[1L],
      ") of ", possessive, " total intangible investment in ",
      as.integer(composition_assets$composition_year[1L]),
      following_assets_text(
        composition_names[-1L],
        composition_labels[-1L]
      ),
      "."
    )

  } else {
    insight6 <- paste0(
      "Intangible asset composition data for ", country,
      " are not available in this dataset."
    )
  }


  # ----------------------------------------------------------
  # Insight 7 — Unmeasured intangible investment
  #
  # Use Chunk 10's own selected year, independently of the
  # composition year or other headline observation years.
  # ----------------------------------------------------------

  insight7 <- if (
    is.finite(r$unmeasured_pct) &&
    is.finite(r$split_year)
  ) {
    paste0(
      format_profile_percent(r$unmeasured_pct, prose = TRUE),
      " of ", possessive,
      " total intangible investment goes unmeasured in ",
      "official statistics, as of ",
      as.integer(r$split_year), "."
    )
  } else {
    paste0(
      "The measured and unmeasured split of intangible ",
      "investment in ", country,
      " is not available in this dataset."
    )
  }


  # One row per insight, in the required order.
  tibble(
    geo_code = r$geo_code,
    economy_name = r$economy_name,
    insight_id = seq_len(7L),
    text = c(
      insight1, insight2, insight3, insight4,
      insight5, insight6, insight7
    )
  )
}


# ------------------------------------------------------------
# 6. Generate the seven insights for every economy
#
# Keep text without a bullet prefix for the Excel table.
# The final layout can add bullets when displaying the text.
# ------------------------------------------------------------

key_insights <- bind_rows(
  lapply(
    seq_len(nrow(insight_inputs)),
    function(i) {
      build_economy_insights(insight_inputs[i, , drop = FALSE])
    }
  )
) %>%
  arrange(
    match(geo_code, geo_list),
    insight_id
  )


key_insights

```





# I prepared the numerical tables, display labels and colours for Figures 1–6, covering intensity, indexed investment, aggregate growth, asset growth, composition and measured/unmeasured shares. I carried forward the economy-specific calculation rules.

```{r country_figure_values, message=FALSE}

# ============================================================
# CHUNK 15 — PREPARE VALUES FOR COUNTRY FIGURES 1–6
#
# Figure 1: Investment intensity over time.
# Figure 2: Indexed real investment.
# Figure 3: Intangible/tangible real CAGR over the comparison period.
# Figure 4: Real CAGR by intangible asset.
# Figure 5: Nominal intangible asset composition.
# Figure 6: Measured versus unmeasured intangible investment.
#
# Economy specificities inherited from earlier calculations:
# - India: latest intangible intensity includes the 10% override.
# - Canada: intensity uses GDP_Adj.
# - Brazil: Figures 2–3 use the configured 2011–2021 window.
# - Brazil and Poland: missing timeline observations remain NA.
# - India, Netherlands and Poland: OIPP excluded from Figure 4,
#   but retained in Figure 5.
# - Canada and Netherlands: retain configured asset-growth windows.
# - Japan: retain small negative growth values and negative labels.
#
# Original calculation tables remain available at full precision.
# These new tables prepare the display values for the master Excel.
# ============================================================


# ------------------------------------------------------------
# 1. Shared series styles for Figures 1–3
#
# Figure 3's intangible value label is black.
# Its tangible value label matches the orange bar.
# ------------------------------------------------------------

figure_series_style <- tibble::tribble(
  ~series_order, ~series_label, ~nominal_variable, ~real_variable,
  ~series_colour, ~label_colour,

  1L, "Intangible", "I_Intang", "Iq_Intang",
  "#1e3a5f", "#1c1c1c",

  2L, "Tangible", "I_TangNRes", "Iq_TangNRes",
  "#ea580c", "#ea580c"
)


# ------------------------------------------------------------
# 2. Asset palettes
#
# Colours follow display order after ranking full-precision
# values. Negative asset-growth bars override the palette
# with grey.
# ------------------------------------------------------------

asset_growth_palette <- c(
  "#9a3412",
  "#c2410c",
  "#ea580c",
  "#f97316",
  "#fb923c",
  "#fdba74",
  "#fed7aa"
)

asset_composition_palette <- c(
  "#9a3412",
  "#c2410c",
  "#ea580c",
  "#f97316",
  "#fdba74",
  "#fed7aa",
  "#f59e0b"
)


# ------------------------------------------------------------
# 3. Growth values for chart tables
#
# Normally store one decimal place.
#
# Japan's Design example:
# A small negative rate must remain negative even when its
# one-decimal label is "-0.0%". Preserve the original numeric
# percentage in that case.
#
# Apply this rule to any economy with the same situation.
# ------------------------------------------------------------

chart_growth_value <- function(pct) {

  result <- round_profile(pct, digits = 1L)

  preserve_negative <- (
    is.finite(pct) &
    pct < 0 &
    !is.na(result) &
    result == 0
  )

  result[preserve_negative] <- pct[preserve_negative]

  result
}


# ------------------------------------------------------------
# 4. Figure 1 — Investment intensity over time
#
# Intensity is already expressed in percentage units.
# Preserve the completed annual timeline from Chunk 7.
#
# NA values remain numeric NA here. The Excel export will
# handle their final representation as missing chart points.
# ------------------------------------------------------------

figure1_values <- intensity_history %>%
  transmute(
    geo_code,
    year,
    series = factor(
      series,
      levels = c("I_Intang", "I_TangNRes"),
      labels = c(
        "Intangible (% of GDP)",
        "Tangible (% of GDP)"
      )
    ),
    value = round_profile(intensity_pct, digits = 1L)
  ) %>%
  pivot_wider(
    id_cols = c(geo_code, year),
    names_from = series,
    values_from = value,
    names_expand = TRUE
  ) %>%
  arrange(
    match(geo_code, geo_list),
    year
  )


# ------------------------------------------------------------
# 5. Figure 2 — Indexed real investment
#
# Use the existing indices from Chunk 6.
# Keep their original base year and missing observations.
# No rebasing or interpolation is performed here.
# ------------------------------------------------------------

figure2_values <- indexed_real_series %>%
  transmute(
    geo_code,
    year,
    series = factor(
      series,
      levels = c("Iq_Intang", "Iq_TangNRes"),
      labels = c(
        "Intangible (index)",
        "Tangible (index)"
      )
    ),
    value = round_profile(index_100, digits = 1L)
  ) %>%
  pivot_wider(
    id_cols = c(geo_code, year),
    names_from = series,
    values_from = value,
    names_expand = TRUE
  ) %>%
  arrange(
    match(geo_code, geo_list),
    year
  )


# ------------------------------------------------------------
# 6. Figure 3 — Real investment growth over the comparison period
#
# Use Chunk 6's comparison CAGR.
#
# Brazil: this preserves the common 2011–2021 comparison.
# Its headline CAGR remains separately available for the
# KPI, key insight and global Figure 10.
# ------------------------------------------------------------

figure3_values <- investment_comparison_cagr %>%
  left_join(
    figure_series_style,
    by = c("series" = "real_variable")
  ) %>%
  mutate(
    cagr_pct_raw = cagr_real * 100,
    cagr_pct = chart_growth_value(cagr_pct_raw),

    value_label = format_profile_percent(
      cagr_pct_raw,
      keep_negative_zero = TRUE
    )
  ) %>%
  select(
    geo_code,
    row_order = series_order,
    series = series_label,
    start_year,
    end_year,
    cagr_pct_raw,
    cagr_pct,
    value_label,
    bar_colour = series_colour,
    label_colour
  ) %>%
  arrange(
    match(geo_code, geo_list),
    row_order
  )


# ------------------------------------------------------------
# 7. Figure 4 — Real CAGR by intangible asset
#
# Rank on unrounded growth, then assign colours.
# Missing eligible asset results remain NA and appear last.
#
# India, Netherlands and Poland:
# Use the six eligible categories already selected in Chunk 8.
#
# Japan and other economies with negative asset growth:
# Grey bars and grey value labels; preserve the negative sign.
# ------------------------------------------------------------

figure4_values <- asset_cagr_real %>%
  group_by(geo_code) %>%
  arrange(
    desc(cagr_real),
    asset_order,
    .by_group = TRUE
  ) %>%
  mutate(
    row_order = row_number(),

    cagr_pct_raw = cagr_real * 100,
    cagr_pct = chart_growth_value(cagr_pct_raw),

    value_label = format_profile_percent(
      cagr_pct_raw,
      keep_negative_zero = TRUE
    ),

    bar_colour = case_when(
      !is.finite(cagr_real) ~ "#cbc9c4",
      cagr_real < 0        ~ "#78716c",
      TRUE                 ~ asset_growth_palette[row_order]
    ),

    label_colour = bar_colour
  ) %>%
  ungroup() %>%
  select(
    geo_code,
    row_order,
    asset_code,
    asset_name,
    start_year,
    end_year,
    cagr_pct_raw,
    cagr_pct,
    value_label,
    bar_colour,
    label_colour
  ) %>%
  arrange(
    match(geo_code, geo_list),
    row_order
  )


# ------------------------------------------------------------
# 8. Figure 5 — Nominal intangible asset composition
#
# Keep all seven categories for every economy.
# Rank using unrounded shares from Chunk 9.
#
# Display rule:
# - Numeric shares: one decimal place.
# - White slice labels only when the unrounded share is >=3%.
# - All categories remain available for the legend.
#
# Each economy retains its selected composition year.
# Rounded shares are not renormalized a second time.
# ------------------------------------------------------------

figure5_values <- asset_composition %>%
  group_by(geo_code) %>%
  arrange(
    desc(share),
    asset_order,
    .by_group = TRUE
  ) %>%
  mutate(
    row_order = row_number(),

    share_pct_raw = share_pct,
    share_pct_display = round_profile(
      share_pct_raw,
      digits = 1L
    ),

    slice_colour = asset_composition_palette[row_order],
    label_colour = "#ffffff",

    label_shown = if_else(
      is.finite(share_pct_raw) & share_pct_raw >= 3,
      "Yes",
      "No"
    ),

    value_label = if_else(
      label_shown == "Yes",
      format_profile_percent(share_pct_raw),
      NA_character_
    )
  ) %>%
  ungroup() %>%
  select(
    geo_code,
    row_order,
    asset_code,
    asset_name,
    year = composition_year,
    share_pct_raw,
    share_pct = share_pct_display,
    slice_colour,
    label_colour,
    label_shown,
    value_label
  ) %>%
  arrange(
    match(geo_code, geo_list),
    row_order
  )


# ------------------------------------------------------------
# 9. Figure 6 — Measured versus unmeasured investment
#
# Preserve the reference's fixed slice order:
# 1. Unmeasured: grey fill, dark text.
# 2. Measured: navy fill, white text.
#
# Both available slices receive percentage labels.
# Use this indicator's own latest valid year from Chunk 10.
# ------------------------------------------------------------

figure6_values <- measured_unmeasured_long %>%
  mutate(
    row_order = match(
      component,
      c("Unmeasured", "Measured")
    ),

    share_pct_raw = share_pct,
    share_pct_display = round_profile(
      share_pct_raw,
      digits = 1L
    ),

    fill_colour = if_else(
      component == "Unmeasured",
      "#d6d3d1",
      "#1e3a5f"
    ),

    label_colour = if_else(
      component == "Unmeasured",
      "#44403c",
      "#ffffff"
    ),

    label_shown = if_else(
      is.finite(share_pct_raw),
      "Yes",
      "No"
    ),

    value_label = format_profile_percent(share_pct_raw)
  ) %>%
  select(
    geo_code,
    row_order,
    component,
    year = split_year,
    share_pct_raw,
    share_pct = share_pct_display,
    fill_colour,
    label_colour,
    label_shown,
    value_label
  ) %>%
  arrange(
    match(geo_code, geo_list),
    row_order
  )


# ------------------------------------------------------------
# 10. Collect the six tables for later profile assembly
#
# Each table contains geo_code, allowing the export to
# extract the appropriate rows for each economy's sheet.
# ------------------------------------------------------------

country_figure_values <- list(
  figure1 = figure1_values,
  figure2 = figure2_values,
  figure3 = figure3_values,
  figure4 = figure4_values,
  figure5 = figure5_values,
  figure6 = figure6_values
)


country_figure_values

```



# I generated the titles, subtitles, periods, axis ticks and chart instructions for Figures 1–6. I specified details including (UK) and (US) titles, missing-data gaps, asset exclusions and label rules.
```{r country_figure_metadata, message=FALSE}

# ============================================================
# CHUNK 16 — METADATA FOR COUNTRY FIGURES 1–6
#
# All economies:
# - Derive chart periods from the preceding calculation tables.
# - Generate titles, subtitles, ranges and axis ticks.
# - Preserve missing observations as gaps in line charts.
# - Record doughnut cutouts and label rules explicitly.
#
# Economy specificities:
# - US: use "(US)" in figure titles.
# - UK: use "(UK)" in figure titles.
# - Canada: preserve the reference's final-tick spacing.
# - Brazil and Poland: identify shorter tangible coverage
#   when it differs from the overall intensity timeline.
# - Brazil: describe Figure 2 as a common comparison period.
# - India, Netherlands and Poland: identify asset exclusions.
# - Japan: document small negative growth labels dynamically.
# ============================================================


# ------------------------------------------------------------
# 1. Helpers for observed ranges and selected periods
# ------------------------------------------------------------

observed_year_range <- function(years) {

  years <- years[is.finite(years)]

  if (length(years) == 0L) {
    return(c(start = NA_integer_, end = NA_integer_))
  }

  c(
    start = min(years),
    end = max(years)
  )
}

chart_range_text <- function(period) {

  if (!all(is.finite(period))) {
    return("Not available")
  }

  if (period[1L] == period[2L]) {
    return(as.character(as.integer(period[1L])))
  }

  period_text(period[1L], period[2L])
}


# A figure with a common period must have one distinct pair.
# Stop if inconsistent periods would produce a misleading
# single subtitle.

selected_chart_period <- function(
  df,
  start_column,
  end_column,
  context
) {

  periods <- df %>%
    transmute(
      start = .data[[start_column]],
      end = .data[[end_column]]
    ) %>%
    filter(
      is.finite(start),
      is.finite(end)
    ) %>%
    distinct()

  if (nrow(periods) == 0L) {
    return(c(start = NA_integer_, end = NA_integer_))
  }

  if (
    nrow(periods) > 1L ||
    any(periods$end < periods$start)
  ) {
    stop(
      "Inconsistent chart period: ",
      context,
      ". Check the preceding calculation tables.",
      call. = FALSE
    )
  }

  c(
    start = periods$start[1L],
    end = periods$end[1L]
  )
}


# ------------------------------------------------------------
# 2. Generate year ticks
#
# Figure 1: five-year intervals plus the final year.
# Figure 2: two-year intervals plus the final year.
#
# Canada, Figure 1:
# When the final year falls between regular ticks, replace
# the last regular tick with the final year.
#
# This reproduces the reference's terminal spacing without
# hardcoding its historical list of years.
# ------------------------------------------------------------

chart_year_ticks <- function(
  period,
  step,
  replace_last_tick = FALSE
) {

  if (
    !all(is.finite(period)) ||
    period[2L] < period[1L]
  ) {
    return(integer())
  }

  ticks <- seq.int(
    from = period[1L],
    to = period[2L],
    by = step
  )

  if (
    replace_last_tick &&
    length(ticks) > 1L &&
    tail(ticks, 1L) != period[2L]
  ) {
    ticks <- head(ticks, -1L)
  }

  as.integer(sort(unique(c(ticks, period[2L]))))
}

ticks_as_text <- function(ticks) {
  if (length(ticks) == 0L) {
    "Not available"
  } else {
    paste(ticks, collapse = ", ")
  }
}

has_chart_values <- function(df, columns) {
  any(vapply(
    df[columns],
    function(x) any(is.finite(x)),
    logical(1)
  ))
}


# ------------------------------------------------------------
# 3. Shared figure titles
#
# Economy suffixes are added below using figure_label.
# This preserves "(US)" and "(UK)", while other economies
# use their full names in parentheses.
# ------------------------------------------------------------

country_figure_titles <- c(
  "Investment intensity over time",
  "Intangible versus tangible investment growth",
  "Real growth of intangible & tangible investment",
  "Real growth by asset type",
  "Intangible asset composition",
  "Measured versus unmeasured intangible investment"
)


# ------------------------------------------------------------
# 4. Build one economy's figure metadata
# ------------------------------------------------------------

build_country_figure_metadata <- function(g) {

  economy <- economy_lookup %>%
    filter(geo_code == g)

  figures <- lapply(
    country_figure_values,
    function(df) filter(df, geo_code == g)
  )

  comparison <- comparison_windows %>%
    filter(geo_code == g)


  # ----------------------------------------------------------
  # Actual periods for each figure
  # ----------------------------------------------------------

  periods <- list(
    observed_year_range(figures$figure1$year),

    selected_chart_period(
      comparison,
      "start_year", "end_year",
      paste(g, "Figure 2")
    ),

    selected_chart_period(
      figures$figure3,
      "start_year", "end_year",
      paste(g, "Figure 3")
    ),

    selected_chart_period(
      figures$figure4,
      "start_year", "end_year",
      paste(g, "Figure 4")
    ),

    selected_chart_period(
      figures$figure5,
      "year", "year",
      paste(g, "Figure 5")
    ),

    selected_chart_period(
      figures$figure6,
      "year", "year",
      paste(g, "Figure 6")
    )
  )

  ranges <- vapply(
    periods,
    chart_range_text,
    character(1)
  )

  has_data <- c(
    has_chart_values(
      figures$figure1,
      c("Intangible (% of GDP)", "Tangible (% of GDP)")
    ),

    has_chart_values(
      figures$figure2,
      c("Intangible (index)", "Tangible (index)")
    ),

    has_chart_values(figures$figure3, "cagr_pct_raw"),
    has_chart_values(figures$figure4, "cagr_pct_raw"),
    has_chart_values(figures$figure5, "share_pct_raw"),
    has_chart_values(figures$figure6, "share_pct_raw")
  )


  # ----------------------------------------------------------
  # Figure 1: identify different series coverage
  #
  # Brazil: tangible observations may end earlier.
  # Poland: tangible observations may begin later.
  #
  # Apply the same check to every economy in future editions.
  # ----------------------------------------------------------

  coverage_notes <- character()

  for (series_code in c("I_Intang", "I_TangNRes")) {

    observations <- intensity_history %>%
      filter(
        geo_code == g,
        series == series_code,
        is.finite(intensity_pct)
      )

    series_period <- observed_year_range(observations$year)

    series_name <- if (series_code == "I_Intang") {
      "intangible"
    } else {
      "tangible"
    }

    if (!all(is.finite(series_period))) {

      coverage_notes <- c(
        coverage_notes,
        paste(series_name, "unavailable")
      )

    } else if (!same_period(
      series_period[1L],
      series_period[2L],
      periods[[1L]][1L],
      periods[[1L]][2L]
    )) {

      coverage_notes <- c(
        coverage_notes,
        paste(series_name, chart_range_text(series_period))
      )
    }
  }

  if (has_data[1L] && length(coverage_notes) > 0L) {
    ranges[1L] <- paste0(
      ranges[1L],
      " (",
      paste(coverage_notes, collapse = "; "),
      ")"
    )
  }


  # ----------------------------------------------------------
  # Dynamic axis ticks
  # ----------------------------------------------------------

  intensity_ticks <- chart_year_ticks(
    periods[[1L]],
    step = 5L,
    replace_last_tick = g == "CA"
  )

  index_ticks <- chart_year_ticks(
    periods[[2L]],
    step = 2L
  )


  # ----------------------------------------------------------
  # Subtitles
  #
  # Brazil's common comparison window is explicitly described
  # without calling it the latest intangible-only decade.
  # ----------------------------------------------------------

  index_subtitle <- if (all(is.finite(periods[[2L]]))) {
    paste0(
      "Indexed ",
      as.integer(periods[[2L]][1L]),
      " = 100, ",
      if (g == "BR") {
        "common comparison period"
      } else {
        "latest 10-year period available"
      }
    )
  } else {
    "No data available"
  }

  subtitles <- c(
    paste0(
      "Share of GDP, ",
      chart_range_text(periods[[1L]])
    ),

    index_subtitle,

    paste0(
      "10-year CAGR, ",
      chart_range_text(periods[[3L]])
    ),

    paste0(
      "10-year CAGR, ",
      chart_range_text(periods[[4L]])
    ),

    paste0(
      "Share of intangible investment, ",
      chart_range_text(periods[[5L]])
    ),

    paste0(
      "Share of intangible investment not covered in ",
      "National Accounts, ",
      chart_range_text(periods[[6L]])
    )
  )

  subtitles[!has_data] <- "No data available"


  # ----------------------------------------------------------
  # Chart/display instructions
  # ----------------------------------------------------------

  chart_display <- c(
    paste0(
      "Two lines: Intangible #1e3a5f, Tangible #ea580c. ",
      "Y-axis from 0, %. X ticks: ",
      ticks_as_text(intensity_ticks),
      ". Missing observations remain gaps."
    ),

    paste0(
      "Two lines: Intangible #1e3a5f, Tangible #ea580c. ",
      "Index. X ticks: ",
      ticks_as_text(index_ticks),
      ". Missing observations remain gaps."
    ),

    paste0(
      "Vertical bars: Intangible #1e3a5f, Tangible #ea580c. ",
      "Value labels: intangible black, tangible orange. ",
      "Use a zero baseline and include negative values."
    ),

    paste0(
      "Horizontal bars in the supplied row order; ",
      "palette by row; negative bars grey #78716c; ",
      "value labels in bar colour. ",
      "Use the supplied value_label text."
    ),

    paste0(
      "Doughnut, 50% cutout, legend below. ",
      "Keep all seven asset categories. ",
      "White value labels only for slices with ",
      "an unrounded share of at least 3%."
    ),

    paste0(
      "Doughnut, 55% cutout, legend below. ",
      "Unmeasured first: #d6d3d1, label #44403c. ",
      "Measured second: #1e3a5f, label #ffffff. ",
      "Both available slices labelled."
    )
  )


  # ----------------------------------------------------------
  # Figure 4: document eligible categories
  #
  # India, Netherlands and Poland inherit OIPP exclusions.
  # Determine exclusions from the actual figure table.
  # ----------------------------------------------------------

  excluded_codes <- setdiff(
    asset_lookup$asset_code,
    figures$figure4$asset_code
  )

  chart_display[4L] <- paste0(
    chart_display[4L],
    " Eligible asset categories: ",
    nrow(figures$figure4),
    "."
  )

  if (length(excluded_codes) > 0L) {

    excluded_names <- asset_lookup$asset_name[
      match(excluded_codes, asset_lookup$asset_code)
    ]

    chart_display[4L] <- paste0(
      chart_display[4L],
      " Excluded: ",
      paste(excluded_names, collapse = ", "),
      "."
    )
  }


  # ----------------------------------------------------------
  # Figure 4: document negative values rounding to "-0.0%"
  #
  # Covers Japan's Design case and any future equivalent.
  # Preserve the numeric negative value from Chunk 15.
  # ----------------------------------------------------------

  small_negative_assets <- figures$figure4 %>%
    filter(
      is.finite(cagr_pct_raw),
      cagr_pct_raw < 0,
      round_profile(cagr_pct_raw, digits = 1L) == 0
    )

  if (nrow(small_negative_assets) > 0L) {
    chart_display[4L] <- paste0(
      chart_display[4L],
      " Small negative labels: ",
      paste(
        paste(
          small_negative_assets$asset_name,
          small_negative_assets$value_label
        ),
        collapse = ", "
      ),
      "; retain the underlying negative numeric values."
    )
  }


  # ----------------------------------------------------------
  # Return six metadata rows
  #
  # has_data is used internally when assembling the workbook.
  # Structured fields also support later HTML reproduction.
  # ----------------------------------------------------------

  tibble(
    geo_code = g,
    economy_name = economy$economy_name,
    figure_id = seq_len(6L),

    # figure_label already includes parentheses.
    # US -> "(US)"; UK -> "(UK)"; others -> full names.
    title = paste0(
      country_figure_titles,
      " ",
      economy$figure_label
    ),

    subtitle = subtitles,
    chart_display = chart_display,
    range = ranges,

    start_year = vapply(
      periods, function(x) unname(x[1L]), numeric(1)
    ),

    end_year = vapply(
      periods, function(x) unname(x[2L]), numeric(1)
    ),

    chart_type = c(
      "Line", "Line", "Vertical bar",
      "Horizontal bar", "Doughnut", "Doughnut"
    ),

    unit = c(
      "Percent of GDP",
      "Index",
      "Percent per year",
      "Percent per year",
      "Percent of intangible investment",
      "Percent of intangible investment"
    ),

    x_ticks = c(
      ticks_as_text(intensity_ticks),
      ticks_as_text(index_ticks),
      rep(NA_character_, 4L)
    ),

    y_axis_min = c(0, rep(NA_real_, 5L)),

    cutout_pct = c(
      rep(NA_real_, 4L),
      50,
      55
    ),

    legend_position = c(
      rep(NA_character_, 4L),
      "below",
      "below"
    ),

    connect_missing_points = c(
      FALSE,
      FALSE,
      rep(NA, 4L)
    ),

    has_data = has_data
  )
}


# ------------------------------------------------------------
# 5. Build metadata for all economies
# ------------------------------------------------------------

country_figure_metadata <- bind_rows(
  lapply(geo_list, build_country_figure_metadata)
) %>%
  arrange(
    match(geo_code, geo_list),
    figure_id
  )


country_figure_metadata %>%
  select(
    geo_code,
    figure_id,
    title,
    subtitle,
    chart_display,
    range
  )

```



# I formatted Figures 7–10 and generated their titles, subtitles and display instructions. I created a version for each profile that highlights its economy in navy while keeping other economies grey.
```{r global_figure_values_and_metadata, message=FALSE}

# ============================================================
# CHUNK 17 — GLOBAL FIGURES 7–10
#
# Figure 7: Top five intangible investment levels.
# Figure 8: Remaining intangible investment levels.
# Figure 9: Latest intangible investment intensity.
# Figure 10: Latest 10-year real intangible investment CAGR.
#
# All economies:
# - Preserve the full-precision ordering from Chunk 11.
# - Use Chunk 12's unit conversion and display formatting.
# - Derive subtitle years and exceptions from the actual rows.
# - Highlight the profile economy in navy with a bold label.
# - Use grey for every other economy.
# - If the profile economy is absent from a figure, all bars
#   remain grey.
#
# Economy specificities:
# - India: Figure 9 retains its accepted 10% intensity.
# - Canada: Figure 9 retains the GDP_Adj-based intensity.
# - Brazil: Figure 10 uses its headline CAGR from Chunk 5,
#   independently of the comparison period in Figures 2–3.
# - US: numeric levels remain in billions; the label can use T.
#
# Output:
# - Shared values for each global figure.
# - Profile-specific copies carrying highlight instructions.
# - Four metadata rows per economy.
# ============================================================


# ------------------------------------------------------------
# 1. Global chart colours
# ------------------------------------------------------------

global_plot_colours <- c(
  highlight = "#1e3a5f",
  other     = "#cbc9c4"
)


# ------------------------------------------------------------
# 2. Prepare investment level tables
#
# Numeric chart values are consistently in USD billions.
# Display labels use B or T according to the amount.
#
# Restart row numbering within each panel:
# the first economy in Figure 8 is row 1 of that figure.
# ------------------------------------------------------------

prepare_global_level_values <- function(df) {

  df %>%
    arrange(level_position) %>%
    transmute(
      geo_code,
      economy_name,
      row_order = row_number(),
      year = latest_year_intangible,

      usd_billions_raw = ppp_to_billions(intangible_raw),

      usd_billions = round_profile(
        usd_billions_raw,
        digits = 1L
      ),

      value_label = format_profile_level(
        intangible_raw,
        style = "axis"
      )
    )
}

figure7_values <- prepare_global_level_values(
  figure7_top5_levels
)

figure8_values <- prepare_global_level_values(
  figure8_remaining_levels
)


# ------------------------------------------------------------
# 3. Figure 9 — Latest intangible intensity
#
# Values are already in percentage units.
# Retain the recorded source, including India's manual input
# and Canada's GDP_Adj denominator.
# ------------------------------------------------------------

figure9_values <- figure9_intensity_all %>%
  arrange(intensity_position) %>%
  transmute(
    geo_code,
    economy_name,
    row_order = row_number(),
    year = intensity_year_intangible,

    intensity_pct_raw = intangible_intensity_pct,

    intensity_pct = round_profile(
      intensity_pct_raw,
      digits = 1L
    ),

    value_label = format_profile_percent(
      intensity_pct_raw
    ),

    value_source = intangible_intensity_source
  )


# ------------------------------------------------------------
# 4. Figure 10 — Latest real intangible investment CAGR
#
# Chunk 11 already converted this rate into percentage units.
# Keep each economy's own start and end years.
#
# Preserve small negative numeric values and signed labels
# using the helper introduced in Chunk 15.
# ------------------------------------------------------------

figure10_values <- figure10_cagr_all %>%
  arrange(cagr_position) %>%
  transmute(
    geo_code,
    economy_name,
    row_order = row_number(),

    start_year = start_year_intangible,
    end_year = end_year_intangible,

    cagr_pct_raw = intangible_cagr_pct,
    cagr_pct = chart_growth_value(cagr_pct_raw),

    value_label = format_profile_percent(
      cagr_pct_raw,
      keep_negative_zero = TRUE
    )
  )


# Shared values: these are the same for every profile.

global_figure_values <- list(
  figure7  = figure7_values,
  figure8  = figure8_values,
  figure9  = figure9_values,
  figure10 = figure10_values
)


# ------------------------------------------------------------
# 5. Helper for readable lists of economy exceptions
# ------------------------------------------------------------

join_prose_items <- function(x) {

  x <- x[!is.na(x) & nzchar(x)]

  if (length(x) == 0L) return("")
  if (length(x) == 1L) return(x)
  if (length(x) == 2L) return(paste(x, collapse = " and "))

  paste0(
    paste(head(x, -1L), collapse = ", "),
    ", and ",
    tail(x, 1L)
  )
}


# ------------------------------------------------------------
# 6. Dynamic subtitles for Figures 7 and 8
#
# Describe the observation years actually present in each
# panel. Do not permanently assign an economy to the top five
# or hardcode its latest year.
# ------------------------------------------------------------

global_levels_subtitle <- function(df) {

  if (nrow(df) == 0L) {
    return("No data available")
  }

  known_years <- df$year[is.finite(df$year)]

  if (length(known_years) == 0L) {
    return(
      paste0(
        "Absolute levels, observation years unavailable",
        " \u2014 in USD billions"
      )
    )
  }

  reference_year <- max(known_years)

  exceptions <- df %>%
    filter(
      !is.finite(year) |
      year != reference_year
    ) %>%
    arrange(match(geo_code, geo_list)) %>%
    mutate(
      exception_text = if_else(
        is.finite(year),
        paste0(geo_code, " (", as.integer(year), ")"),
        paste0(geo_code, " (year unavailable)")
      )
    )

  if (nrow(exceptions) == 0L) {
    return(paste0(
      "Absolute levels, latest available year (",
      as.integer(reference_year),
      ") \u2014 in USD billions"
    ))
  }

  paste0(
    "Absolute levels, latest available year for all economies (",
    as.integer(reference_year),
    "), except ",
    join_prose_items(exceptions$exception_text),
    " \u2014 in USD billions"
  )
}


# ------------------------------------------------------------
# 7. Dynamic subtitle for Figure 10
#
# Identify the latest decade represented in the chart and
# explicitly list economies using a different decade.
#
# Brazil: the listed period comes from its headline CAGR.
# ------------------------------------------------------------

global_cagr_subtitle <- function(df) {

  if (nrow(df) == 0L) {
    return("No data available")
  }

  reference_end <- max(df$end_year)
  reference_start <- reference_end - 10L

  exceptions <- df %>%
    filter(
      start_year != reference_start |
      end_year != reference_end
    ) %>%
    arrange(match(geo_code, geo_list)) %>%
    mutate(
      exception_text = paste0(
        geo_code,
        " (",
        period_text(start_year, end_year),
        ")"
      )
    )

  subtitle <- paste0(
    "CAGR for ",
    period_text(reference_start, reference_end),
    " for all economies"
  )

  if (nrow(exceptions) > 0L) {
    subtitle <- paste0(
      subtitle,
      ", except ",
      join_prose_items(exceptions$exception_text)
    )
  }

  subtitle
}


# ------------------------------------------------------------
# 8. Figure 10 axis limits
#
# Preserve the reference's -3% to 12% range when it contains
# all values. Expand either boundary when updated data would
# otherwise fall outside the chart.
# ------------------------------------------------------------

global_cagr_axis <- c(
  lower = -3,
  upper = 12
)

available_global_cagr <- figure10_values$cagr_pct_raw
available_global_cagr <- available_global_cagr[
  is.finite(available_global_cagr)
]

if (length(available_global_cagr) > 0L) {

  global_cagr_axis[["lower"]] <- min(
    global_cagr_axis[["lower"]],
    floor(min(available_global_cagr))
  )

  global_cagr_axis[["upper"]] <- max(
    global_cagr_axis[["upper"]],
    ceiling(max(available_global_cagr))
  )
}


# ------------------------------------------------------------
# 9. Shared global figure metadata
#
# Highlight instructions are added separately for each profile.
# ------------------------------------------------------------

global_figure_metadata_base <- tibble(
  figure_id = 7:10,

  title = c(
    "Top 5 economies by intangible investment",
    "Remaining economies by intangible investment",
    "Intangible investment intensity",
    "Real growth of intangible investment"
  ),

  subtitle = c(
    global_levels_subtitle(figure7_values),

    global_levels_subtitle(figure8_values),

    if (nrow(figure9_values) > 0L) {
      "Share of GDP, latest available year"
    } else {
      "No data available"
    },

    global_cagr_subtitle(figure10_values)
  ),

  chart_type = "Horizontal bar",

  unit = c(
    "USD billions",
    "USD billions",
    "Percent of GDP",
    "Percent per year"
  ),

  chart_display = c(
    paste0(
      "Horizontal bars in supplied row order. ",
      "Numeric values in USD billions; labels use B, ",
      "or T for amounts of at least 1,000 billion."
    ),

    paste0(
      "Horizontal bars in supplied row order. ",
      "Numeric values in USD billions; labels use B, ",
      "or T for amounts of at least 1,000 billion."
    ),

    paste0(
      "Horizontal bars in supplied row order. ",
      "Unit: % of GDP. Use the supplied value labels."
    ),

    paste0(
      "Horizontal bars in supplied row order. ",
      "Use the supplied signed percentage labels. X-axis ",
      format_profile_percent(global_cagr_axis[["lower"]]),
      " to ",
      format_profile_percent(global_cagr_axis[["upper"]]),
      "."
    )
  ),

  x_axis_min = c(
    0,
    0,
    0,
    global_cagr_axis[["lower"]]
  ),

  x_axis_max = c(
    NA_real_,
    NA_real_,
    NA_real_,
    global_cagr_axis[["upper"]]
  ),

  has_data = vapply(
    global_figure_values,
    function(df) nrow(df) > 0L,
    logical(1)
  )
)


# ------------------------------------------------------------
# 10. Add highlight instructions to each profile's values
#
# geo_code identifies the economy represented by a bar.
# profile_geo_code identifies the profile containing the chart.
#
# The underlying values and row order remain shared.
# ------------------------------------------------------------

global_figure_values_by_profile <- setNames(
  lapply(
    geo_list,
    function(g) {

      lapply(
        global_figure_values,
        function(df) {

          df %>%
            mutate(
              profile_geo_code = g,
              highlighted = geo_code == g,

              bar_colour = if_else(
                highlighted,
                global_plot_colours[["highlight"]],
                global_plot_colours[["other"]]
              ),

              label_weight = if_else(
                highlighted,
                "bold",
                "normal"
              )
            ) %>%
            relocate(profile_geo_code)
        }
      )
    }
  ),
  geo_list
)


# ------------------------------------------------------------
# 11. Add profile-specific highlight descriptions
#
# If an economy is absent from a panel, state explicitly that
# all bars are grey. This covers economies outside the top five
# and economies without an available result for an indicator.
# ------------------------------------------------------------

build_global_figure_metadata <- function(g) {

  highlight_present <- vapply(
    global_figure_values,
    function(df) g %in% df$geo_code,
    logical(1)
  )

  global_figure_metadata_base %>%
    mutate(
      geo_code = g,

      highlighted_geo_code = if_else(
        highlight_present,
        g,
        NA_character_
      ),

      chart_display = paste(
        chart_display,
        case_when(
          !has_data ~ "No data available.",

          highlight_present ~ paste0(
            g,
            " navy #1e3a5f with a bold label; ",
            "other economies grey #cbc9c4."
          ),

          TRUE ~ paste0(
            "All bars grey #cbc9c4; ",
            g,
            " is not shown in this figure."
          )
        )
      )
    ) %>%
    relocate(geo_code)
}

global_figure_metadata <- bind_rows(
  lapply(geo_list, build_global_figure_metadata)
) %>%
  left_join(
    economy_lookup %>%
      select(geo_code, economy_name),
    by = "geo_code"
  ) %>%
  relocate(economy_name, .after = geo_code) %>%
  arrange(
    match(geo_code, geo_list),
    figure_id
  )


global_figure_metadata %>%
  select(
    geo_code,
    figure_id,
    title,
    subtitle,
    chart_display
  )

```





# I assembled each economy’s settings, cards, insights and figure information, together with one tailored Notes paragraph and one separate Source row. I preserved the approved note wording, updated the specified date references and Japan’s rounding explanation, and added checks for the assembled notes.

```{r profile_settings_notes_and_assembly_, message=FALSE}
# ============================================================
# CHUNK 18 — PROFILE SETTINGS, CONSOLIDATED NOTES AND ASSEMBLY
#
# Responsibility:
# Prepare every economy's complete profile object.
# Each profile contains exactly:
# - One Source row.
# - One consolidated Notes row.
#
# Approved note wording and sentence order are preserved.
# Availability years and Brazil's periods update from data.
# Japan's rounding explanation uses its unrounded Design CAGR.
#
# Chunk 19 will export these objects and verify the saved file.
# No workbook is created or saved in this chunk.
# ============================================================


# ------------------------------------------------------------
# 1. Check required inputs
# ------------------------------------------------------------

required_assembly_objects <- c(
  "economy_lookup",
  "geo_list",
  "investment_levels",
  "investment_cagr",
  "kpi_cards",
  "key_insights",
  "country_figure_values",
  "country_figure_metadata",
  "global_figure_metadata",
  "global_figure_values_by_profile",
  "ppp_input_unit",
  "report_year",
  "report_title",
  "source_text"
)

missing_assembly_objects <- required_assembly_objects[
  !vapply(
    required_assembly_objects,
    exists,
    logical(1),
    inherits = TRUE
  )
]

if (length(missing_assembly_objects) > 0L) {
  stop(
    "Run the preceding chunks first. Missing objects: ",
    paste(missing_assembly_objects, collapse = ", "),
    call. = FALSE
  )
}

if (
  !is.character(source_text) ||
  length(source_text) != 1L ||
  is.na(source_text) ||
  !nzchar(trimws(source_text))
) {
  stop(
    "source_text must contain one nonempty source statement.",
    call. = FALSE
  )
}


# ------------------------------------------------------------
# 2. Index and profile headers
#
# All economies:
# The Index's latest data year refers to the headline
# intangible investment level.
#
# Other indicators retain their own observation years.
# Worksheet names retain the approved full economy names.
# ------------------------------------------------------------

profile_index <- economy_lookup %>%
  select(
    display_order,
    geo_code,
    economy_name
  ) %>%
  left_join(
    investment_levels %>%
      select(
        geo_code,
        latest_data_year = latest_year_intangible
      ),
    by = "geo_code"
  ) %>%
  mutate(sheet_name = economy_name) %>%
  arrange(display_order)

note_geos <- profile_index$geo_code

profile_headers <- economy_lookup %>%
  transmute(
    geo_code,
    economy_name,

    header_title = paste0(
      economy_name,
      " (", geo_code, ")",
      " \u2014 Intangible Economy Profile ",
      report_year
    ),

    header_subtitle = paste0(
      report_title,
      " \u00b7 Intangible Economy Profile"
    )
  )


# ------------------------------------------------------------
# 3. Profile settings
#
# US / UK:
# Reuse figure_label, preserving "(US)" and "(UK)".
#
# All economies:
# Preserve the existing layout, colours and numeric conventions.
# ------------------------------------------------------------

shared_profile_settings <- tibble(
  field = c(
    "Titles",
    "Layout",
    "Colours",
    "Formatting",
    "Missing observations",
    "Source PPP level unit"
  ),

  value = c(
    paste0(
      report_title,
      " \u00b7 Intangible Economy Profile",
      " \u00b7 Key indicators",
      " \u00b7 Key insights",
      " \u00b7 Cross-country comparison"
    ),

    paste0(
      "Page 1: 4 KPI cards, Key insights, Figs 1–3, Figs 4–6. ",
      "Page 2: Figs 7–10 (2×2), source and notes. ",
      "Luiss and WIPO logos on both pages."
    ),

    paste0(
      "Cards #fff7ed; border #fed7aa; ",
      "Intangible #1e3a5f; Tangible #ea580c; ",
      "Gridlines #fde68a; Other economies #cbc9c4; ",
      "Titles #1c1c1c; Subtitles #78716c."
    ),

    paste0(
      "Half-up rounding. Levels, growth, intensity and chart ",
      "shares: one decimal place. Composition insights: whole ",
      "percentages. Omit trailing .0 in text. Rank using ",
      "unrounded values. Use USD and the word percent in insights."
    ),

    paste0(
      "Missing chart observations are represented by null ",
      "in the export. Preserve gaps in time series; ",
      "missing observations do not represent zero."
    ),

    gsub("_", " ", ppp_input_unit, fixed = TRUE)
  )
)

profile_settings <- bind_rows(
  lapply(note_geos, function(g) {

    economy <- economy_lookup %>%
      filter(geo_code == g)

    bind_rows(
      tibble(
        field = c(
          "Country name",
          "Code",
          "Figure label (Figs 1–6)"
        ),
        value = c(
          economy$economy_name,
          g,
          economy$figure_label
        )
      ),
      shared_profile_settings
    ) %>%
      mutate(geo_code = g, .before = 1L)
  })
)


# ------------------------------------------------------------
# 4. Approved note components
#
# Preserve the wording already approved for the consolidated
# notes, including the formal-sector GDP wording.
#
# Availability, Brazil's period and Japan's rounding sentence
# are populated below using the calculated results.
#
# Netherlands (NL):
# Keep "Kingdom of the Netherlands" in the code glossary,
# while retaining "Netherlands" as the worksheet/profile name.
# ------------------------------------------------------------

wiih_approved_note_parts <- list(

  availability = "",

  ppp = paste0(
    "All investment levels are expressed in purchasing power ",
    "parity-adjusted current prices."
  ),

  real_growth = paste0(
    "Compound annual growth rates (CAGRs) in Figures 3, 4 and 10 ",
    "are computed using investment data in chain-linked volume ",
    "terms to reflect real growth."
  ),

  brazil_period = "",

  intensity = paste0(
    "Investment intensities in Figures 1 and 9 are measured as ",
    "a share of GDP, where GDP is proxied by “adjusted” ",
    "value-added i.e. taking into account the full set of intangibles."
  ),

  comparability_singular = paste0(
    "In Figure 9, the share for India is based on formal sector ",
    "GDP only; comparable estimates for Brazil and the Philippines ",
    "are not yet available and are under development for future editions."
  ),

  codes = paste0(
    "Country codes in Figures 7-10 refer to the following economies: ",
    "BG: Bulgaria; BR: Brazil; CA: Canada; CZ: Czech Republic; ",
    "DE: Germany; DK: Denmark; EE: Estonia; EL: Greece; ES: Spain; ",
    "FI: Finland; FR: France; HR: Croatia; HU: Hungary; IND: India; ",
    "IT: Italy; JP: Japan; LT: Lithuania; LU: Luxembourg; LV: Latvia; ",
    "NL: Kingdom of the Netherlands; PHL: Philippines; PL: Poland; ",
    "PT: Portugal; RO: Romania; SE: Sweden; SI: Slovenia; SK: Slovakia; ",
    "UK: United Kingdom; US: United States of America."
  ),

  comparability_plural = paste0(
    "In Figure 9, shares for India are based on formal sector GDP ",
    "only; comparable estimates for Brazil and the Philippines ",
    "are not yet available and are under development for future editions."
  ),

  intensity_general = paste0(
    "In general, investment intensities in Figures 1 and 9 are ",
    "measured as a share of GDP, where GDP is proxied by “adjusted” ",
    "value-added i.e. taking into account the full set of intangibles."
  ),

  comparability_india = paste0(
    "In Figures 1 and 9, the share for India is based on formal ",
    "sector GDP only; comparable estimates for Brazil and the ",
    "Philippines are not yet available and are under development ",
    "for future editions."
  ),

  india_exclusion = paste0(
    "Figure 4 excludes the \"other IP products\" category due to ",
    "gaps in the time-series from 2013 to 2016."
  ),

  real_growth_japan = paste0(
    "Compound annual growth rates (CAGRs) in Figures 3, 4 and 10 ",
    "are computed using investment data in chain-linked volume ",
    "terms to reflect real growth and are rounded to one decimal point."
  ),

  japan_design = "",

  oipp_exclusion = paste0(
    "Figure 4 excludes the \"other IP products\" category due to ",
    "gaps in the time-series."
  )
)


# ------------------------------------------------------------
# 5. Assign the approved sentence order to every economy
#
# Brazil: comparison-period explanation.
# India: its approved intensity and OIPP-exclusion wording.
# Japan: its rounding wording and conditional Design explanation.
# Netherlands / Poland: their shared OIPP-exclusion wording.
# Canada / Philippines and remaining economies: standard wording.
#
# Every economy has an explicit assignment.
# ------------------------------------------------------------

wiih_note_component_order <- list(

  standard = c(
    "availability", "ppp", "real_growth",
    "intensity", "comparability_plural", "codes"
  ),

  brazil = c(
    "availability", "ppp", "real_growth", "brazil_period",
    "intensity", "comparability_singular", "codes"
  ),

  india = c(
    "availability", "ppp", "real_growth", "intensity_general",
    "comparability_india", "india_exclusion", "codes"
  ),

  japan = c(
    "availability", "ppp", "real_growth_japan", "japan_design",
    "intensity", "comparability_plural", "codes"
  ),

  netherlands_poland = c(
    "availability", "ppp", "real_growth", "intensity_general",
    "comparability_singular", "oipp_exclusion", "codes"
  )
)

wiih_note_group <- c(
  BR  = "brazil",
  BG  = "standard",
  CA  = "standard",
  HR  = "standard",
  CZ  = "standard",
  DK  = "standard",
  EE  = "standard",
  FI  = "standard",
  FR  = "standard",
  DE  = "standard",
  EL  = "standard",
  HU  = "standard",
  IND = "india",
  IT  = "standard",
  JP  = "japan",
  LV  = "standard",
  LT  = "standard",
  LU  = "standard",
  NL  = "netherlands_poland",
  PHL = "standard",
  PL  = "netherlands_poland",
  PT  = "standard",
  RO  = "standard",
  SK  = "standard",
  SI  = "standard",
  ES  = "standard",
  SE  = "standard",
  UK  = "standard",
  US  = "standard"
)

if (
  anyDuplicated(note_geos) ||
  !setequal(note_geos, names(wiih_note_group))
) {
  stop(
    "The note mapping must cover exactly the economies in the Index.",
    call. = FALSE
  )
}


# ------------------------------------------------------------
# 6. Update the availability sentence
#
# All economies:
# Preserve the approved sentence structure.
# Update its years and economy exceptions from profile_index.
# ------------------------------------------------------------

wiih_join_note_items <- function(x) {
  if (length(x) == 0L) return("")
  if (length(x) == 1L) return(x)

  paste0(
    paste(head(x, -1L), collapse = ", "),
    " and ",
    tail(x, 1L)
  )
}

note_years <- suppressWarnings(
  as.numeric(as.character(profile_index$latest_data_year))
)

if (any(
  !is.finite(note_years) |
    note_years != trunc(note_years)
)) {
  stop(
    "Every economy needs a valid latest data year for its note.",
    call. = FALSE
  )
}

latest_note_year <- max(note_years)
older <- which(note_years != latest_note_year)

if (length(older) == 0L) {

  wiih_approved_note_parts$availability <- paste0(
    "Notes: All estimates are based on the most recent data available—",
    latest_note_year,
    " for all economies."
  )

} else {

  older <- older[order(profile_index$economy_name[older])]

  exception_names <- profile_index$economy_name[older]
  exception_codes <- profile_index$geo_code[older]

  articles <- c(
    CZ  = "the Czech Republic",
    NL  = "the Netherlands",
    PHL = "the Philippines",
    UK  = "the United Kingdom",
    US  = "the United States of America"
  )

  use_article <- exception_codes %in% names(articles)

  exception_names[use_article] <- unname(
    articles[exception_codes[use_article]]
  )

  wiih_approved_note_parts$availability <- paste0(
    "Notes: All estimates are based on the most recent data available—",
    latest_note_year,
    " for all economies, with the exception of ",
    wiih_join_note_items(
      paste0(
        exception_names,
        " (", note_years[older], ")"
      )
    ),
    "."
  )
}


# ------------------------------------------------------------
# 7. Brazil: populate its approved comparison-period sentence
#
# Read Figure 3's actual comparison window directly from
# country_figure_values, which already exists from Chunk 15.
#
# Read the separate latest real-series years from Chunk 5.
# This avoids depending on all_geo_profiles before assembly.
# ------------------------------------------------------------

br_window <- country_figure_values$figure3 %>%
  filter(geo_code == "BR") %>%
  distinct(start_year, end_year)

br_latest <- investment_cagr %>%
  filter(geo_code == "BR")

if (
  nrow(br_window) != 1L ||
  nrow(br_latest) != 1L ||
  any(!is.finite(c(
    br_window$start_year,
    br_window$end_year,
    br_latest$end_year_tangible,
    br_latest$end_year_intangible
  )))
) {
  stop(
    "Brazil's note requires an unambiguous comparison period ",
    "and latest real-series years.",
    call. = FALSE
  )
}

wiih_approved_note_parts$brazil_period <- paste0(
  "For Brazil, growth rates in Figure 3 are reported for ",
  br_window$start_year, "–", br_window$end_year,
  " to ensure comparability between intangible and tangible investment, ",
  "as tangible investment data are available only through ",
  br_latest$end_year_tangible,
  ", while intangible investment data extend through ",
  br_latest$end_year_intangible,
  "."
)


# ------------------------------------------------------------
# 8. Japan: populate its Design rounding explanation
#
# Preserve the approved approach:
# - The chart may show "-0.0%".
# - The note reports the raw percentage rounded to two decimals.
#
# Example:
# -0.017075929249688% becomes -0.02% in the explanation.
#
# Remove this special sentence if the result no longer meets
# the small-negative condition in a future edition.
# ------------------------------------------------------------

wiih_note_two_decimals <- function(x) {

  stopifnot(length(x) == 1L, is.finite(x))

  decimal_text <- format(
    abs(x),
    digits = 15,
    scientific = FALSE,
    trim = TRUE,
    big.mark = "",
    decimal.mark = "."
  )

  pieces <- strsplit(
    decimal_text,
    ".",
    fixed = TRUE
  )[[1L]]

  fraction <- if (length(pieces) > 1L) {
    pieces[2L]
  } else {
    ""
  }

  fraction <- paste0(fraction, "000")

  cents <- as.numeric(
    paste0(pieces[1L], substr(fraction, 1L, 2L))
  )

  if (as.integer(substr(fraction, 3L, 3L)) >= 5L) {
    cents <- cents + 1
  }

  paste0(
    if (x < 0) "-" else "",
    formatC(
      cents / 100,
      format = "f",
      digits = 2,
      decimal.mark = "."
    )
  )
}

jp_design <- country_figure_values$figure4 %>%
  filter(
    geo_code == "JP",
    asset_code == "Design"
  )

if (
  nrow(jp_design) != 1L ||
  !"cagr_pct_raw" %in% names(jp_design) ||
  !is.finite(jp_design$cagr_pct_raw)
) {
  stop(
    "Japan's rounding note requires its unrounded Design CAGR.",
    call. = FALSE
  )
}

jp_raw <- jp_design$cagr_pct_raw

if (jp_raw < 0 && abs(jp_raw) < 0.05) {

  wiih_approved_note_parts$japan_design <- paste0(
    "In Figure 4, the CAGR for design is displayed as -0.0% due to ",
    "rounding; the underlying value, rounded to two decimal places, is ",
    wiih_note_two_decimals(jp_raw),
    "%."
  )

} else {

  wiih_approved_note_parts$japan_design <- ""

}


# ------------------------------------------------------------
# 9. Build one consolidated paragraph per economy
#
# Preserve the approved sentence order and wording.
# Join sentences with spaces into one text value.
# Include "Notes:" once, at the start of the paragraph.
# ------------------------------------------------------------

profile_notes <- bind_rows(
  lapply(note_geos, function(g) {

    note_group <- unname(wiih_note_group[g])
    ids <- wiih_note_component_order[[note_group]]

    if (
      length(ids) == 0L ||
      any(!ids %in% names(wiih_approved_note_parts))
    ) {
      stop(
        "Missing note components for ", g, ".",
        call. = FALSE
      )
    }

    sentences <- unlist(
      wiih_approved_note_parts[ids],
      use.names = FALSE
    )

    if (anyNA(sentences)) {
      stop(
        "Missing note sentence for ", g, ".",
        call. = FALSE
      )
    }

    sentences <- sentences[nzchar(trimws(sentences))]
    note_text <- paste(sentences, collapse = " ")

    if (
      !startsWith(note_text, "Notes: ") ||
      grepl("[\r\n]", note_text) ||
      nchar(note_text, type = "chars") > 32767L
    ) {
      stop(
        "Invalid consolidated note for ", g, ".",
        call. = FALSE
      )
    }

    tibble(
      geo_code = g,
      note_group = note_group,
      notes_text = note_text
    )
  })
)


# ------------------------------------------------------------
# 10. Build exactly two source/notes rows per economy
#
# Source: use source_text from Chunk 1.
# Notes: use the single approved paragraph created above.
#
# Rebuild this table on every run so notes cannot accumulate.
# ------------------------------------------------------------

profile_source_notes <- bind_rows(
  lapply(note_geos, function(g) {

    note_text <- profile_notes$notes_text[
      match(g, profile_notes$geo_code)
    ]

    tibble(
      geo_code = g,
      note_order = 0:1,
      element = c("Source", "Notes"),
      text = c(source_text, note_text)
    )
  })
)


# ------------------------------------------------------------
# 11. Assemble the complete profile objects
#
# All economies:
# Collect the existing calculations and presentation fields.
# Attach the already finalized Source and Notes rows.
#
# Global figure tables retain all comparison economies and
# their profile-specific highlighting.
# ------------------------------------------------------------

all_geo_profiles <- setNames(
  lapply(note_geos, function(g) {

    index_row <- profile_index %>%
      filter(geo_code == g)

    header_row <- profile_headers %>%
      filter(geo_code == g)

    list(
      geo_code = g,
      economy_name = index_row$economy_name,
      sheet_name = index_row$sheet_name,
      header_title = header_row$header_title,
      header_subtitle = header_row$header_subtitle,

      settings = profile_settings %>%
        filter(geo_code == g) %>%
        select(field, value),

      kpis = kpi_cards %>%
        filter(geo_code == g) %>%
        arrange(kpi_id),

      insights = key_insights %>%
        filter(geo_code == g) %>%
        arrange(insight_id),

      country_figure_metadata = country_figure_metadata %>%
        filter(geo_code == g) %>%
        arrange(figure_id),

      country_figure_values = lapply(
        country_figure_values,
        function(df) filter(df, geo_code == g)
      ),

      global_figure_metadata = global_figure_metadata %>%
        filter(geo_code == g) %>%
        arrange(figure_id),

      global_figure_values =
        global_figure_values_by_profile[[g]],

      source_and_notes = profile_source_notes %>%
        filter(geo_code == g) %>%
        arrange(note_order) %>%
        select(element, text)
    )
  }),
  note_geos
)


# ------------------------------------------------------------
# 12. Verify the assembled Source and Notes rows
#
# These checks remain in R and create no additional worksheet.
# Saved-file verification remains the responsibility of Chunk 19.
# ------------------------------------------------------------

for (g in note_geos) {

  source_notes <- all_geo_profiles[[g]]$source_and_notes

  if (
    nrow(source_notes) != 2L ||
    !identical(source_notes$element, c("Source", "Notes")) ||
    anyNA(source_notes$text) ||
    any(!nzchar(trimws(source_notes$text)))
  ) {
    stop(
      "Expected one Source row and one complete Notes row for ",
      g, ".",
      call. = FALSE
    )
  }
}

profile_index

```



# Finally I built the final export step to write a formatted master workbook with an Index and 29 economy worksheets containing all profile content. I added checks that reopen the saved workbook to verify worksheet order and consolidated note completeness, without creating a Review flags sheet.
```{r chunk19_export_master, message=FALSE, eval=FALSE}

# ============================================================
# CHUNK 19 — EXPORT AND VERIFY THE MASTER EXCEL WORKBOOK
#
# Inputs:
#   all_geo_profiles — assembled in Chunk 18
#   profile_index    — assembled in Chunk 18
#   report_title     — defined in Chunk 1
#   output_file      — defined in Chunk 1
#
# Responsibilities:
# - Export one Index sheet and one sheet per economy.
# - Preserve the calculations, labels, insights and notes
#   already assembled upstream.
# - Write exactly one Source row and one consolidated
#   Notes row per economy.
# - Verify the saved worksheet order and complete note text.
#
# Notes are constructed exclusively in Chunk 18.
# This chunk reads and exports them without rebuilding them.
#
# Economy-specific presentation:
# - India: preserve the approved intensity values.
# - UK: preserve "(UK)" in figure titles.
# - Japan: preserve small negative numeric growth values
#   and their separately supplied display labels.
# - All other exceptions remain those established upstream.
#
# No review-flags worksheet is created.
# ============================================================

library(openxlsx)
library(dplyr)


# ------------------------------------------------------------
# 1. Validate the assembled profiles
# ------------------------------------------------------------

required_export_objects <- c(
  "all_geo_profiles",
  "profile_index",
  "report_title",
  "output_file"
)

missing_export_objects <- required_export_objects[
  !vapply(
    required_export_objects,
    exists,
    logical(1),
    inherits = TRUE
  )
]

if (length(missing_export_objects) > 0L) {
  stop(
    "Run the earlier chunks first. Missing objects: ",
    paste(missing_export_objects, collapse = ", "),
    call. = FALSE
  )
}

export_index <- profile_index %>%
  arrange(display_order)

export_geos <- export_index$geo_code
export_sheet_names <- export_index$sheet_name

if (
  nrow(export_index) == 0L ||
  anyNA(export_geos) ||
  anyDuplicated(export_geos) ||
  anyDuplicated(names(all_geo_profiles)) ||
  length(all_geo_profiles) != length(export_geos) ||
  !setequal(names(all_geo_profiles), export_geos)
) {
  stop(
    "The Index and assembled economy profiles do not match.",
    call. = FALSE
  )
}


# Validate worksheet names without silently changing them.

invalid_sheet_name <- function(x) {

  if (is.na(x) || !nzchar(trimws(x))) {
    return(TRUE)
  }

  characters <- strsplit(x, "", fixed = TRUE)[[1L]]

  nchar(x) > 31 ||
    any(characters %in% c("\\", "/", ":", "*", "?", "[", "]")) ||
    grepl("[[:cntrl:]]", x) ||
    startsWith(x, "'") ||
    endsWith(x, "'")
}

if (
  any(vapply(
    export_sheet_names,
    invalid_sheet_name,
    logical(1)
  )) ||
  anyDuplicated(tolower(c("Index", export_sheet_names)))
) {
  stop(
    "Worksheet names are invalid or are not unique.",
    call. = FALSE
  )
}


# Check profile identifiers and required sections.

for (g in export_geos) {

  p <- all_geo_profiles[[g]]

  expected_sheet <- export_index$sheet_name[
    match(g, export_index$geo_code)
  ]

  if (
    !identical(p$geo_code, g) ||
    !identical(p$sheet_name, expected_sheet)
  ) {
    stop(
      "Profile identifiers do not match the Index for ",
      g, ".",
      call. = FALSE
    )
  }

  if (
    !identical(sort(as.integer(p$kpis$kpi_id)), 1:4) ||
    !identical(sort(as.integer(p$insights$insight_id)), 1:7) ||
    !identical(
      sort(as.integer(p$country_figure_metadata$figure_id)),
      1:6
    ) ||
    !identical(
      sort(as.integer(p$global_figure_metadata$figure_id)),
      7:10
    )
  ) {
    stop(
      "Incomplete or duplicated profile sections for ", g,
      ": expected 4 KPIs, 7 insights and Figures 1–10.",
      call. = FALSE
    )
  }

  if (
    !all(
      paste0("figure", 1:6) %in%
        names(p$country_figure_values)
    ) ||
    !all(
      paste0("figure", 7:10) %in%
        names(p$global_figure_values)
    )
  ) {
    stop(
      "Missing figure-value tables for ", g, ".",
      call. = FALSE
    )
  }
}


# ------------------------------------------------------------
# 1A. Validate the finalized Source and Notes rows
#
# All economies:
# Read the two rows prepared in Chunk 18.
# Do not reconstruct, replace or append note sentences here.
# ------------------------------------------------------------

for (g in export_geos) {

  source_notes <- all_geo_profiles[[g]]$source_and_notes

  if (
    !is.data.frame(source_notes) ||
    !all(c("element", "text") %in% names(source_notes))
  ) {
    stop(
      "Invalid Source/Notes structure for ", g,
      ". Rerun the replacement Chunk 18.",
      call. = FALSE
    )
  }

  if (
    nrow(source_notes) != 2L ||
    !identical(
      as.character(source_notes$element),
      c("Source", "Notes")
    ) ||
    anyNA(source_notes$text) ||
    any(!nzchar(trimws(source_notes$text)))
  ) {
    stop(
      "Expected exactly one Source row and one Notes row for ",
      g, ". Rerun the replacement Chunk 18.",
      call. = FALSE
    )
  }
}

# Record the exported cell positions for saved-file verification.
note_export_positions <- list()


# ------------------------------------------------------------
# 2. Create the workbook and shared styles
# ------------------------------------------------------------

wb_master <- createWorkbook()

excel_styles <- list(

  title = createStyle(
    fontName = "Calibri",
    fontSize = 16,
    fontColour = "#1e3a5f",
    textDecoration = "bold",
    wrapText = TRUE,
    valign = "top"
  ),

  subtitle = createStyle(
    fontName = "Calibri",
    fontSize = 10,
    fontColour = "#78716c",
    wrapText = TRUE,
    valign = "top"
  ),

  section = createStyle(
    fontName = "Calibri",
    fontSize = 11,
    fontColour = "#1e3a5f",
    textDecoration = "bold",
    fgFill = "#ffedd5",
    wrapText = TRUE,
    valign = "top"
  ),

  header = createStyle(
    fontName = "Calibri",
    fontSize = 10,
    fontColour = "#1e3a5f",
    textDecoration = "bold",
    fgFill = "#f3f4f6",
    border = "Bottom",
    borderColour = "#d6d3d1",
    wrapText = TRUE,
    valign = "top"
  ),

  body = createStyle(
    fontName = "Calibri",
    fontSize = 10,
    fontColour = "#1c1c1c",
    wrapText = TRUE,
    valign = "top"
  ),

  highlight = createStyle(
    fontColour = "#1e3a5f",
    textDecoration = "bold",
    fgFill = "#eaf0f7"
  ),

  link = createStyle(
    fontColour = "#1e3a5f",
    textDecoration = "underline"
  )
)


# Numeric values were prepared upstream.
#
# Percentages are already in percentage units:
# 10 means 10%, not 0.10.
# Do not apply Excel's multiplying percentage format.

excel_numeric_styles <- list(
  integer = createStyle(numFmt = "0"),
  decimal = createStyle(numFmt = "0.0"),
  tiny_negative = createStyle(numFmt = "General")
)


# Two figure-table panels, separated by column F.
# Notes occupy one wrapped cell spanning columns C:K.

profile_column_widths <- c(
  6, 26, 14, 14, 14, 3, 6, 26, 14, 14, 14
)

index_column_widths <- c(6, 36, 10, 20, 40)


# Keep the maximum required height when panels share a row.

excel_layout <- new.env(parent = emptyenv())
excel_layout$widths <- list()
excel_layout$heights <- list()


# ------------------------------------------------------------
# 3. Reusable writing and layout helpers
# ------------------------------------------------------------

wiih_register_sheet <- function(sheet, widths) {

  addWorksheet(
    wb_master,
    sheet,
    gridLines = FALSE
  )

  setColWidths(
    wb_master,
    sheet,
    cols = seq_along(widths),
    widths = widths
  )

  excel_layout$widths[[sheet]] <- widths
  excel_layout$heights[[sheet]] <- numeric(0)
}


wiih_record_heights <- function(sheet, rows, heights) {

  stored <- excel_layout$heights[[sheet]]

  for (i in seq_along(rows)) {

    key <- as.character(rows[i])

    previous <- if (key %in% names(stored)) {
      stored[[key]]
    } else {
      0
    }

    stored[key] <- max(previous, heights[i])
  }

  excel_layout$heights[[sheet]] <- stored
}


wiih_text_height <- function(
  sheet,
  text,
  columns,
  font_size = 10
) {

  if (length(text) == 0L || is.na(text)) {
    text <- ""
  }

  available_width <- sum(
    excel_layout$widths[[sheet]][columns]
  )

  wrap_width <- max(
    1L,
    floor(available_width * 0.82)
  )

  paragraphs <- strsplit(
    as.character(text),
    "\n",
    fixed = TRUE
  )[[1L]]

  if (length(paragraphs) == 0L) {
    paragraphs <- ""
  }

  line_count <- sum(vapply(
    paragraphs,
    function(paragraph) {
      max(
        1L,
        length(strwrap(paragraph, width = wrap_width))
      )
    },
    integer(1)
  ))

  height <- max(
    22,
    line_count * 15 * font_size / 10 + 8
  )

  if (height > 409) {
    stop(
      "A text row exceeds Excel's row-height limit on sheet '",
      sheet,
      "'. Widen the text span while retaining one Notes row.",
      call. = FALSE
    )
  }

  height
}


wiih_banner <- function(
  sheet,
  row,
  text,
  columns,
  style = excel_styles$section,
  font_size = 11
) {

  writeData(
    wb_master,
    sheet,
    x = text,
    startRow = row,
    startCol = min(columns),
    colNames = FALSE
  )

  if (length(columns) > 1L) {
    mergeCells(
      wb_master,
      sheet,
      cols = columns,
      rows = row
    )
  }

  addStyle(
    wb_master,
    sheet,
    style = style,
    rows = row,
    cols = columns,
    gridExpand = TRUE,
    stack = TRUE
  )

  wiih_record_heights(
    sheet,
    row,
    wiih_text_height(
      sheet,
      text,
      columns,
      font_size
    )
  )

  row + 1L
}


wiih_write_table <- function(
  sheet,
  row,
  df,
  spans,
  missing_text = "Not available"
) {

  df <- as.data.frame(
    df,
    stringsAsFactors = FALSE
  )

  if (
    ncol(df) != length(spans) ||
    any(vapply(df, is.list, logical(1)))
  ) {
    stop(
      "Invalid table structure on sheet '",
      sheet, "'.",
      call. = FALSE
    )
  }

  columns <- seq.int(
    min(unlist(spans)),
    max(unlist(spans))
  )


  # Write headers into their designated column spans.

  for (j in seq_along(spans)) {

    writeData(
      wb_master,
      sheet,
      x = names(df)[j],
      startRow = row,
      startCol = min(spans[[j]]),
      colNames = FALSE
    )

    if (length(spans[[j]]) > 1L) {
      mergeCells(
        wb_master,
        sheet,
        cols = spans[[j]],
        rows = row
      )
    }
  }

  addStyle(
    wb_master,
    sheet,
    style = excel_styles$header,
    rows = row,
    cols = columns,
    gridExpand = TRUE,
    stack = TRUE
  )

  header_height <- max(vapply(
    seq_along(spans),
    function(j) {
      wiih_text_height(
        sheet,
        names(df)[j],
        spans[[j]]
      )
    },
    numeric(1)
  ))

  wiih_record_heights(
    sheet,
    row,
    header_height
  )


  # Empty tables receive an explicit availability message.

  if (nrow(df) == 0L) {

    wiih_banner(
      sheet,
      row + 1L,
      "No data available",
      columns,
      style = excel_styles$body,
      font_size = 10
    )

    return(list(
      next_row = row + 3L,
      data_start = row + 1L,
      last_row = row + 1L,
      columns = columns
    ))
  }

  data_rows <- row + seq_len(nrow(df))

  addStyle(
    wb_master,
    sheet,
    style = excel_styles$body,
    rows = data_rows,
    cols = columns,
    gridExpand = TRUE,
    stack = TRUE
  )


  # Write each column without converting numeric values to text.

  for (j in seq_along(spans)) {

    values <- df[[j]]
    first_column <- min(spans[[j]])

    writeData(
      wb_master,
      sheet,
      x = df[j],
      startRow = row + 1L,
      startCol = first_column,
      colNames = FALSE,
      rowNames = FALSE,
      keepNA = TRUE,
      na.string = missing_text
    )

    if (length(spans[[j]]) > 1L) {
      for (r in data_rows) {
        mergeCells(
          wb_master,
          sheet,
          cols = spans[[j]],
          rows = r
        )
      }
    }

    if (is.numeric(values)) {

      valid <- is.finite(values)

      if (any(is.infinite(values))) {
        stop(
          "Infinite numeric value in '", names(df)[j],
          "' on sheet '", sheet, "'.",
          call. = FALSE
        )
      }

      formats <- rep(
        "decimal",
        length(values)
      )

      formats[
        valid & values == trunc(values)
      ] <- "integer"

      # Japan and any equivalent future observation:
      # preserve small negative growth numerically.

      formats[
        valid &
          values < 0 &
          abs(values) < 0.05
      ] <- "tiny_negative"

      for (format_name in unique(formats[valid])) {

        selected_rows <- data_rows[
          valid & formats == format_name
        ]

        addStyle(
          wb_master,
          sheet,
          style = excel_numeric_styles[[format_name]],
          rows = selected_rows,
          cols = first_column,
          gridExpand = TRUE,
          stack = TRUE
        )
      }
    }
  }


  # Estimate row heights from the widest required text block.

  body_heights <- vapply(
    seq_len(nrow(df)),
    function(i) {

      max(vapply(
        seq_along(spans),
        function(j) {

          value <- df[[j]][i]

          if (is.na(value)) {
            value <- missing_text
          }

          wiih_text_height(
            sheet,
            value,
            spans[[j]]
          )
        },
        numeric(1)
      ))
    },
    numeric(1)
  )

  wiih_record_heights(
    sheet,
    data_rows,
    body_heights
  )

  list(
    next_row = max(data_rows) + 2L,
    data_start = min(data_rows),
    last_row = max(data_rows),
    columns = columns
  )
}


wiih_write_block <- function(
  sheet,
  row,
  title,
  df,
  spans,
  missing_text = "Not available"
) {

  columns <- seq.int(
    min(unlist(spans)),
    max(unlist(spans))
  )

  table_row <- wiih_banner(
    sheet,
    row,
    title,
    columns
  )

  wiih_write_table(
    sheet,
    table_row,
    df,
    spans,
    missing_text = missing_text
  )
}


wiih_apply_text_colour <- function(
  sheet,
  row,
  columns,
  colour_text
) {

  if (is.na(colour_text)) {
    return(invisible(NULL))
  }

  match_position <- regexpr(
    "#[[:xdigit:]]{6}",
    colour_text
  )

  if (match_position[1L] < 0) {
    return(invisible(NULL))
  }

  hex_colour <- regmatches(
    colour_text,
    match_position
  )

  addStyle(
    wb_master,
    sheet,
    style = createStyle(fontColour = hex_colour),
    rows = row,
    cols = columns,
    gridExpand = TRUE,
    stack = TRUE
  )
}


wiih_highlight_profile <- function(
  sheet,
  block,
  figure_data
) {

  selected <- which(
    figure_data$highlighted %in% TRUE
  )

  if (length(selected) > 0L) {
    addStyle(
      wb_master,
      sheet,
      style = excel_styles$highlight,
      rows = block$data_start + selected - 1L,
      cols = block$columns,
      gridExpand = TRUE,
      stack = TRUE
    )
  }
}


wiih_figure_title <- function(metadata, figure_id) {

  paste0(
    "Figure ", figure_id, " — ",
    metadata$title[
      match(figure_id, metadata$figure_id)
    ]
  )
}


# ------------------------------------------------------------
# 4. Export the Index
#
# Latest data years remain numeric.
# Hyperlinks point to the matching economy worksheets.
# ------------------------------------------------------------

wiih_register_sheet(
  "Index",
  index_column_widths
)

wiih_banner(
  "Index",
  1,
  paste0(
    report_title,
    " — Intangible Economy Profiles"
  ),
  1:5,
  style = excel_styles$title,
  font_size = 16
)

wiih_banner(
  "Index",
  2,
  "Master table · one sheet per economy",
  1:5,
  style = excel_styles$subtitle,
  font_size = 10
)

index_table <- export_index %>%
  transmute(
    `#` = row_number(),
    Economy = economy_name,
    Code = geo_code,
    `Latest data year` = latest_data_year,
    Sheet = paste("Go to", economy_name)
  )

index_block <- wiih_write_table(
  "Index",
  4,
  index_table,
  spans = as.list(1:5)
)

for (i in seq_len(nrow(export_index))) {

  target_sheet <- gsub(
    "'",
    "''",
    export_index$sheet_name[i],
    fixed = TRUE
  )

  link_label <- gsub(
    '"',
    '""',
    index_table$Sheet[i],
    fixed = TRUE
  )

  link_formula <- paste0(
    "HYPERLINK(\"#'",
    target_sheet,
    "'!A1\",\"",
    link_label,
    "\")"
  )

  writeFormula(
    wb_master,
    "Index",
    x = link_formula,
    startRow = index_block$data_start + i - 1L,
    startCol = 5
  )
}

addStyle(
  wb_master,
  "Index",
  style = excel_styles$link,
  rows = index_block$data_start +
    seq_len(nrow(export_index)) - 1L,
  cols = 5,
  gridExpand = TRUE,
  stack = TRUE
)

freezePane(
  wb_master,
  "Index",
  firstActiveRow = 5
)


# ------------------------------------------------------------
# 5. Export the economy worksheets
#
# Each worksheet receives:
# - Profile settings.
# - Four KPI cards.
# - Seven key insights.
# - Country-figure metadata and values.
# - Global-figure metadata and values.
# - One Source row and one Notes row.
# ------------------------------------------------------------

for (g in export_geos) {

  p <- all_geo_profiles[[g]]
  sheet <- p$sheet_name

  wiih_register_sheet(
    sheet,
    profile_column_widths
  )

  wiih_banner(
    sheet,
    1,
    p$header_title,
    1:11,
    style = excel_styles$title,
    font_size = 16
  )

  wiih_banner(
    sheet,
    2,
    p$header_subtitle,
    1:11,
    style = excel_styles$subtitle,
    font_size = 10
  )

  row_pos <- 4L


  # ---- Profile settings -------------------------------------

  settings_table <- p$settings %>%
    transmute(
      Field = field,
      Value = value
    )

  block <- wiih_write_block(
    sheet,
    row_pos,
    "Profile settings",
    settings_table,
    spans = list(1:2, 3:11)
  )

  row_pos <- block$next_row


  # ---- Four KPI cards ---------------------------------------

  kpis <- p$kpis %>%
    arrange(kpi_id)

  kpi_table <- kpis %>%
    transmute(
      `#` = kpi_id,
      Label = label,
      Value = value,
      `Value colour` = value_colour,
      `Sub-text` = subtext,
      `Sub-text colour` = subtext_colour
    )

  block <- wiih_write_block(
    sheet,
    row_pos,
    "Key indicators — KPI cards",
    kpi_table,
    spans = list(1, 2, 3, 4:5, 6:9, 10:11)
  )

  for (i in seq_len(nrow(kpis))) {

    kpi_row <- block$data_start + i - 1L

    wiih_apply_text_colour(
      sheet,
      kpi_row,
      3,
      kpis$value_colour[i]
    )

    wiih_apply_text_colour(
      sheet,
      kpi_row,
      6:9,
      kpis$subtext_colour[i]
    )
  }

  row_pos <- block$next_row


  # ---- Seven key insights -----------------------------------

  insight_table <- p$insights %>%
    arrange(insight_id) %>%
    transmute(
      `#` = insight_id,
      Text = text
    )

  block <- wiih_write_block(
    sheet,
    row_pos,
    "Key insights",
    insight_table,
    spans = list(1, 2:11)
  )

  row_pos <- block$next_row


  # ---- Country-figure metadata ------------------------------

  country_metadata <- p$country_figure_metadata %>%
    arrange(figure_id)

  country_metadata_table <- country_metadata %>%
    transmute(
      Fig = figure_id,
      Title = title,
      Subtitle = subtitle,
      `Chart / display` = chart_display,
      Range = range
    )

  block <- wiih_write_block(
    sheet,
    row_pos,
    "Country figures — metadata",
    country_metadata_table,
    spans = list(1, 2:3, 4:6, 7:9, 10:11)
  )

  row_pos <- block$next_row


  # ---- Country-figure values --------------------------------

  f <- p$country_figure_values

  country_tables <- list(

    figure1 = f$figure1 %>%
      arrange(year) %>%
      select(
        Year = year,
        `Intangible (% of GDP)`,
        `Tangible (% of GDP)`
      ),

    figure2 = f$figure2 %>%
      arrange(year) %>%
      select(
        Year = year,
        `Intangible (index)`,
        `Tangible (index)`
      ),

    figure3 = f$figure3 %>%
      arrange(row_order) %>%
      transmute(
        `#` = row_order,
        Series = series,
        `CAGR (%)` = cagr_pct,
        `Bar colour` = bar_colour,
        `Label colour` = label_colour
      ),

    figure4 = f$figure4 %>%
      arrange(row_order) %>%
      transmute(
        `#` = row_order,
        `Asset type` = asset_name,
        `CAGR (%)` = cagr_pct,
        `Bar colour` = bar_colour,
        `Shown as` = value_label
      ),

    figure5 = f$figure5 %>%
      arrange(row_order) %>%
      transmute(
        `#` = row_order,
        `Asset type` = asset_name,
        `Share (%)` = share_pct,
        `Slice colour` = slice_colour,
        `Label shown` = label_shown
      ),

    figure6 = f$figure6 %>%
      arrange(row_order) %>%
      transmute(
        `#` = row_order,
        Slice = component,
        `Share (%)` = share_pct,
        `Fill colour` = fill_colour,
        `Label colour` = label_colour
      )
  )

  row_pos <- wiih_banner(
    sheet,
    row_pos,
    "Country figures — values",
    1:11
  )


  # Left panel: Figure 1.

  left_block <- wiih_write_block(
    sheet,
    row_pos,
    wiih_figure_title(country_metadata, 1),
    country_tables$figure1,
    spans = as.list(1:3),
    missing_text = "null"
  )


  # Right panel: Figures 2, 3 and 6, stacked vertically.

  right_block <- wiih_write_block(
    sheet,
    row_pos,
    wiih_figure_title(country_metadata, 2),
    country_tables$figure2,
    spans = as.list(7:9),
    missing_text = "null"
  )

  right_block <- wiih_write_block(
    sheet,
    right_block$next_row,
    wiih_figure_title(country_metadata, 3),
    country_tables$figure3,
    spans = as.list(7:11),
    missing_text = "null"
  )

  right_block <- wiih_write_block(
    sheet,
    right_block$next_row,
    wiih_figure_title(country_metadata, 6),
    country_tables$figure6,
    spans = as.list(7:11),
    missing_text = "null"
  )

  row_pos <- max(
    left_block$next_row,
    right_block$next_row
  )


  # Next pair: Figures 4 and 5.
  #
  # India, Netherlands and Poland:
  # Preserve the growth exclusions selected upstream.
  # Composition retains all seven asset categories.

  left_block <- wiih_write_block(
    sheet,
    row_pos,
    wiih_figure_title(country_metadata, 4),
    country_tables$figure4,
    spans = as.list(1:5),
    missing_text = "null"
  )

  right_block <- wiih_write_block(
    sheet,
    row_pos,
    wiih_figure_title(country_metadata, 5),
    country_tables$figure5,
    spans = as.list(7:11),
    missing_text = "null"
  )

  row_pos <- max(
    left_block$next_row,
    right_block$next_row
  )


  # ---- Cross-country metadata -------------------------------

  global_metadata <- p$global_figure_metadata %>%
    arrange(figure_id)

  global_metadata_table <- global_metadata %>%
    transmute(
      Fig = figure_id,
      Title = title,
      Subtitle = subtitle,
      `Chart / highlight` = chart_display
    )

  block <- wiih_write_block(
    sheet,
    row_pos,
    "Cross-country comparison — metadata",
    global_metadata_table,
    spans = list(1, 2:3, 4:8, 9:11)
  )

  row_pos <- block$next_row


  # ---- Cross-country values ---------------------------------
  #
  # Preserve every comparison economy in these tables.
  # The highlighted field identifies the current profile.

  gf <- lapply(
    p$global_figure_values,
    function(x) arrange(x, row_order)
  )

  level_export_table <- function(x) {
    x %>%
      transmute(
        `#` = row_order,
        `Economy code` = geo_code,
        `USD bn` = usd_billions,
        Year = year,
        `Shown as` = value_label
      )
  }

  global_tables <- list(

    figure7 = level_export_table(gf$figure7),

    figure8 = level_export_table(gf$figure8),

    figure9 = gf$figure9 %>%
      transmute(
        `#` = row_order,
        Economy = geo_code,
        `Intensity (%)` = intensity_pct,
        Year = year
      ),

    figure10 = gf$figure10 %>%
      transmute(
        `#` = row_order,
        Economy = geo_code,
        `CAGR (%)` = cagr_pct,
        `Start year` = start_year,
        `End year` = end_year
      )
  )

  row_pos <- wiih_banner(
    sheet,
    row_pos,
    "Cross-country comparison — values",
    1:11
  )


  # First pair: Figures 7 and 8.

  left_block <- wiih_write_block(
    sheet,
    row_pos,
    wiih_figure_title(global_metadata, 7),
    global_tables$figure7,
    spans = as.list(1:5),
    missing_text = "null"
  )

  wiih_highlight_profile(
    sheet,
    left_block,
    gf$figure7
  )

  right_block <- wiih_write_block(
    sheet,
    row_pos,
    wiih_figure_title(global_metadata, 8),
    global_tables$figure8,
    spans = as.list(7:11),
    missing_text = "null"
  )

  wiih_highlight_profile(
    sheet,
    right_block,
    gf$figure8
  )

  row_pos <- max(
    left_block$next_row,
    right_block$next_row
  )


  # Second pair: Figures 9 and 10.

  left_block <- wiih_write_block(
    sheet,
    row_pos,
    wiih_figure_title(global_metadata, 9),
    global_tables$figure9,
    spans = as.list(1:4),
    missing_text = "null"
  )

  wiih_highlight_profile(
    sheet,
    left_block,
    gf$figure9
  )

  right_block <- wiih_write_block(
    sheet,
    row_pos,
    wiih_figure_title(global_metadata, 10),
    global_tables$figure10,
    spans = as.list(7:11),
    missing_text = "null"
  )

  wiih_highlight_profile(
    sheet,
    right_block,
    gf$figure10
  )

  row_pos <- max(
    left_block$next_row,
    right_block$next_row
  )


  # ---- Source and one consolidated Notes row -----------------
  #
  # Use the exact text supplied by Chunk 18.
  # The complete note occupies one wrapped cell across C:K.

  source_notes_table <- p$source_and_notes %>%
    transmute(
      Element = element,
      Text = text
    )

  notes_block <- wiih_write_block(
    sheet,
    row_pos,
    "Source and notes",
    source_notes_table,
    spans = list(1:2, 3:11)
  )

  note_export_positions[[g]] <- list(
    sheet = sheet,
    source_row = notes_block$data_start,
    notes_row = notes_block$data_start + 1L,
    source_text = source_notes_table$Text[1L],
    notes_text = source_notes_table$Text[2L]
  )

  freezePane(
    wb_master,
    sheet,
    firstActiveRow = 4
  )
}


# ------------------------------------------------------------
# 6. Apply the final row heights
# ------------------------------------------------------------

for (sheet in names(excel_layout$heights)) {

  heights <- excel_layout$heights[[sheet]]

  setRowHeights(
    wb_master,
    sheet,
    rows = as.integer(names(heights)),
    heights = unname(heights)
  )
}


# ------------------------------------------------------------
# 7. Save the workbook once
#
# Use output_file from Chunk 1.
# Its filename is:
# WIIH_2026_Intangible_Economy_Profiles_MASTER_FINAL.xlsx
# ------------------------------------------------------------

export_directory <- dirname(output_file)

if (!dir.exists(export_directory)) {
  dir.create(
    export_directory,
    recursive = TRUE,
    showWarnings = FALSE
  )
}

if (!dir.exists(export_directory)) {
  stop(
    "Cannot create the output directory: ",
    export_directory,
    call. = FALSE
  )
}

tryCatch(
  saveWorkbook(
    wb_master,
    file = output_file,
    overwrite = TRUE
  ),

  error = function(e) {
    stop(
      "The master workbook could not be saved to:\n",
      output_file,
      "\nCheck that the destination is writable and that the ",
      "workbook is not open in Excel.\n",
      "Original error: ",
      conditionMessage(e),
      call. = FALSE
    )
  }
)


# ------------------------------------------------------------
# 8. Verify the saved worksheet names and order
# ------------------------------------------------------------

expected_sheets <- c(
  "Index",
  export_sheet_names
)

saved_sheets <- getSheetNames(output_file)

if (!identical(saved_sheets, expected_sheets)) {
  stop(
    "The saved workbook's worksheets do not match the Index.",
    call. = FALSE
  )
}


# ------------------------------------------------------------
# 9. Verify the saved Source and Notes cells
#
# Read the actual saved text and compare it with Chunk 18's
# intended text, recorded during export.
#
# Confirm that every economy sheet contains one Notes row.
# These checks remain in R and create no extra worksheet.
# ------------------------------------------------------------

notes_export_check <- bind_rows(
  lapply(export_geos, function(g) {

    position <- note_export_positions[[g]]

    saved_rows <- read.xlsx(
      output_file,
      sheet = position$sheet,
      rows = c(
        position$source_row,
        position$notes_row
      ),
      cols = c(1, 3),
      colNames = FALSE,
      skipEmptyRows = FALSE,
      skipEmptyCols = TRUE
    )

    if (
      nrow(saved_rows) != 2L ||
      ncol(saved_rows) != 2L ||
      !identical(
        as.character(saved_rows[[1L]]),
        c("Source", "Notes")
      ) ||
      !identical(
        as.character(saved_rows[[2L]]),
        c(
          position$source_text,
          position$notes_text
        )
      )
    ) {
      stop(
        "The saved Source/Notes cells do not match ",
        "the intended text for ", g, ".",
        call. = FALSE
      )
    }

    labels <- read.xlsx(
      output_file,
      sheet = position$sheet,
      cols = 1,
      colNames = FALSE,
      skipEmptyRows = FALSE,
      skipEmptyCols = FALSE
    )[[1L]]

    note_row_count <- sum(
      grepl(
        "^notes?($|[[:space:]]+[0-9])",
        trimws(as.character(labels)),
        ignore.case = TRUE
      ),
      na.rm = TRUE
    )

    if (note_row_count != 1L) {
      stop(
        "Expected one saved Notes row, found ",
        note_row_count,
        " for ", g, ".",
        call. = FALSE
      )
    }

    data.frame(
      geo_code = g,
      sheet = position$sheet,
      notes_rows = note_row_count,
      full_text_matches = TRUE,
      stringsAsFactors = FALSE
    )
  })
)

stopifnot(
  nrow(notes_export_check) == length(export_geos),
  all(notes_export_check$notes_rows == 1L),
  all(notes_export_check$full_text_matches)
)

message(
  "Master workbook saved: ",
  normalizePath(
    output_file,
    winslash = "/",
    mustWork = TRUE
  ),
  "\nSheets: 1 Index + ",
  length(export_geos),
  " economy profiles.",
  "\nOne complete tailored Notes row verified on every economy sheet.",
  "\nNo review-flags sheet included."
)

```








