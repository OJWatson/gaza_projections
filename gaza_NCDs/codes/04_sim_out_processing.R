# Load libraries for reading/writing Excel and data manipulation
library(readxl)
library(openxlsx)

# Read age distribution for NCD deaths (sheet "Death" contains age groups and distribution by NCD):contentReference[oaicite:36]{index=36}
age_dist_data <- read_excel("inputs/NCD_M2_model_para_1.26.2024.xlsx", sheet = "Death")
# Ensure the data frame has an 'Age' column and columns for each NCD's age distribution (e.g., "CKD_Age_Distribution")
# (We assume the sheet has columns named like "<NCD>_Age_Distribution" and a column "Age".)
age_groups <- age_dist_data$Age  # vector of age categories
# Prepare a list or environment for quick access to each NCD's age distribution vector
age_distribution <- list()
for (col in names(age_dist_data)) {
  if (grepl("_Age_Distribution$", col)) {
    # Extract NCD name from the column name (remove suffix "_Age_Distribution")
    ncd_name <- sub("_Age_Distribution$", "", col)
    age_distribution[[ncd_name]] <- as.numeric(age_dist_data[[col]])
  }
}

# Define helper functions for distributing cumulative deaths by age

cum_1346_by_age_sce <- function(ncd_name, res_total) {
  # res_total: a list or environment with scenario names as keys and numeric matrix (time x simulations) as values
  dist <- age_distribution[[ncd_name]]  # age distribution vector for this NCD
  df_out <- data.frame()
  # Iterate scenarios: Escalation, Status Quo, Ceasefire
  for (s in c('Escalation', 'Status Quo', 'Ceasefire')) {
    # Sum total deaths in months 1-3 (Feb, Mar, Apr) -> indices 5,6,7 in 1-indexed R (4:6 in 0-index Python)
    total_13 <- sum(res_total[[s]][5:7, ])
    # Distribute this total across ages
    values <- dist * total_13   # this produces a vector of length = number of age groups
    # Compute mean, 2.5th percentile, and 97.5th percentile for each age group
    # (Since values for each age are constant (no distribution across runs after summing), mean = lb = ub = that value)
    means <- values
    lower_bounds <- values
    upper_bounds <- values
    # Add columns to output data frame for this scenario and period
    df_out[[paste0(ncd_name, s, "_Month1-3_mean")]] <- round(means, 2)
    df_out[[paste0(ncd_name, s, "_Month1-3_lb")]]   <- round(lower_bounds, 2)
    df_out[[paste0(ncd_name, s, "_Month1-3_ub")]]   <- round(upper_bounds, 2)
    
    # Sum total deaths in months 4-6 (May, Jun, Jul) -> indices 8,9,10 in R (7:9 in 0-index)
    total_46 <- sum(res_total[[s]][8:10, ])
    values <- dist * total_46
    means <- values; lower_bounds <- values; upper_bounds <- values
    df_out[[paste0(ncd_name, s, "_Month4-6_mean")]] <- round(means, 2)
    df_out[[paste0(ncd_name, s, "_Month4-6_lb")]]   <- round(lower_bounds, 2)
    df_out[[paste0(ncd_name, s, "_Month4-6_ub")]]   <- round(upper_bounds, 2)
    
    # Sum total deaths in months 1-6 combined (Feb–Jul)
    total_all <- sum(res_total[[s]][5:10, ])
    values <- dist * total_all
    means <- values; lower_bounds <- values; upper_bounds <- values
    df_out[[paste0(ncd_name, s, "_All_mean")]] <- round(means, 2)
    df_out[[paste0(ncd_name, s, "_All_lb")]]   <- round(lower_bounds, 2)
    df_out[[paste0(ncd_name, s, "_All_ub")]]   <- round(upper_bounds, 2)
  }
  # Insert the Age column at the front
  df_out <- cbind(Age = age_groups, df_out)
  # Write the result to an Excel file in the NCD-specific outputs folder
  out_dir <- file.path("outputs", ncd_name)
  if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
  out_file <- file.path(out_dir, paste0("scenario_total_13-46M_by_age_.xlsx"))
  write.xlsx(df_out, file = out_file, sheetName = ncd_name, overwrite = TRUE)
}

cum_1346_by_age_base <- function(ncd_name, baseline_mat) {
  # baseline_mat: matrix of baseline scenario deaths (time x simulations)
  dist <- age_distribution[[ncd_name]]
  df_out <- data.frame()
  # Sum baseline deaths in months 1-3 (Feb–Apr) and distribute
  total_13 <- sum(baseline_mat[5:7, ])
  values <- dist * total_13
  means <- values; lower_bounds <- values; upper_bounds <- values
  df_out[[paste0(ncd_name, "_Month1-3_mean")]] <- round(means, 2)
  df_out[[paste0(ncd_name, "_Month1-3_lb")]]   <- round(lower_bounds, 2)
  df_out[[paste0(ncd_name, "_Month1-3_ub")]]   <- round(upper_bounds, 2)
  # Months 4-6
  total_46 <- sum(baseline_mat[8:10, ])
  values <- dist * total_46
  means <- values; lower_bounds <- values; upper_bounds <- values
  df_out[[paste0(ncd_name, "_Month4-6_mean")]] <- round(means, 2)
  df_out[[paste0(ncd_name, "_Month4-6_lb")]]   <- round(lower_bounds, 2)
  df_out[[paste0(ncd_name, "_Month4-6_ub")]]   <- round(upper_bounds, 2)
  # All 6 months
  total_all <- sum(baseline_mat[5:10, ])
  values <- dist * total_all
  means <- values; lower_bounds <- values; upper_bounds <- values
  df_out[[paste0(ncd_name, "_All_mean")]] <- round(means, 2)
  df_out[[paste0(ncd_name, "_All_lb")]]   <- round(lower_bounds, 2)
  df_out[[paste0(ncd_name, "_All_ub")]]   <- round(upper_bounds, 2)
  # Add Age column
  df_out <- cbind(Age = age_groups, df_out)
  # Write to Excel file
  out_dir <- file.path("outputs", ncd_name)
  if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
  out_file <- file.path(out_dir, paste0("baseline_total_13-46M_by_age_.xlsx"))
  write.xlsx(df_out, file = out_file, sheetName = ncd_name, overwrite = TRUE)
}

cum_1346_by_age_excess <- function(ncd_name, excess) {
  # excess: list with scenario keys and excess deaths matrix (time x simulations)
  dist <- age_distribution[[ncd_name]]
  df_out <- data.frame()
  for (s in c('Escalation', 'Status Quo', 'Ceasefire')) {
    total_13 <- sum(excess[[s]][5:7, ])
    values <- dist * total_13
    means <- values; lower_bounds <- values; upper_bounds <- values
    df_out[[paste0(ncd_name, s, "_Month1-3_mean")]] <- round(means, 2)
    df_out[[paste0(ncd_name, s, "_Month1-3_lb")]]   <- round(lower_bounds, 2)
    df_out[[paste0(ncd_name, s, "_Month1-3_ub")]]   <- round(upper_bounds, 2)
    total_46 <- sum(excess[[s]][8:10, ])
    values <- dist * total_46
    means <- values; lower_bounds <- values; upper_bounds <- values
    df_out[[paste0(ncd_name, s, "_Month4-6_mean")]] <- round(means, 2)
    df_out[[paste0(ncd_name, s, "_Month4-6_lb")]]   <- round(lower_bounds, 2)
    df_out[[paste0(ncd_name, s, "_Month4-6_ub")]]   <- round(upper_bounds, 2)
    total_all <- sum(excess[[s]][5:10, ])
    values <- dist * total_all
    means <- values; lower_bounds <- values; upper_bounds <- values
    df_out[[paste0(ncd_name, s, "_All_mean")]] <- round(means, 2)
    df_out[[paste0(ncd_name, s, "_All_lb")]]   <- round(lower_bounds, 2)
    df_out[[paste0(ncd_name, s, "_All_ub")]]   <- round(upper_bounds, 2)
  }
  df_out <- cbind(Age = age_groups, df_out)
  out_dir <- file.path("outputs", ncd_name)
  if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
  out_file <- file.path(out_dir, paste0("scenario_excess_13-46M_by_age_.xlsx"))
  write.xlsx(df_out, file = out_file, sheetName = ncd_name, overwrite = TRUE)
}

cum_1346_by_age_all_ncds <- function(ncd_name, total_ncd) {
  # total_ncd: list with scenario keys and combined all-NCDs excess matrix (time x simulations)
  dist <- age_distribution[[ncd_name]]
  df_out <- data.frame()
  for (s in c('Escalation', 'Status Quo', 'Ceasefire')) {
    total_13 <- sum(total_ncd[[s]][5:7, ])
    values <- dist * total_13
    means <- values; lower_bounds <- values; upper_bounds <- values
    df_out[[paste0(ncd_name, s, "_Month1-3_mean")]] <- round(means, 2)
    df_out[[paste0(ncd_name, s, "_Month1-3_lb")]]   <- round(lower_bounds, 2)
    df_out[[paste0(ncd_name, s, "_Month1-3_ub")]]   <- round(upper_bounds, 2)
    total_46 <- sum(total_ncd[[s]][8:10, ])
    values <- dist * total_46
    means <- values; lower_bounds <- values; upper_bounds <- values
    df_out[[paste0(ncd_name, s, "_Month4-6_mean")]] <- round(means, 2)
    df_out[[paste0(ncd_name, s, "_Month4-6_lb")]]   <- round(lower_bounds, 2)
    df_out[[paste0(ncd_name, s, "_Month4-6_ub")]]   <- round(upper_bounds, 2)
    total_all <- sum(total_ncd[[s]][5:10, ])
    values <- dist * total_all
    means <- values; lower_bounds <- values; upper_bounds <- values
    df_out[[paste0(ncd_name, s, "_All_mean")]] <- round(means, 2)
    df_out[[paste0(ncd_name, s, "_All_lb")]]   <- round(lower_bounds, 2)
    df_out[[paste0(ncd_name, s, "_All_ub")]]   <- round(upper_bounds, 2)
  }
  df_out <- cbind(Age = age_groups, df_out)
  # Write combined all-NCD results to a single Excel file
  out_file <- "outputs/Total_ncd_1346M.xlsx"
  write.xlsx(df_out, file = out_file, sheetName = ncd_name, overwrite = TRUE)
}

# -------------------- READ SIMULATION OUTPUTS AND PROCESS --------------------

# List of NCD names of interest (consistent with earlier scripts)
NCD_list <- c('CKD','IHD','HS','IS','Bcancer','Ccancer','Lcancer','DM1')

# 1. Process baseline outputs for each NCD
for (ncd in NCD_list) {
  # Read baseline CSV (no header, time x sims matrix)
  baseline_file <- sprintf("outputs/Raw/baseline_%s.csv", ncd)
  baseline_mat <- as.matrix(read.csv(baseline_file, header = FALSE))
  # Distribute baseline deaths by age and output to Excel
  cum_1346_by_age_base(ncd, baseline_mat)
}

# 2. Process scenario total outputs for each NCD
for (ncd in NCD_list) {
  scenario_file <- sprintf("outputs/Raw/scenario_total_%s.xlsx", ncd)
  xls <- loadWorkbook(scenario_file)
  # Read each scenario sheet into a matrix
  res_total <- list()
  for (s in Scenario_names) {
    # readWorkbook returns a data frame; convert to matrix and treat blank header if any
    df <- read.xlsx(xls, sheet = s, colNames = FALSE)
    mat <- as.matrix(df)
    res_total[[s]] <- mat
  }
  cum_1346_by_age_sce(ncd, res_total)
}

# 3. Process scenario excess outputs for each NCD
for (ncd in NCD_list) {
  excess_file <- sprintf("outputs/Raw/scenario_excess_%s.xlsx", ncd)
  xls <- loadWorkbook(excess_file)
  excess_list <- list()
  for (s in Scenario_names) {
    df <- read.xlsx(xls, sheet = s, colNames = FALSE)
    mat <- as.matrix(df)
    excess_list[[s]] <- mat
  }
  cum_1346_by_age_excess(ncd, excess_list)
}

# 4. Combine all NCDs' excess deaths into a total and distribute by age
# First, read all scenario excess files and sum across NCDs
total_ncd <- list('Escalation' = NULL, 'Status Quo' = NULL, 'Ceasefire' = NULL)
# Initialize total matrices for each scenario (assuming all have same dimensions, e.g., 10 x nsim)
for (s in Scenario_names) {
  # Use the first NCD to get dimensions
  sample_file <- sprintf("outputs/Raw/scenario_excess_%s.xlsx", NCD_list[1])
  sample_df <- read.xlsx(sample_file, sheet = s, colNames = FALSE)
  total_ncd[[s]] <- matrix(0, nrow = nrow(sample_df), ncol = ncol(sample_df))
}
# Sum up all NCD excess deaths for each scenario
for (ncd in NCD_list) {
  excess_file <- sprintf("outputs/Raw/scenario_excess_%s.xlsx", ncd)
  xls <- loadWorkbook(excess_file)
  for (s in Scenario_names) {
    df <- read.xlsx(xls, sheet = s, colNames = FALSE)
    mat <- as.matrix(df)
    total_ncd[[s]] <- total_ncd[[s]] + mat
  }
}
# Use one of the NCDs (e.g., DM1) for distribution (original code uses ncd_name parameter, likely DM1):contentReference[oaicite:37]{index=37}
cum_1346_by_age_all_ncds("DM1", total_ncd)
