# Load necessary libraries for data I/O and manipulation
library(readxl)   # to read Excel input files
library(openxlsx) # to write output Excel files

# Set random seed for reproducibility (if desired)
# (Original script sets RNG seed in master script, if applicable)
# set.seed(123)

# Scenario names and timeline (months of interest)
Scenario_names <- c('Escalation', 'Status Quo', 'Ceasefire')
months <- c('23 Oct','23 Nov','23 Dec','24 Jan','24 Feb','24 Mar','24 Apr','24 May','24 Jun','24 Jul')
Nt <- length(months)  # total number of months in simulation output timeline (Oct 2023 to Jul 2024)
periods <- c('Months 1-3 (Feb-Apr)', 'Months 4-6 (May-Jul)')  # labels for output periods

# -------------------- IMPORT NCD BASELINE DEATH DATA ---------------------------
ts_past <- 2015:2022  # years of available death data
# Read the baseline total death data for NCDs (sheet "Total Baseline Death")
baseline_file <- "inputs/NCD_M2_model_para_1.26.2024.xlsx"
df_bd <- read_excel(baseline_file, sheet = "Total Baseline Death")
# The sheet is assumed to have columns named by NCD (and possibly Year)
# If a Year column exists, drop it for analysis (since ts_past provides the year vector)
if ("Year" %in% names(df_bd) || "year" %in% names(df_bd)) {
  # Remove or ignore the Year column (since we have ts_past explicitly)
  # (This step ensures df_bd columns correspond to NCD names directly)
  df_bd <- df_bd[ , !(names(df_bd) %in% c("Year", "year")) ]
}

# Prepare baseline data as a list of (NCD, death counts vector) for quick access (similar to Ds_past in MATLAB)
Ds_past <- list()
for (col in names(df_bd)) {
  Ds_past[[col]] <- as.numeric(df_bd[[col]])
}

# ----------------- IMPORT COVERAGE RATE OVER TIME -----------------------------
cs_names <- c('CKD','CVD','Cancer','DM')  # coverage categories corresponding to groupings of NCDs
# Read the coverage rate data (sheet "coverage rate")
coverage_data <- read_excel(baseline_file, sheet = "coverage rate")
# Initialize lists to hold lower-bound and upper-bound coverage matrices for each category
cs_all_LBs <- vector("list", length(cs_names))
cs_all_UBs <- vector("list", length(cs_names))
# The sheet is expected to have columns named like "Worst_CKD_lb", "Central_CKD_lb", "Best_CKD_lb", etc., for each category
for (i in seq_along(cs_names)) {
  cat <- cs_names[i]
  # Construct column names for worst, central, best (lower bound)
  col_worst_lb   <- paste0("Worst_", cat, "_lb")
  col_central_lb <- paste0("Central_", cat, "_lb")
  col_best_lb    <- paste0("Best_", cat, "_lb")
  # Extract the columns as vectors (assuming they exist in coverage_data)
  r1 <- as.numeric(coverage_data[[col_worst_lb]])
  r2 <- as.numeric(coverage_data[[col_central_lb]])
  r3 <- as.numeric(coverage_data[[col_best_lb]])
  # Combine into a matrix: each column is one scenario (Worst=1, Central=2, Best=3)
  cs_all_LBs[[i]] <- cbind(r1, r2, r3)
  
  # Similarly for upper bound columns
  col_worst_ub   <- paste0("Worst_", cat, "_ub")
  col_central_ub <- paste0("Central_", cat, "_ub")
  col_best_ub    <- paste0("Best_", cat, "_ub")
  u1 <- as.numeric(coverage_data[[col_worst_ub]])
  u2 <- as.numeric(coverage_data[[col_central_ub]])
  u3 <- as.numeric(coverage_data[[col_best_ub]])
  cs_all_UBs[[i]] <- cbind(u1, u2, u3)
}

# ----------------- LINK EACH NCD WITH ITS CATEGORY ------------------------
# Define the mapping of each specific NCD to a coverage category (mirrors NCD_categories in MATLAB)
NCD_categories <- data.frame(
  NCD = c('DM1','DM2','CKD','IHD','CDIS','HD','HS','IS','COPD','Bcancer','Ccancer','Lcancer'),
  Category = c('DM','DM','CKD','CVD','CVD','CVD','CVD','CVD','COPD','Cancer','Cancer','Cancer'),
  stringsAsFactors = FALSE
)

# -------------------- PREPARE FOR THE MAIN SIMULATION ---------------------
# Choose the NCD for simulation by index (user can change k_NCD_categories to select a different NCD)
k_NCD_categories <- 8  # default index (e.g., 8 corresponds to 'IS' in the list above)
# Determine the specific NCD name and its category
NCD <- NCD_categories$NCD[k_NCD_categories]                # e.g., "IS"
NCD_category <- NCD_categories$Category[k_NCD_categories]  # e.g., "CVD"
# Read hazard rate time series for this NCD from the pre-generated Excel file
hazard_file <- "inputs/HR_new_logistic_normal.xlsx"
HR_ALL_t <- read_excel(hazard_file, sheet = NCD, col_names = TRUE)
# Extract hazard rate vectors (each of length 500, representing tau=0 to 499 months)
HR_tlb  <- HR_ALL_t$HR_t_tlb   # treated, lower-bound parameters
HR_tub  <- HR_ALL_t$HR_t_tub   # treated, upper-bound parameters
HR_utlb <- HR_ALL_t$HR_t_utlb  # untreated, lower-bound parameters
HR_utub <- HR_ALL_t$HR_t_utub  # untreated, upper-bound parameters

# Determine if this NCD uses the "single cohort" strategy (DM1 case):contentReference[oaicite:12]{index=12}
# DM1 starts with an existing prevalent cohort, others start from zero prevalence.
flag_single_cohort <- ifelse(NCD == "DM1", TRUE, FALSE)

# Extract coverage trajectories for the appropriate category of this NCD
# Find index of this category in cs_names
k_cs <- match(NCD_category, cs_names)  # e.g., "CVD" corresponds to index 2
# Retrieve the pre-loaded coverage matrices for this category
cov_LB_mat <- cs_all_LBs[[k_cs]]  # matrix of lower bound coverage (rows=timepoints, cols=Worst/Central/Best)
cov_UB_mat <- cs_all_UBs[[k_cs]]  # matrix of upper bound coverage
# Number of time points provided in coverage data for this category
nc <- nrow(cov_LB_mat)

# ---------------- LINEAR REGRESSION TO GET 2023 BASELINE DEATH ----------------
# Get historical death counts for this NCD (2015–2022) from Ds_past list
D_past_full <- Ds_past[[NCD]]
# Remove any missing (NaN) years if present
valid_idx <- which(!is.na(D_past_full))
D_past <- D_past_full[valid_idx]         # deaths in years with data
t_past <- ts_past[valid_idx]            # corresponding years
# Perform a simple linear regression: D ~ 1 + Year (to project to 2023):contentReference[oaicite:13]{index=13}
reg_model <- lm(D_past ~ t_past)
coeff <- coef(reg_model)                # (Intercept) and slope
D_2023 <- coeff[1] + coeff[2] * 2023    # estimated baseline (counterfactual) deaths in 2023

# Optionally, plot the baseline trend and 2023 estimate (disabled by default, set flag_plot_D_past = TRUE to enable)
flag_plot_D_past <- FALSE
if (flag_plot_D_past) {
  plot(t_past, D_past, type='b', pch=16, col='blue', xlab='Year', ylab='Annual deaths',
       main=paste(NCD, "- baseline deaths trend"))
  abline(reg_model, col='darkgreen')
  points(2023, D_2023, col='red', pch=18, cex=1.2)
}

# -------------------- DEFINE SIMULATION FUNCTION ----------------------------
# f_D implements the cohort simulation for one scenario (returns mean deaths and all simulation draws):contentReference[oaicite:14]{index=14}:contentReference[oaicite:15]{index=15}
f_D <- function(nt, tC, Nsim, nsim, cs_all_LB_vec, cs_all_UB_vec, flag_single_cohort, scenario_index,
                HR_tlb_vec, HR_tub_vec, HR_utlb_vec, HR_utub_vec) {
  # nt: total simulation time steps (months)
  # tC: index of conflict start (after baseline "burn-in")
  # Nsim: base incidence per month (number of individuals per cohort)
  # nsim: number of simulation runs
  # cs_all_LB_vec, cs_all_UB_vec: vectors of coverage fraction (lower and upper bounds) of length = coverage data duration
  # flag_single_cohort: whether to use single initial cohort (TRUE for DM1)
  # scenario_index: 1=Worst (Escalation), 2=Central (Status Quo), 3=Best (Ceasefire)
  # HR_*_vec: hazard rate vectors for treated/untreated, lower/upper bound (length 500)
  
  # Prepare coverage fraction over the whole timeline (length nt)
  cs_LB <- rep(1, nt)  # default 1 (100% coverage) before conflict
  cs_UB <- rep(1, nt)
  # Apply scenario-specific coverage from conflict start for the duration of provided data
  cs_LB[tC:(tC + nc - 1)] <- cs_all_LB_vec[, scenario_index]
  cs_UB[tC:(tC + nc - 1)] <- cs_all_UB_vec[, scenario_index]
  # After coverage data ends, hold the last value constant for remaining time
  if ((tC + nc) <= nt) {
    cs_LB[(tC + nc):nt] <- cs_LB[tC + nc - 1]
    cs_UB[(tC + nc):nt] <- cs_UB[tC + nc - 1]
  }
  
  # Set coverage to 1 (full) for all pre-conflict months
  if (tC > 1) {
    cs_LB[1:(tC - 1)] <- 1
    cs_UB[1:(tC - 1)] <- 1
  }
  
  # Initialize arrays to store results for each simulation run
  Dsim    <- matrix(0, nrow = nt, ncol = nsim)  # deaths per time in crisis scenario
  D_BLsim <- matrix(0, nrow = nt, ncol = nsim)  # deaths per time in baseline scenario
  D_excsim <- matrix(0, nrow = nt, ncol = nsim) # excess deaths per time (difference)
  # (Psim and P_BLsim will store prevalence if needed, but are not returned explicitly)
  
  # Monte Carlo simulation runs
  for (k in 1:nsim) {
    # Initialize deaths vectors for this run
    D <- numeric(nt)       # deaths at each time (scenario)
    D_BL <- numeric(nt)    # deaths at each time (baseline)
    D_exc <- numeric(nt)   # excess deaths at each time
    
    # Initialize cohort incidence matrices P and P_BL (nt x nt) for this run:contentReference[oaicite:16]{index=16}
    if (flag_single_cohort) {
      # Single cohort: start all Nsim individuals at t=1, no new incidences afterwards
      P <- matrix(0, nt, nt)
      P[1, 1] <- Nsim
      P_BL <- P
    } else {
      # Continuous incidence: each month t0 (1..nt) starts a cohort of Nsim individuals
      P <- diag(Nsim, nrow = nt, ncol = nt)    # each diagonal entry P(t0,t0) = Nsim
      P_BL <- diag(Nsim, nrow = nt, ncol = nt) # baseline scenario starts same number
    }
    
    # Draw a random coverage trajectory between the LB and UB bounds (one random offset per simulation run):contentReference[oaicite:17]{index=17}
    c_rand <- runif(1)  # random number in [0,1) for this run
    cs <- cs_LB + c_rand * (cs_UB - cs_LB)  # interpolate coverage between lower and upper trajectories
    
    # Loop over cohort start times t0 (each cohort introduced at time t0)
    for (t0 in 1:nt) {
      # Loop over each time t from t0 (cohort introduction) to end of simulation
      for (t in t0:nt) {
        # Current coverage fraction at time t
        cs_t <- cs[t]
        # Determine hazard rates at (t - t0) months since this cohort's start:contentReference[oaicite:18]{index=18}
        # (Use hazard vectors, adjusting index: +1 because R index starts at 1 for t0 itself corresponding to tau=0)
        idx <- (t - t0) + 1  # index into hazard vectors (1 corresponds to acute phase at cohort start)
        HR_with_lb  <- HR_tlb_vec[idx]   # treated lower-bound hazard at this interval
        HR_with_ub  <- HR_tub_vec[idx]   # treated upper-bound hazard
        HR_without_lb <- HR_utlb_vec[idx] # untreated lower-bound hazard
        HR_without_ub <- HR_utub_vec[idx] # untreated upper-bound hazard
        # Randomly pick actual hazard for treated and untreated within the bounds:contentReference[oaicite:19]{index=19}
        HR_with    <- HR_with_lb    + runif(1) * (HR_with_lb    - HR_with_ub)
        HR_without <- HR_without_lb + runif(1) * (HR_without_lb - HR_without_ub)
        # Baseline hazard (with full treatment) at this time
        h0 <- HR_with
        # Actual hazard considering coverage cs_t (mixture of with vs without treatment):contentReference[oaicite:20]{index=20}
        h  <- cs_t * HR_with + (1 - cs_t) * HR_without
        # Ensure hazard in crisis scenario is not below baseline hazard (take max):contentReference[oaicite:21]{index=21}
        h <- max(h0, h)
        
        # Determine how many individuals die in this cohort at time t for scenario vs baseline:contentReference[oaicite:22]{index=22}
        nP    <- P[t0, t]    # number of people alive in this cohort at time t (scenario)
        nP_BL <- P_BL[t0, t] # number alive in baseline cohort at time t
        nh <- max(nP, nP_BL) # total individuals to consider for random draw (max of both)
        if (nh > 0) {
          # Generate random uniforms for each individual (coupled for scenario vs baseline):contentReference[oaicite:23]{index=23}
          r <- runif(nh)
          # Deaths in scenario cohort (if any alive)
          d <- if (nP > 0) sum(r[1:nP] < h) else 0
          # Deaths in baseline cohort
          d_BL <- if (nP_BL > 0) sum(r[1:nP_BL] < h0) else 0
          # Record deaths at time t
          D[t]    <- D[t] + d
          D_BL[t] <- D_BL[t] + d_BL
          D_exc[t] <- D_exc[t] + (d - d_BL)  # excess deaths in this cohort at time t
          # Update survivors for next time step (if not at end):contentReference[oaicite:24]{index=24}
          if (t < nt) {
            P[t0, t + 1]    <- P[t0, t]    - d
            P_BL[t0, t + 1] <- P_BL[t0, t] - d_BL
          }
        } else {
          # If no one is alive in either cohort at this time, break out early for this cohort
          next  # (skip to next t if no survivors; in R, 'next' in inner loop continues to next t)
        }
      }  # end loop over t for cohort t0
    }  # end loop over cohort start t0
    
    # Store results of this simulation run in the result matrices
    Dsim[, k]    <- D
    D_BLsim[, k] <- D_BL
    D_excsim[, k] <- D_exc
    # (Optionally, could store prevalence Psim and P_BLsim if needed:
    # Psim <- colSums(P)
    # P_BLsim <- colSums(P_BL)
    # Not returned by this function as they're not used in later analysis.)
  }  # end simulation runs
  
  # Compute mean deaths over simulations for scenario and baseline (per time point)
  D_mean    <- rowMeans(Dsim)
  D_mean_BL <- rowMeans(D_BLsim)
  # Return a list of outputs analogous to [D_mean, D_mean_BL, D_C, D_BL, D_exc]
  return(list(D_mean = D_mean,
              D_mean_BL = D_mean_BL,
              D_C = Dsim,
              D_BL = D_BLsim,
              D_exc = D_excsim))
}

# -------------------- RUN SIMULATION FOR THREE SCENARIOS ----------------------
# Determine base incidence (Nsim) for this NCD (monthly incidence), as given in original Nsims mapping:contentReference[oaicite:25]{index=25}:contentReference[oaicite:26]{index=26}
Nsims_values <- c(CKD=25, IHD=473, HS=58, IS=26, Bcancer=30, Ccancer=6, Lcancer=25, DM1=2947)
Nsim <- round(Nsims_values[NCD])  # base number of individuals per monthly cohort

nsim <- 100  # number of simulation iterations (could be increased for smoother estimates)
# Set conflict timing: burn-in + conflict + projection as per original:contentReference[oaicite:27]{index=27}
if (flag_single_cohort) {
  tC <- 1            # conflict starts immediately for DM1 case (we model existing prevalence as at conflict start)
} else {
  tC <- round(12 * 30)  # ~360 months burn-in before conflict to reach steady state baseline
}
tnow <- 5          # months from conflict start to "current" time (approx war duration up to end Jan 2024)
nt <- tC + tnow + 6 + 1  # total timeline length (conflict + 6 months projection + small buffer)

# Prepare storage for scenario results
D_exc_means <- list()    # mean excess deaths per time for each scenario
D_C_total <- list()      # list of matrices of total deaths per scenario (time x nsim)
D_C_excess_sorted <- list()  # list of matrices of sorted excess deaths per scenario (time x nsim)

# Loop over scenarios 1=Escalation, 2=Status Quo, 3=Ceasefire:contentReference[oaicite:28]{index=28}
for (k_Scenario in 1:3) {
  scenario_name <- Scenario_names[k_Scenario]
  # Run the simulation function for this scenario
  sim_res <- f_D(nt, tC, Nsim, nsim, cs_all_LBs[[k_cs]], cs_all_UBs[[k_cs]], flag_single_cohort,
                 k_Scenario, HR_tlb, HR_tub, HR_utlb, HR_utub)
  D_mean0   <- sim_res$D_mean      # mean deaths (scenario)
  D_mean_BL0<- sim_res$D_mean_BL   # mean deaths (baseline)
  D_C0     <- sim_res$D_C         # matrix of scenario deaths (time x nsim)
  D_BL0    <- sim_res$D_BL        # matrix of baseline deaths (time x nsim)
  D_exc0   <- sim_res$D_exc       # matrix of excess deaths (time x nsim)
  
  # Scale results so that baseline simulation matches the real 2023 baseline death estimate:contentReference[oaicite:29]{index=29}
  if (flag_single_cohort) {
    D_mult <- 1.0
  } else {
    # Compute scaling factor: ratio of estimated 2023 deaths to simulated baseline (annualized)
    D_mult <- D_2023 / (mean(D_mean_BL0) * 12)
    message(sprintf("D_mult = %.2f (multiplying simulation outputs by this factor)", D_mult))
  }
  # Apply scaling factor to all outputs (scenario and baseline)
  D_mean   <- D_mean0   * D_mult
  D_mean_BL<- D_mean_BL0* D_mult
  D_C      <- D_C0      * D_mult
  D_C_BL   <- D_BL0     * D_mult
  D_C_exc  <- D_exc0    * D_mult  # scaled excess deaths matrix
  
  # Sort each time point's simulations to derive uncertainty intervals:contentReference[oaicite:30]{index=30}
  D_C_sorted    <- t(apply(D_C, 1, sort))
  D_C_BL_sorted <- t(apply(D_C_BL, 1, sort))
  # Calculate 95% uncertainty interval (2.5th and 97.5th percentiles) for scenario and baseline deaths
  idx_low <- round(nsim * 0.025)
  idx_high<- round(nsim * 0.975)
  if (idx_low < 1) idx_low <- 1
  if (idx_high < 1) idx_high <- 1
  if (idx_low > nsim) idx_low <- nsim
  if (idx_high > nsim) idx_high <- nsim
  D_C_CI    <- pmax(0, cbind(D_C_sorted[, idx_low], D_C_sorted[, idx_high]))
  D_C_BL_CI <- pmax(0, cbind(D_C_BL_sorted[, idx_low], D_C_BL_sorted[, idx_high]))
  
  # Store mean excess deaths for this scenario (vector of length nt)
  D_exc_means[[k_Scenario]] <- (D_mean - D_mean_BL)
  # Store the full scaled results matrix for scenario total deaths and sorted excess deaths
  D_C_total[[k_Scenario]] <- D_C
  D_C_excess_sorted[[k_Scenario]] <- t(apply(D_C_exc, 1, sort))
}

# Scenario loop finished. Now save outputs to files (in outputs/Raw directory):contentReference[oaicite:31]{index=31}:contentReference[oaicite:32]{index=32}:
# Ensure output directory exists
if (!dir.exists("outputs/Raw")) dir.create("outputs/Raw", recursive = TRUE)

# 1. Save all runs results for Baseline deaths (one CSV per NCD)
baseline_filename <- sprintf("outputs/Raw/baseline_%s.csv", NCD)
# Write the baseline deaths matrix (time x nsim) with no row or column names (raw format)
write.table(D_C_BL, file = baseline_filename, sep = ",", row.names = FALSE, col.names = FALSE)

# 2. Save all runs results for Total deaths by scenario (one Excel file with 3 sheets per NCD)
scenario_total_file <- sprintf("outputs/Raw/scenario_total_%s.xlsx", NCD)
wb_tot <- createWorkbook()
for (i in seq_along(Scenario_names)) {
  sheet <- Scenario_names[i]
  addWorksheet(wb_tot, sheetName = sheet)
  writeData(wb_tot, sheet = sheet, x = D_C_total[[i]], colNames = FALSE, rowNames = FALSE)
}
saveWorkbook(wb_tot, file = scenario_total_file, overwrite = TRUE)

# 3. Save all runs results for Excess deaths by scenario (one Excel file with 3 sheets per NCD)
scenario_excess_file <- sprintf("outputs/Raw/scenario_excess_%s.xlsx", NCD)
wb_exc <- createWorkbook()
for (i in seq_along(Scenario_names)) {
  sheet <- Scenario_names[i]
  addWorksheet(wb_exc, sheetName = sheet)
  writeData(wb_exc, sheet = sheet, x = D_C_excess_sorted[[i]], colNames = FALSE, rowNames = FALSE)
}
saveWorkbook(wb_exc, file = scenario_excess_file, overwrite = TRUE)
