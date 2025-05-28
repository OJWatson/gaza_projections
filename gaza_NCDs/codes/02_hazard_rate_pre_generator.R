# Load required library for writing Excel files
# (Install openxlsx via install.packages("openxlsx") if not already installed)
library(openxlsx)

# Define NCDs and their corresponding hazard rate function names (as in original script)
NCDs <- c('CKD', 'IHD', 'HS', 'IS', 'Bcancer', 'Ccancer', 'Lcancer', 'DM1')

# Map each NCD to the hazard function to use (matching original dictionary keys)
ncd_survival_function <- list(
  CKD = "log_normal_hazard_rate",
  IHD = "log_normal_hazard_rate",
  HS  = "log_logistic_hazard_rate",
  IS  = "log_logistic_hazard_rate",
  Bcancer = "log_logistic_hazard_rate",
  Ccancer = "log_logistic_hazard_rate",
  Lcancer = "log_logistic_hazard_rate",
  DM1 = "negative_expoential_hazard_rate"  # note: "expoential" is intentionally misspelled to mirror original
)

# Parameter values for survival curves (treated/untreated, lower/upper bounds) for each NCD
survival_curve_para <- list(
  CKD = list(
    treated_lb   = c(100, 3.99524753, 1.68136759),
    treated_ub   = c(100, 3.99524753, 1.68136759),
    untreated_lb = c(100, 3.14329388, 1.58136759),
    untreated_ub = c(100, 3.25922182, 1.58136759)
  ),
  IHD = list(
    treated_lb   = c(94.3, 8.10297550, 5.19497573),
    treated_ub   = c(97.2, 8.10297547, 5.19497571),
    untreated_lb = c(60.0, 2.90401215, 4.50269156),
    untreated_ub = c(60.0, 2.90401215, 4.50269156)
  ),
  HS = list(
    treated_lb   = c(68.0, 0.70166823, 0.12837254),
    treated_ub   = c(68.0, 20.71026698, 0.32401575),
    untreated_lb = c(51.0, 0.24012431, 0.17809074),
    untreated_ub = c(51.0, 0.24012431, 0.17809074)
  ),
  IS = list(
    treated_lb   = c(93.0, 63.17745322, 0.58047708),
    treated_ub   = c(93.0, 63.17745322, 0.58047708),
    untreated_lb = c(88.0, 18.55677113, 0.58644473),
    untreated_ub = c(88.0, 18.55677113, 0.58644473)
  ),
  Bcancer = list(
    treated_lb   = c(100, 221.35196983, 1.06114687),
    treated_ub   = c(100, 221.35196983, 1.06114687),
    untreated_lb = c(100, 146.60253830, 1.06114687),
    untreated_ub = c(100, 146.60253830, 1.06114687)
  ),
  Ccancer = list(
    treated_lb   = c(100, 84.40664198, 0.72789679),
    treated_ub   = c(100, 84.40664198, 0.72789679),
    untreated_lb = c(100, 64.49123835, 0.65451315),
    untreated_ub = c(100, 64.49123835, 0.65451315)
  ),
  Lcancer = list(
    treated_lb   = c(100, 23.64434258, 0.66017550),
    treated_ub   = c(100, 23.64434258, 0.66017550),
    untreated_lb = c(100, 7.15460378, 0.66017550),
    untreated_ub = c(100, 7.15460378, 0.66017550)
  ),
  DM1 = list(
    treated_lb   = c(100, 100, 1/75.6/12),
    treated_ub   = c(100, 100, 1/75.6/12),
    untreated_lb = c(100, 100, 1/1.3/12),
    untreated_ub = c(100, 100, 1/2.6/12)
  )
)

# Define hazard rate functions corresponding to each distribution
log_logistic_hazard_rate <- function(t, alpha, beta) {
  # Hazard for Log-Logistic distribution at time t with scale=alpha, shape=beta
  pdf_ll  <- (beta/alpha) * ( (t/alpha)^(beta-1) ) / (1 + (t/alpha)^beta)^2
  cdf_ll  <- 1 / (1 + (alpha/t)^beta)
  return(pdf_ll / (1 - cdf_ll))
}
log_normal_hazard_rate <- function(t, mu, sigma) {
  # Hazard for Log-Normal distribution at time t with parameters mu, sigma (log-space mean and std)
  pdf_ln <- (1 / (t * sigma * sqrt(2*pi))) * exp(- (log(t) - mu)^2 / (2 * sigma^2))
  cdf_ln <- pnorm(log(t), mean = mu, sd = sigma)
  return(pdf_ln / (1 - cdf_ln))
}
negative_expoential_hazard_rate <- function(t, acute, h) {
  # Hazard for Negative Exponential distribution (constant hazard = h after acute phase)
  return(h)
}

# Prepare a list to store hazard rate data frames for each NCD
HR_ncds <- list()

# Compute hazard rates for each NCD in the list
for (ncd_name in NCDs) {
  # Retrieve parameter sets for this NCD
  params <- survival_curve_para[[ncd_name]]
  acute_tlb  <- params$treated_lb[1];   para1_tlb  <- params$treated_lb[2];   para2_tlb  <- params$treated_lb[3]
  acute_tub  <- params$treated_ub[1];   para1_tub  <- params$treated_ub[2];   para2_tub  <- params$treated_ub[3]
  acute_utlb <- params$untreated_lb[1]; para1_utlb <- params$untreated_lb[2]; para2_utlb <- params$untreated_lb[3]
  acute_utub <- params$untreated_ub[1]; para1_utub <- params$untreated_ub[2]; para2_utub <- params$untreated_ub[3]
  
  # Initialize a list to collect hazard at each month for this NCD
  HR_t_all <- list()
  
  # Loop over time from t = 0 (acute phase) to t = 499 months (approximately 40 years)
  for (t in 0:499) {
    if (t == 0) {
      # Acute phase (t = 0): calculate immediate mortality fraction (1 - acute survival%)
      HR_t_tlb  <- 1 - acute_tlb/100
      HR_t_tub  <- 1 - acute_tub/100
      HR_t_utlb <- 1 - acute_utlb/100
      HR_t_utub <- 1 - acute_utub/100
    } else {
      # Post-acute hazard (t > 0): use appropriate hazard function for each scenario
      # Determine which hazard function to use (by name) and call it with parameters
      func_name <- ncd_survival_function[[ncd_name]]           # e.g., "log_normal_hazard_rate"
      # Use get() to retrieve the actual function object by name and call it
      HR_t_tlb  <- get(func_name)(t, para1_tlb,  para2_tlb)
      HR_t_tub  <- get(func_name)(t, para1_tub,  para2_tub)
      HR_t_utlb <- get(func_name)(t, para1_utlb, para2_utlb)
      HR_t_utub <- get(func_name)(t, para1_utub, para2_utub)
      # (For DM1, the hazard function returns constant h for t > 0)
    }
    # Append the hazard values for this time point as a numeric vector
    HR_t_all[[length(HR_t_all) + 1]] <- c(HR_t_tlb, HR_t_tub, HR_t_utlb, HR_t_utub)
  }
  
  # Convert the hazard list to a data frame with appropriate column names
  HR_df <- as.data.frame(do.call(rbind, HR_t_all))
  colnames(HR_df) <- c('HR_t_tlb', 'HR_t_tub', 'HR_t_utlb', 'HR_t_utub')
  
  # Store the data frame in the list with the NCD name as key
  HR_ncds[[ncd_name]] <- HR_df
}

# Write the hazard rates to an Excel file with one sheet per NCD
wb <- createWorkbook()
for (ncd_name in names(HR_ncds)) {
  addWorksheet(wb, sheetName = ncd_name)
  writeData(wb, sheet = ncd_name, x = HR_ncds[[ncd_name]], colNames = TRUE, rowNames = FALSE)
}
saveWorkbook(wb, file = "HR_new_logistic_normal.xlsx", overwrite = TRUE)
