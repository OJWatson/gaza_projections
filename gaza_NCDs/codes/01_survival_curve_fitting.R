# 01_survival_curve_fitting.R – Translate Python survival curve fitting to R
setwd("gaza_NCDs")

# Load required libraries
library(readxl)    # for reading Excel files
library(ggplot2)   # for plotting (optional, used for diagnostic plots)
# (No additional libraries needed for curve fitting; using base R nls with port algorithm)

# Define the list of NCDs to process (same as in Python script)
NCDs <- c('CKD', 'IHD', 'HS', 'IS', 'Bcancer', 'Ccancer', 'Lcancer')

# (Optional) Full names and category mappings for NCDs (mirroring Python dictionaries for completeness)
NCD_full_name <- list(
  DM1    = 'Diabetes mellitus type 1 (E10)',
  # DM2 repeated as DM1 in original Python (likely a typo); skipping duplicate key
  CKD    = 'Chronic kidney disease (N18)',
  IHD    = 'Ischaemic heart diseases (I20-25)',
  CDIS   = 'Cerebrovascular diseases including stroke (I63-I66)',
  HD     = 'Hypertensive diseases (I11-12)',
  HS     = 'Haemorrhagic Stroke',
  IS     = 'Ischemic Stroke',
  COPD   = 'Chronic obstructive pulmonary disease (J44)',
  Bcancer= 'Breast cancer',
  Ccancer= 'Colorectal cancer',
  Lcancer= 'Lung cancer'
)
NCD_cata <- list(
  DM1    = 'DM',
  CKD    = 'CKD',
  IHD    = 'CVD',
  CDIS   = 'CVD',
  HD     = 'CVD',
  HS     = 'CVD',
  IS     = 'CVD',
  COPD   = 'COPD',
  Bcancer= 'Cancer',
  Ccancer= 'Cancer',
  Lcancer= 'Cancer'
)

# Initialize structures to record fit goodness and parameters
fit_goodness <- data.frame()   # will hold MSE (mean squared error) for each fit
survival_curve_para <- list()  # nested list to store fitted parameters for each NCD & scenario

# Read survival percentage data from Excel (sheet "Treatment")
# The Excel is expected in the inputs/ subfolder. Adjust path if necessary.
df <- read_excel("inputs/NCD_M2_model_para_1.26.2024.xlsx", sheet = "Treatment")

# Convert the first row (immediate survival rates) from percentage to fraction
# and scale subsequent rows by these fractions to get actual survival probabilities.
df[1, ] <- df[1, ] / 100            # first row now in 0–1 range (acute survival fraction for each scenario)
first_row <- df[1, ]               # store the first row of fractions
other_rows <- df[-1, ]             # all subsequent rows (remaining survival percentages)

# Element-wise multiply each subsequent row by the corresponding acute fraction, then divide by 100
# This yields cumulative survival fraction of the original cohort at each time point.
# (We use sweep to multiply each column of other_rows by the value in first_row for that column.)
other_vals <- sweep(as.data.frame(other_rows), 2, as.numeric(first_row), `*`)
other_vals <- other_vals / 100

# Combine the acute row back with the scaled survival data
curve_vals <- rbind(as.data.frame(first_row), other_vals)

# Add a "Time (Month)" column: 0 for the first (acute) row, then 1..N for subsequent rows.
curve_vals$`Time (Month)` <- 0:(nrow(curve_vals) - 1)

# Reorder columns if necessary so that "Time (Month)" is first (for convenience in plotting/analysis)
curve_vals <- curve_vals[, c(ncol(curve_vals), 1:(ncol(curve_vals)-1))]

# Define survival function models (as in Python, using `acute` from environment)
weibull_survival <- function(t, lambda, kappa) {
  # Weibull survival: S(t) = acute * exp(- (t/lambda)^kappa )
  acute * exp(- (t / lambda) ^ kappa)
}
log_normal_survival <- function(t, mu, sigma) {
  # Log-Normal survival: S(t) = acute * [1 - Φ((ln t - mu)/sigma) ], where Φ is CDF of normal distribution
  acute * (1 - pnorm(log(t), mean = mu, sd = sigma))
}
log_logistic_survival <- function(t, alpha, beta) {
  # Log-Logistic survival: S(t) = acute / [1 + (t/alpha)^beta]
  acute / (1 + (t / alpha) ^ beta)
}
gamma_survival <- function(t, shape, scale) {
  # Gamma (alternate) survival: S(t) = acute * [1 - F_gamma(t; shape, scale) ], using gamma CDF
  acute * (1 - pgamma(t, shape = shape, scale = scale))
}
exponential_survival <- function(t, lambda) {
  # Exponential survival: S(t) = acute * exp(-lambda * t)
  acute * exp(- lambda * t)
}

# Select the distribution for fitting. (Default is log-logistic as per the original script.)
curve_model <- 'log_logistic_survival'
# (To choose a different model, set curve_model to 'weibull_survival', 'log_normal_survival', etc.)

# Set parameter bounds and initial guesses for certain distributions (following Python logic)
param_lower_bounds <- NULL
initial_guess <- NULL
if (curve_model %in% c('weibull_survival', 'log_logistic_survival')) {
  # For Weibull and log-logistic, use small initial guesses and enforce positive bounds
  initial_guess <- list(param1 = 0.01, param2 = 0.01)       # initial guess for two parameters
  param_lower_bounds <- c(param1 = 0.001, param2 = 0.001)   # both parameters > 0
} else {
  # For other models (log-normal, gamma, exponential), we use default or moderate initial guesses
  # Determine number of parameters needed:
  if (curve_model == 'exponential_survival') {
    initial_guess <- list(param1 = 1.0)    # exponential has one parameter (lambda)
    # (lambda > 0 is inherently required, but we won't explicitly set bounds to mimic original behavior)
  } else {
    initial_guess <- list(param1 = 1.0, param2 = 1.0)  # log-normal and gamma have two parameters; start at 1
  }
  # No explicit bounds for these in original script
  param_lower_bounds <- NULL  
}

# Prepare vectors to accumulate fit names and errors
ncd_name_list <- c()   # will store strings like "CKD_treated_lb"
ncd_error_list <- c()  # will store MSE for each corresponding entry

# Loop over each NCD and each treatment scenario to fit survival curves
for (ncd in NCDs) {
  # Initialize an inner list for this NCD to hold scenario parameters
  survival_curve_para[[ncd]] <- list()
  
  for (cond in c('_treated_lb', '_treated_ub', '_untreated_lb', '_untreated_ub')) {
    col_name <- paste0(ncd, cond)  # e.g., "CKD_treated_lb"
    if (!(col_name %in% names(curve_vals))) {
      next  # if this scenario column is not present in data, skip it
    }
    # Extract the time and survival values for this NCD scenario, dropping any missing data
    info <- curve_vals[, c("Time (Month)", col_name)]
    info <- info[!is.na(info[[col_name]]), ]  # drop rows where this scenario has NA
    
    # Separate the time and survival vectors, excluding the first time point (t=0) from the fit
    # The first row corresponds to time 0 (acute phase), which we handle separately.
    months <- info[["Time (Month)"]][-1]           # times from 1 onwards
    survival_rates <- info[[col_name]][-1]         # observed survival fractions from time 1 onwards
    acute <- info[[col_name]][1]                   # initial survival fraction at t=0 (acute phase)
    
    # Choose the model function by name (from those defined above)
    model_fn <- switch(curve_model,
                       weibull_survival    = weibull_survival,
                       log_normal_survival = log_normal_survival,
                       log_logistic_survival = log_logistic_survival,
                       gamma_survival      = gamma_survival,
                       exponential_survival = exponential_survival)
    
    # Prepare data for nls fitting
    fit_data <- data.frame(t = months, y = survival_rates)
    
    # Fit the survival curve to the data using non-linear least squares (nls)
    # We'll use the port algorithm if bounds are specified; otherwise default algorithm.
    if (!is.null(param_lower_bounds)) {
      # Two-parameter model with bounds (for Weibull or log-logistic)
      fit_model <- nls(
        y ~ model_fn(t, param1, param2), data = fit_data,
        start = list(param1 = initial_guess$param1, param2 = initial_guess$param2),
        algorithm = "port", lower = param_lower_bounds
      )
    } else {
      # Model without explicit bounds (log-normal, gamma, or exponential)
      if (length(initial_guess) == 1) {
        # One-parameter model (exponential)
        fit_model <- nls(
          y ~ model_fn(t, param1), data = fit_data,
          start = list(param1 = initial_guess$param1)
        )
      } else {
        # Two-parameter model (log-normal or gamma)
        fit_model <- nls(
          y ~ model_fn(t, param1, param2), data = fit_data,
          start = list(param1 = initial_guess$param1, param2 = initial_guess$param2)
        )
      }
    }
    
    # Extract fitted parameters
    params <- coef(fit_model)
    # Compute fitted values for the observed time points (to calculate error)
    fitted_values <- predict(fit_model, newdata = data.frame(t = months))
    # Calculate mean squared error (MSE) between observed and fitted survival rates
    mse <- mean((survival_rates - fitted_values)^2)
    
    # Store the results.
    # **Note**: We store acute as a percentage (acute*100) in the output to match the original format
    # The output vector is [acute_percent, param1, param2] for each scenario.
    if (length(params) == 1) {
      # For exponential (single parameter), include a placeholder second parameter for consistency
      survival_curve_para[[ncd]][[substring(cond, 2)]] <- c(acute * 100, params[[1]], NA)
    } else {
      survival_curve_para[[ncd]][[substring(cond, 2)]] <- c(acute * 100, params[[1]], params[[2]])
    }
    
    # Record the scenario name and error for the goodness-of-fit table
    ncd_name_list <- c(ncd_name_list, paste0(ncd, cond))
    ncd_error_list <- c(ncd_error_list, mse)
  }
}

# Create the fit goodness data frame with NCD scenario and MSE for the chosen curve model
fit_goodness <- data.frame(
  NCD = ncd_name_list,
  # Use the model name as the column for MSE, e.g., "log_logistic_survival"
  # (Replace any special characters if needed to make a valid name)
  MSE = ncd_error_list
)
colnames(fit_goodness)[2] <- curve_model  # name second column after the model

# Save the fitted parameters and goodness-of-fit metrics to output files
if (!dir.exists("outputs")) dir.create("outputs")
write.csv(fit_goodness, file = file.path("outputs", paste0("fit_goodness_", curve_model, ".csv")), row.names = FALSE)
# Prepare a flat table of survival_curve_para for saving: columns = NCD, scenario, acute_percent, param1, param2
param_records <- data.frame()
for (ncd in names(survival_curve_para)) {
  for (scenario in names(survival_curve_para[[ncd]])) {
    vals <- survival_curve_para[[ncd]][[scenario]]
    param_records <- rbind(param_records, data.frame(
      NCD = ncd,
      scenario = scenario,
      acute_percent = vals[1],
      param1 = vals[2],
      param2 = vals[3]
    ))
  }
}
write.csv(param_records, file = file.path("outputs", paste0("survival_curve_parameters_", curve_model, ".csv")), row.names = FALSE)

