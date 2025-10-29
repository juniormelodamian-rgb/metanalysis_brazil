# ============================================================
# Meta-analysis (lnRR, multilevel) — Amazon 0–30 cm 
# ============================================================

# ---------- Packages ----------
required_pkgs <- c("readxl","writexl","metafor","dplyr","ggplot2")
new_pkgs <- required_pkgs[!(required_pkgs %in% installed.packages()[,"Package"])]
if (length(new_pkgs)) install.packages(new_pkgs, dependencies = TRUE)

library(readxl); library(writexl)
library(metafor); library(dplyr); library(ggplot2)

# ---------- Config (edit here if needed) ----------
# Base directory: use here::here() if available, else getwd()
if (requireNamespace("here", quietly = TRUE)) {
  base_dir <- here::here()
} else {
  base_dir <- getwd()
}

# Input locations (relative to repo root). Adjust if your files are elsewhere.
data_dir    <- file.path(base_dir, "data")
results_dir <- file.path(base_dir, "results")

# Input file & sheet
infile  <- file.path(data_dir, "Amazon_Int_0-30.xlsx")
insheet <- "1_2"

# Output files
out_main_xlsx <- file.path(results_dir, "Amazon_1_2f.xlsx")
out_bias_xlsx <- file.path(results_dir, "Atlantic_Forest_trimfill_failsafe_summary.xlsx")  # mantém seu nome original

# Ensure output dir exists
if (!dir.exists(results_dir)) dir.create(results_dir, recursive = TRUE)

# ---------- Helpers ----------
parse_depth <- function(depth_str){
  s <- gsub("cm|\\s", "", depth_str)
  s <- gsub("–|—|−", "-", s)
  parts <- unlist(strsplit(s, "-"))
  if (length(parts) < 2) return(c(NA, NA))
  as.numeric(parts[1:2])
}

calc_lnRR <- function(m_t, sd_t, n_t, m_c, sd_c, n_c, zero_adj = 1e-6){
  m_t <- as.numeric(m_t); m_c <- as.numeric(m_c)
  sd_t <- as.numeric(sd_t); sd_c <- as.numeric(sd_c)
  n_t <- as.numeric(n_t);   n_c <- as.numeric(n_c)
  
  m_t[m_t == 0] <- zero_adj
  m_c[m_c == 0] <- zero_adj
  
  v  <- (sd_t^2)/(n_t * m_t^2) + (sd_c^2)/(n_c * m_c^2)
  yi <- log(m_t / m_c)
  list(yi = yi, vi = v)
}

prepare_es <- function(dat,
                       col_m_t = "treatment_mean",
                       col_sd_t = "treatment_sd",
                       col_n_t = "treatment_n",
                       col_m_c = "control_mean",
                       col_sd_c = "control_sd",
                       col_n_c = "control_n",
                       col_depth = "depth",
                       col_study = "study_id"){
  needed  <- c(col_m_t, col_sd_t, col_n_t, col_m_c, col_sd_c, col_n_c, col_depth, col_study)
  missing <- setdiff(needed, names(dat))
  if (length(missing) > 0){
    stop("The following columns were not found in the sheet: ", paste(missing, collapse=", "))
  }
  
  dat2 <- dat %>%
    mutate(
      depth_raw = !!rlang::sym(col_depth),
      study_id  = !!rlang::sym(col_study)
    ) %>%
    rowwise() %>%
    mutate(
      depth_start = parse_depth(depth_raw)[1],
      depth_end   = parse_depth(depth_raw)[2]
    ) %>% ungroup()
  
  ln <- calc_lnRR(
    m_t = dat2[[col_m_t]],
    sd_t = dat2[[col_sd_t]],
    n_t = dat2[[col_n_t]],
    m_c = dat2[[col_m_c]],
    sd_c = dat2[[col_sd_c]],
    n_c = dat2[[col_n_c]]
  )
  
  dat2$yi <- ln$yi
  dat2$vi <- ln$vi
  dat2
}

hetero_info <- function(model, vi = NULL, tau2_null = NULL){
  tau2 <- if (!is.null(model$tau2)) model$tau2 else NA
  
  if (!is.null(vi) & !is.na(tau2)){
    I2 <- 100 * tau2 / (tau2 + mean(vi, na.rm = TRUE))
    H2 <- 1 / (1 - I2/100)
  } else {
    I2 <- NA; H2 <- NA
  }
  
  if (!is.null(tau2_null) & !is.na(tau2)){
    R2 <- (tau2_null - tau2) / tau2_null * 100
  } else {
    R2 <- NA
  }
  
  data.frame(Tau2 = tau2, I2 = I2, H2 = H2, R2 = R2)
}

calc_perc_change <- function(model){
  b     <- as.numeric(coef(model)[1])
  se_b  <- as.numeric(model$se[1])
  ci_lb <- b - 1.96 * se_b
  ci_ub <- b + 1.96 * se_b
  
  perc_change <- (exp(b)     - 1) * 100
  ci_low      <- (exp(ci_lb) - 1) * 100
  ci_high     <- (exp(ci_ub) - 1) * 100
  
  list(lnRR = b, lnRR_CI_lb = ci_lb, lnRR_CI_ub = ci_ub,
       PercChange = perc_change, PercChange_CI_lb = ci_low, PercChange_CI_ub = ci_high)
}

# ---------- 1) Read worksheet ----------
if (!file.exists(infile)) {
  stop("Input file not found: ", infile,
       "\nPlace your Excel file at: ", infile, "\n(or edit 'infile' at the top).")
}
dat <- read_excel(infile, sheet = insheet)

# ---------- 2) Compute lnRR ----------
es <- prepare_es(dat,
                 col_m_t = "treatment_mean",
                 col_sd_t = "treatment_sd",
                 col_n_t = "treatment_n",
                 col_m_c = "control_mean",
                 col_sd_c = "control_sd",
                 col_n_c = "control_n",
                 col_depth = "depth",
                 col_study = "study_id")

# ---------- 2b) Standardize moderators ----------
# ensure numeric scale output
es <- es %>%
  mutate(
    Temperature_z = as.numeric(scale(Temperature)),
    Rainfall_z    = as.numeric(scale(Rainfall)),
    Lat_z         = as.numeric(scale(Lat))
  )

# factors for random effects
es$study_id  <- as.factor(es$study_id)
es$depth_raw <- as.factor(es$depth_raw)

# ---------- 3) Null model ----------
res_null <- rma.mv(yi, vi,
                   random = ~ 1 | study_id/depth_raw,
                   data = es,
                   method = "REML")

# ---------- 4) Multilevel model (no moderators) ----------
res_mv <- rma.mv(yi, vi,
                 random = ~ 1 | study_id/depth_raw,
                 data = es,
                 method = "REML")

# ---------- 5) Model with moderators ----------
res_mv_mod <- rma.mv(yi, vi,
                     mods = ~ Temperature_z + Rainfall_z + Lat_z,
                     random = ~ 1 | study_id/depth_raw,
                     data = es,
                     method = "REML")

# ---------- 6) Percent change ----------
res_null_perc   <- calc_perc_change(res_null)
res_mv_perc     <- calc_perc_change(res_mv)
res_mv_mod_perc <- calc_perc_change(res_mv_mod)

# ---------- 7) Heterogeneity ----------
hetero_null <- hetero_info(res_null, vi = es$vi)
hetero_mv   <- hetero_info(res_mv,   vi = es$vi)
hetero_mod  <- hetero_info(res_mv_mod, vi = es$vi, tau2_null = res_null$tau2)

# ---------- 8) Export main results ----------
write_xlsx(list(
  effect_sizes     = es,
  results_null     = data.frame(res_null_perc, hetero_null),
  results_mv       = data.frame(res_mv_perc, hetero_mv),
  results_mod      = data.frame(res_mv_mod_perc, hetero_mod)
), path = out_main_xlsx)

# ---------- 9) Forest plots (saved to results/) ----------
tiff(file.path(results_dir, "forest_null.tif"), width = 2000, height = 1500, res = 300)
forest(res_null, slab = paste(es$study_id, es$depth_raw), xlab = "lnRR (Null Model)")
dev.off()

tiff(file.path(results_dir, "forest_mv.tif"), width = 2000, height = 1500, res = 300)
forest(res_mv, slab = paste(es$study_id, es$depth_raw), xlab = "lnRR (Multilevel, no moderators)")
dev.off()

tiff(file.path(results_dir, "forest_mv_mod.tif"), width = 2000, height = 1500, res = 300)
forest(res_mv_mod, slab = paste(es$study_id, es$depth_raw), xlab = "lnRR (Multilevel with moderators)")
dev.off()

# ---------- 10) Egger test ----------
res_egger <- rma(yi, vi, data = es, method = "REML")
egger_out <- capture.output(print(regtest(res_egger, model = "lm")))
writeLines(egger_out, con = file.path(results_dir, "egger_test.txt"))

# ---------- 11) Funnel plot ----------
tiff(file.path(results_dir, "funnel_res_egger.tif"), width = 2000, height = 1500, res = 300)
funnel(res_egger, xlab = "lnRR", ylab = "SE(lnRR)", main = "Funnel plot")
dev.off()

# ---------- 12) Trim-and-Fill & Fail-safe N ----------
res_simple <- rma(yi, vi, data = es, method = "REML")

# Trim-and-Fill
res_tf <- trimfill(res_simple)
tf_summary <- data.frame(
  k_original        = length(res_simple$yi),
  k_adjusted        = length(res_tf$yi),
  estimate_original = as.numeric(coef(res_simple)),
  estimate_adjusted = as.numeric(coef(res_tf))
)
tiff(file.path(results_dir, "funnel_trimfill.tif"), width = 2000, height = 1500, res = 300)
funnel(res_tf, xlab = "lnRR", ylab = "SE(lnRR)", main = "Funnel plot (Trim-and-Fill)")
dev.off()

# Fail-safe N (with guard)
fsn_res <- try(fsn(res_simple, type = "Rosenthal"), silent = TRUE)
if (inherits(fsn_res, "try-error") || is.null(fsn_res$fsn)) {
  fsn_summary <- data.frame(
    fail_safe_N   = NA,
    z_crit        = NA,
    p_crit        = NA,
    est_mean      = NA,
    est_mean_fill = NA
  )
} else {
  fsn_summary <- data.frame(
    fail_safe_N   = fsn_res$fsn,
    z_crit        = fsn_res$z.crit,
    p_crit        = fsn_res$p.crit,
    est_mean      = fsn_res$mean,
    est_mean_fill = fsn_res$mean.fill
  )
}

# Export bias summaries
write_xlsx(list(
  trimfill_summary  = tf_summary,
  fail_safe_summary = fsn_summary
), path = out_bias_xlsx)

message("Done. Results written to: ", results_dir)
